import os
import glob
import csv
import argparse
import subprocess
from datetime import datetime, timedelta, timezone

PERU_TZ = timezone(timedelta(hours=-5))


def parse_args():
    parser = argparse.ArgumentParser(
        description="Procesa Not Responding por rango de fechas y opcionalmente hace un único refresh final."
    )
    parser.add_argument(
        "--start-date",
        required=True,
        help="Fecha inicial en formato YYYY-MM-DD"
    )
    parser.add_argument(
        "--end-date",
        default=datetime.now(PERU_TZ).strftime("%Y-%m-%d"),
        help="Fecha final en formato YYYY-MM-DD. Default: hoy en Perú"
    )
    parser.add_argument(
        "--mode",
        choices=["conversation", "event"],
        default="conversation",
        help="Fuente de datos para el populate. Default: conversation"
    )
    parser.add_argument(
        "--generator-script",
        default="notRespondingV5.py",
        help="Script generador de la base diaria"
    )
    parser.add_argument(
        "--populate-script",
        default="populate_nr_excel_v2.py",
        help="Script que puebla el Excel"
    )
    parser.add_argument(
        "--workbook-template",
        default="Not-Responding-Report.xlsx",
        help="Plantilla base de Excel"
    )
    parser.add_argument(
        "--output-workbook",
        default="Reporte-Not-Responding.xlsx",
        help="Archivo Excel acumulado de salida"
    )
    parser.add_argument(
        "--log-file",
        default="run_nr_range_log.csv",
        help="CSV de log"
    )
    parser.add_argument(
        "--refresh-wait",
        type=int,
        default=5,
        help="Segundos de espera para el refresh final"
    )
    parser.add_argument(
        "--final-refresh",
        action="store_true",
        help="Ejecuta un único refresh final al terminar el rango"
    )
    return parser.parse_args()


def daterange(start_date, end_date):
    current = start_date
    while current <= end_date:
        yield current
        current += timedelta(days=1)


def run_command(cmd):
    result = subprocess.run(cmd, capture_output=True, text=True)
    return {
        "returncode": result.returncode,
        "stdout": result.stdout,
        "stderr": result.stderr,
    }


def csv_exists_for_date(date_str, mode):
    if mode == "conversation":
        pattern = os.path.join(
            "outputConversationStart",
            f"conversation_start_por_hora_del_dia_{date_str}_*.csv"
        )
    else:
        pattern = os.path.join(
            "outputEvent",
            f"evento_tnotresponding_por_hora_del_dia_{date_str}_*.csv"
        )
    return len(glob.glob(pattern)) > 0


def append_log(log_file, row):
    file_exists = os.path.exists(log_file)
    with open(log_file, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "timestamp",
                "date",
                "generator_status",
                "populate_status",
                "message"
            ]
        )
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)


def main():
    args = parse_args()

    start_dt = datetime.strptime(args.start_date, "%Y-%m-%d").date()
    end_dt = datetime.strptime(args.end_date, "%Y-%m-%d").date()

    if start_dt > end_dt:
        raise ValueError("start-date no puede ser mayor que end-date")

    print(f"Procesando desde {start_dt} hasta {end_dt}")
    ok_count = 0
    error_count = 0
    processed_dates = []

    for day in daterange(start_dt, end_dt):
        date_str = day.strftime("%Y-%m-%d")
        print(f"\n===== Fecha: {date_str} =====")

        generator_status = "SKIPPED"
        populate_status = "PENDING"
        message = ""

        try:
            # 1) Generar CSVs si aún no existen
            if not csv_exists_for_date(date_str, args.mode):
                print(f"Generando base para {date_str}...")
                gen_result = run_command(["python", args.generator_script, date_str])

                if gen_result["returncode"] != 0:
                    generator_status = "ERROR"
                    populate_status = "SKIPPED"
                    message = (gen_result["stderr"] or gen_result["stdout"] or "Error desconocido").strip()[:1000]
                    error_count += 1

                    append_log(args.log_file, {
                        "timestamp": datetime.now().isoformat(),
                        "date": date_str,
                        "generator_status": generator_status,
                        "populate_status": populate_status,
                        "message": message
                    })
                    print(f"[ERROR] Generador falló para {date_str}")
                    continue

                generator_status = "OK"
            else:
                print(f"CSV ya existe para {date_str}, se omite generación.")
                generator_status = "SKIPPED"

            # 2) Poblar Excel acumulado SIN refresh
            print(f"Poblando Excel para {date_str} sin refresh...")
            pop_result = run_command([
                "python", args.populate_script,
                "--workbook", args.workbook_template,
                "--date", date_str,
                "--mode", args.mode,
                "--output", args.output_workbook
            ])

            if pop_result["returncode"] != 0:
                populate_status = "ERROR"
                message = (pop_result["stderr"] or pop_result["stdout"] or "Error desconocido").strip()[:1000]
                error_count += 1
                print(f"[ERROR] Populate falló para {date_str}")
            else:
                populate_status = "OK"
                ok_count += 1
                processed_dates.append(date_str)
                print(f"[OK] Fecha procesada: {date_str}")

            append_log(args.log_file, {
                "timestamp": datetime.now().isoformat(),
                "date": date_str,
                "generator_status": generator_status,
                "populate_status": populate_status,
                "message": message
            })

        except Exception as e:
            error_count += 1
            append_log(args.log_file, {
                "timestamp": datetime.now().isoformat(),
                "date": date_str,
                "generator_status": generator_status,
                "populate_status": "ERROR",
                "message": str(e)[:1000]
            })
            print(f"[ERROR] Excepción en {date_str}: {e}")

    # 3) Refresh final único opcional
    if processed_dates and args.final_refresh:
        last_date = processed_dates[-1]
        print(f"\n===== REFRESH FINAL ÚNICO ({last_date}) =====")

        refresh_result = run_command([
            "python", args.populate_script,
            "--workbook", args.workbook_template,
            "--date", last_date,
            "--mode", args.mode,
            "--output", args.output_workbook,
            "--refresh",
            "--refresh-wait", str(args.refresh_wait)
        ])

        if refresh_result["returncode"] != 0:
            error_count += 1
            append_log(args.log_file, {
                "timestamp": datetime.now().isoformat(),
                "date": last_date,
                "generator_status": "N/A",
                "populate_status": "REFRESH_ERROR",
                "message": (refresh_result["stderr"] or refresh_result["stdout"] or "Error desconocido").strip()[:1000]
            })
            print("[ERROR] Falló el refresh final")
        else:
            append_log(args.log_file, {
                "timestamp": datetime.now().isoformat(),
                "date": last_date,
                "generator_status": "N/A",
                "populate_status": "REFRESH_OK",
                "message": "Refresh final ejecutado correctamente"
            })
            print("[OK] Refresh final ejecutado correctamente")

    print("\n===== RESUMEN FINAL =====")
    print(f"Fechas OK: {ok_count}")
    print(f"Fechas con error: {error_count}")
    print(f"Log generado: {args.log_file}")


if __name__ == "__main__":
    main()