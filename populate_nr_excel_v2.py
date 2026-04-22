import argparse
import csv
import glob
import os
import shutil
import time
from datetime import datetime
from pathlib import Path

import win32com.client as win32

SPANISH_MONTHS = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}


def find_latest_csv(mode: str, target_date: str) -> str:
    if mode == "conversation":
        pattern = os.path.join(
            "outputConversationStart",
            f"conversation_start_por_hora_del_dia_{target_date}_*.csv"
        )
    elif mode == "event":
        pattern = os.path.join(
            "outputEvent",
            f"evento_tnotresponding_por_hora_del_dia_{target_date}_*.csv"
        )
    else:
        raise ValueError("mode debe ser 'conversation' o 'event'")

    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError(f"No se encontraron archivos con patrón: {pattern}")

    return max(files, key=os.path.getmtime)


def read_hourly_csv(csv_path: str) -> dict:
    hourly = {}

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(2048)
        f.seek(0)

        delimiter = ";" if ";" in sample else ","
        reader = csv.DictReader(f, delimiter=delimiter)

        for row in reader:
            hora = (row.get("hora") or "").strip()
            cantidad = row.get("cantidad")

            if not hora:
                continue

            try:
                hourly[hora] = int(cantidad)
            except (TypeError, ValueError):
                hourly[hora] = 0

    full = {}
    for h in range(24):
        key = f"{h:02d}:00"
        full[key] = hourly.get(key, 0)

    return full


def build_rows(target_date: str, hourly_data: dict):
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    year = dt.year
    month = SPANISH_MONTHS[dt.month]
    day = dt.day

    rows = []
    for h in range(24):
        hour_str = f"{h}:00"
        qty = hourly_data.get(f"{h:02d}:00", 0)

        rows.append([
            target_date,  # FECHA
            year,         # AÑO
            month,        # MES
            day,          # DIA
            hour_str,     # HORA
            qty           # Cantidad
        ])

    return rows


def write_rows_excel(ws, rows, start_row: int, start_col: int = 1):
    if not rows:
        return

    nrows = len(rows)
    ncols = len(rows[0])

    end_row = start_row + nrows - 1
    end_col = start_col + ncols - 1

    target = ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_col))
    target.Value = rows


def sheet_exists(workbook, sheet_name: str) -> bool:
    for ws in workbook.Worksheets:
        if ws.Name == sheet_name:
            return True
    return False


def get_last_used_row(ws, key_col: int = 1, min_row: int = 6) -> int:
    last_row = ws.Cells(ws.Rows.Count, key_col).End(-4162).Row  # xlUp
    return max(last_row, min_row - 1)


def get_existing_rows_for_date(ws, target_date: str, start_row: int = 6, date_col: int = 1):
    last_row = get_last_used_row(ws, key_col=date_col, min_row=start_row)
    matches = []

    for row in range(start_row, last_row + 1):
        value = ws.Cells(row, date_col).Value
        if value is None:
            continue

        value_str = str(value).strip()

        if value_str == target_date:
            matches.append(row)
            continue

        try:
            dt = datetime.strptime(value_str.split(" ")[0], "%d/%m/%Y")
            if dt.strftime("%Y-%m-%d") == target_date:
                matches.append(row)
                continue
        except Exception:
            pass

        try:
            dt = datetime.strptime(value_str.split(" ")[0], "%Y-%m-%d")
            if dt.strftime("%Y-%m-%d") == target_date:
                matches.append(row)
        except Exception:
            pass

    return matches


def clear_rows(ws, start_row: int, end_row: int, start_col: int = 1, end_col: int = 6):
    rng = ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_col))
    rng.ClearContents()


def compact_table(ws, start_row: int = 6, start_col: int = 1, end_col: int = 6):
    last_row = get_last_used_row(ws, key_col=start_col, min_row=start_row)
    data = []

    for row in range(start_row, last_row + 1):
        row_values = []
        has_data = False
        for col in range(start_col, end_col + 1):
            val = ws.Cells(row, col).Value
            row_values.append(val)
            if val not in (None, ""):
                has_data = True
        if has_data:
            data.append(row_values)

    clear_rows(ws, start_row, max(last_row, start_row + 200), start_col, end_col)

    if data:
        target = ws.Range(
            ws.Cells(start_row, start_col),
            ws.Cells(start_row + len(data) - 1, end_col)
        )
        target.Value = data


def do_refresh(excel, wb, wait_seconds: int = 5):
    wb.RefreshAll()
    time.sleep(wait_seconds)
    excel.CalculateUntilAsyncQueriesDone()
    excel.CalculateFull()


def main():
    parser = argparse.ArgumentParser(
        description="Poblar Excel de Not Responding desde CSV por hora del día usando Excel COM."
    )
    parser.add_argument("--workbook", required=True, help="Ruta del Excel plantilla")
    parser.add_argument("--date", required=True, help="Fecha objetivo YYYY-MM-DD")
    parser.add_argument(
        "--mode",
        choices=["conversation", "event"],
        default="conversation",
        help="Fuente de datos: conversationStart o evento"
    )
    parser.add_argument(
        "--csv",
        help="Ruta explícita al CSV por hora del día. Si no se indica, toma el más reciente de la fecha dada."
    )
    parser.add_argument(
        "--output",
        default="Reporte_Final.xlsx",
        help="Archivo de salida"
    )
    parser.add_argument(
        "--start-row",
        type=int,
        default=6,
        help="Fila inicial de la tabla"
    )
    parser.add_argument(
        "--sheet-hour",
        default="RESUMEN POR RANGO DE HORA",
        help="Hoja a poblar"
    )
    parser.add_argument(
        "--replace-date",
        action="store_true",
        help="Si la fecha ya existe, la reemplaza en vez de omitirla."
    )
    parser.add_argument(
        "--refresh",
        action="store_true",
        help="Ejecuta 'Actualizar todo' y recalcula antes de guardar."
    )
    parser.add_argument(
        "--refresh-wait",
        type=int,
        default=5,
        help="Segundos de espera después de RefreshAll(). Default: 5"
    )
    args = parser.parse_args()

    csv_path = args.csv if args.csv else find_latest_csv(args.mode, args.date)
    print(f"CSV usado: {csv_path}")

    hourly_data = read_hourly_csv(csv_path)
    rows = build_rows(args.date, hourly_data)

    if not os.path.exists(args.output):
        shutil.copy2(args.workbook, args.output)
        print(f"Se creó archivo base desde plantilla: {args.output}")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = None
    try:
        output_abs = str(Path(args.output).resolve())
        wb = excel.Workbooks.Open(output_abs)

        if not sheet_exists(wb, args.sheet_hour):
            raise ValueError(f"No existe la hoja: {args.sheet_hour}")

        ws_hour = wb.Worksheets(args.sheet_hour)

        existing_rows = get_existing_rows_for_date(
            ws_hour,
            args.date,
            start_row=args.start_row,
            date_col=1
        )

        if existing_rows:
            if args.replace_date:
                print(f"La fecha {args.date} ya existe. Se reemplazará.")
                clear_rows(
                    ws_hour,
                    min(existing_rows),
                    max(existing_rows),
                    start_col=1,
                    end_col=6
                )
                compact_table(ws_hour, start_row=args.start_row, start_col=1, end_col=6)
            else:
                print(f"La fecha {args.date} ya existe en la hoja. No se agregará nuevamente.")
                if args.refresh:
                    print("Ejecutando refresh final...")
                    do_refresh(excel, wb, wait_seconds=args.refresh_wait)
                wb.Save()
                print(f"Excel generado: {args.output}")
                return

        last_row = get_last_used_row(ws_hour, key_col=1, min_row=args.start_row)
        insert_row = max(last_row + 1, args.start_row)

        write_rows_excel(
            ws_hour,
            rows,
            start_row=insert_row,
            start_col=1
        )

        if args.refresh:
            print("Ejecutando refresh final...")
            do_refresh(excel, wb, wait_seconds=args.refresh_wait)

        wb.Save()
        print(f"Excel generado: {args.output}")
        print(f"Registros insertados desde fila {insert_row} hasta {insert_row + len(rows) - 1}")

    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    main()