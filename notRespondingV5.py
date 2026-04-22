import os
import math
import csv
import argparse
import requests
from datetime import datetime, timezone, timedelta
from collections import defaultdict

import matplotlib.pyplot as plt

GENESYS_REGION = os.getenv("GENESYS_REGION", "mypurecloud.com")
CLIENT_ID = os.getenv("GENESYS_CLIENT_ID")
CLIENT_SECRET = os.getenv("GENESYS_CLIENT_SECRET")

LOGIN_URL = f"https://login.{GENESYS_REGION}"
API_URL = f"https://api.{GENESYS_REGION}"

PAGE_SIZE = 50
PERU_TZ = timezone(timedelta(hours=-5))

DIR_CONVERSATION_START = "outputConversationStart"
DIR_EVENT = "outputEvent"
DIR_DEBUG = "outputDebug"
DIR_CHARTS = "outputCharts"


def ensure_dirs():
    os.makedirs(DIR_CONVERSATION_START, exist_ok=True)
    os.makedirs(DIR_EVENT, exist_ok=True)
    os.makedirs(DIR_DEBUG, exist_ok=True)
    os.makedirs(DIR_CHARTS, exist_ok=True)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Genera la base de Not Responding desde Genesys Cloud."
    )
    parser.add_argument(
        "date",
        nargs="?",
        default=datetime.now(PERU_TZ).strftime("%Y-%m-%d"),
        help="Fecha objetivo en formato YYYY-MM-DD. Si no se envía, usa el día actual de Perú."
    )
    return parser.parse_args()


def build_interval_for_peru_day(date_str: str):
    local_start = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=PERU_TZ)
    local_end = local_start + timedelta(days=1)

    utc_start = local_start.astimezone(timezone.utc)
    utc_end = local_end.astimezone(timezone.utc)

    start_str = utc_start.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    end_str = utc_end.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    return f"{start_str}/{end_str}"


def get_access_token():
    resp = requests.post(
        f"{LOGIN_URL}/oauth/token",
        data={"grant_type": "client_credentials"},
        auth=(CLIENT_ID, CLIENT_SECRET),
        timeout=30
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def build_body(interval: str, page_number: int):
    return {
        "order": "desc",
        "orderBy": "conversationStart",
        "paging": {
            "pageSize": PAGE_SIZE,
            "pageNumber": page_number
        },
        "interval": interval,
        "segmentFilters": [
            {
                "type": "or",
                "predicates": [
                    {"dimension": "direction", "value": "inbound"},
                    {"dimension": "direction", "value": "outbound"}
                ]
            }
        ],
        "conversationFilters": [
            {
                "type": "and",
                "predicates": [
                    {
                        "type": "metric",
                        "metric": "tNotResponding",
                        "operator": "exists"
                    }
                ]
            }
        ],
        "evaluationFilters": [],
        "surveyFilters": []
    }


def query_page(token, interval: str, page_number: int):
    resp = requests.post(
        f"{API_URL}/api/v2/analytics/conversations/details/query",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        json=build_body(interval, page_number),
        timeout=60
    )
    resp.raise_for_status()
    return resp.json()


def parse_iso_z(ts: str):
    return datetime.fromisoformat(ts.replace("Z", "+00:00"))


def fetch_all_conversations(interval: str):
    token = get_access_token()

    first = query_page(token, interval, 1)
    total_hits = first.get("totalHits", 0)
    conversations = first.get("conversations", [])

    total_pages = math.ceil(total_hits / PAGE_SIZE) if total_hits else 0

    print(f"Intervalo consultado: {interval}")
    print(f"totalHits API: {total_hits}")
    print(f"Página 1: {len(conversations)}")
    print(f"Total páginas esperadas: {total_pages}")

    for page in range(2, total_pages + 1):
        data = query_page(token, interval, page)
        batch = data.get("conversations", [])
        print(f"Página {page}: {len(batch)}")
        conversations.extend(batch)

    print(f"Total descargadas: {len(conversations)}")
    return conversations


def count_by_conversation_start_hour(conversations):
    counts = defaultdict(int)

    for conv in conversations:
        ts = conv.get("conversationStart")
        if not ts:
            continue

        dt_peru = parse_iso_z(ts).astimezone(PERU_TZ)
        bucket = dt_peru.strftime("%Y-%m-%d %H:00")
        counts[bucket] += 1

    return dict(sorted(counts.items()))


def extract_all_tnotresponding_metrics(conv):
    rows = []

    for participant in conv.get("participants", []):
        purpose = participant.get("purpose")
        user_id = participant.get("userId")
        participant_id = participant.get("participantId")

        for session in participant.get("sessions", []):
            session_id = session.get("sessionId")

            for metric in session.get("metrics", []):
                if metric.get("name") == "tNotResponding":
                    emit_date = metric.get("emitDate")
                    value = metric.get("value")
                    rows.append({
                        "conversationId": conv.get("conversationId"),
                        "participantId": participant_id,
                        "purpose": purpose,
                        "userId": user_id,
                        "sessionId": session_id,
                        "emitDate": emit_date,
                        "value": value
                    })

    return rows


def count_by_tnotresponding_emitdate(conversations, interval: str):
    counts = defaultdict(int)
    debug_rows = []

    interval_start, interval_end = interval.split("/")
    interval_start_dt = parse_iso_z(interval_start)
    interval_end_dt = parse_iso_z(interval_end)

    found_total = 0

    for conv in conversations:
        metrics_found = extract_all_tnotresponding_metrics(conv)

        if not metrics_found:
            debug_rows.append({
                "conversationId": conv.get("conversationId"),
                "issue": "MISSING_TNOTRESPONDING_METRIC",
                "conversationStart": conv.get("conversationStart"),
                "emitDate": "",
                "purpose": "",
                "userId": "",
                "participantId": "",
                "sessionId": "",
                "value": ""
            })
            continue

        in_range = []
        out_of_range = []

        for row in metrics_found:
            dt = parse_iso_z(row["emitDate"])
            if interval_start_dt <= dt < interval_end_dt:
                in_range.append((dt, row))
            else:
                out_of_range.append((dt, row))

        if in_range:
            dt, row = sorted(in_range, key=lambda x: x[0])[0]
            bucket = dt.astimezone(PERU_TZ).strftime("%Y-%m-%d %H:00")
            counts[bucket] += 1
            found_total += 1
        else:
            for dt, row in sorted(out_of_range, key=lambda x: x[0]):
                debug_rows.append({
                    "conversationId": row["conversationId"],
                    "issue": "TNOTRESPONDING_OUT_OF_INTERVAL",
                    "conversationStart": conv.get("conversationStart"),
                    "emitDate": row["emitDate"],
                    "purpose": row["purpose"],
                    "userId": row["userId"],
                    "participantId": row["participantId"],
                    "sessionId": row["sessionId"],
                    "value": row["value"]
                })

    return dict(sorted(counts.items())), debug_rows, found_total


def collapse_to_hour_of_day(hourly):
    by_hour = defaultdict(int)

    for dt_hour, count in hourly.items():
        _, hour = dt_hour.split(" ")
        by_hour[hour] += count

    ordered = {}
    for h in range(24):
        key = f"{h:02d}:00"
        ordered[key] = by_hour.get(key, 0)

    return ordered


def export_csv(hourly, filename):
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["fecha", "hora", "cantidad"])

        for hour, count in hourly.items():
            fecha, hora_txt = hour.split(" ")
            writer.writerow([fecha, hora_txt, count])

    print(f"CSV generado: {filename}")


def export_hour_of_day_csv(hourly_by_day, filename):
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["hora", "cantidad"])

        for hour, count in hourly_by_day.items():
            writer.writerow([hour, count])

    print(f"CSV generado: {filename}")


def export_debug_csv(rows, filename):
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            delimiter=";",
            fieldnames=[
                "conversationId",
                "issue",
                "conversationStart",
                "emitDate",
                "purpose",
                "userId",
                "participantId",
                "sessionId",
                "value"
            ]
        )
        writer.writeheader()
        writer.writerows(rows)

    print(f"CSV generado: {filename}")


def generate_line_chart(hourly_by_day, title, output_png):
    hours = list(hourly_by_day.keys())
    values = list(hourly_by_day.values())
    total = sum(values)

    plt.figure(figsize=(14, 7))
    plt.plot(hours, values, marker="o")

    plt.title(title)
    plt.xlabel("Hora del día")
    plt.ylabel("Cantidad")
    plt.xticks(rotation=45)
    plt.grid(True, alpha=0.3)

    for x, y in zip(hours, values):
        if y != 0:
            plt.annotate(
                str(y),
                (x, y),
                textcoords="offset points",
                xytext=(0, 8),
                ha="center"
            )

    plt.text(
        0.99,
        0.98,
        f"Total: {total}",
        transform=plt.gca().transAxes,
        ha="right",
        va="top",
        fontsize=11,
        bbox=dict(boxstyle="round,pad=0.3", alpha=0.15)
    )

    plt.tight_layout()
    plt.savefig(output_png, dpi=150)
    plt.close()

    print(f"Gráfica generada: {output_png}")


def main():
    if not CLIENT_ID or not CLIENT_SECRET:
        raise ValueError("Faltan GENESYS_CLIENT_ID o GENESYS_CLIENT_SECRET")

    ensure_dirs()

    args = parse_args()
    target_date = args.date
    interval = build_interval_for_peru_day(target_date)
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")

    print(f"Fecha objetivo Perú: {target_date}")

    conversations = fetch_all_conversations(interval)

    by_start = count_by_conversation_start_hour(conversations)
    by_event, debug_rows, found_total = count_by_tnotresponding_emitdate(conversations, interval)

    print("\nResumen por hora usando conversationStart:")
    total_start = 0
    for hour, count in by_start.items():
        print(f"{hour} -> {count}")
        total_start += count
    print(f"Total por conversationStart: {total_start}")

    print("\nResumen por hora usando emitDate de tNotResponding:")
    total_event = 0
    for hour, count in by_event.items():
        print(f"{hour} -> {count}")
        total_event += count
    print(f"Total por tNotResponding emitDate: {total_event}")

    print(f"\nConversaciones con metric válido dentro del intervalo: {found_total}")
    print(f"Conversaciones con inconsistencias: {len(debug_rows)}")

    by_start_day = collapse_to_hour_of_day(by_start)
    by_event_day = collapse_to_hour_of_day(by_event)

    export_csv(
        by_start,
        os.path.join(DIR_CONVERSATION_START, f"conversation_start_por_fecha_hora_{target_date}_{run_id}.csv")
    )
    export_hour_of_day_csv(
        by_start_day,
        os.path.join(DIR_CONVERSATION_START, f"conversation_start_por_hora_del_dia_{target_date}_{run_id}.csv")
    )

    export_csv(
        by_event,
        os.path.join(DIR_EVENT, f"evento_tnotresponding_por_fecha_hora_{target_date}_{run_id}.csv")
    )
    export_hour_of_day_csv(
        by_event_day,
        os.path.join(DIR_EVENT, f"evento_tnotresponding_por_hora_del_dia_{target_date}_{run_id}.csv")
    )

    export_debug_csv(
        debug_rows,
        os.path.join(DIR_DEBUG, f"inconsistencias_tnotresponding_{target_date}_{run_id}.csv")
    )

    generate_line_chart(
        by_start_day,
        f"Not Responding por hora del día - conversationStart - {target_date}",
        os.path.join(DIR_CHARTS, f"grafica_conversation_start_{target_date}_{run_id}.png")
    )

    generate_line_chart(
        by_event_day,
        f"Not Responding por hora del día - evento tNotResponding - {target_date}",
        os.path.join(DIR_CHARTS, f"grafica_evento_tnotresponding_{target_date}_{run_id}.png")
    )


if __name__ == "__main__":
    main()