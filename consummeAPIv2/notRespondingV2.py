import os
import math
import csv
import requests
from datetime import datetime, timezone, timedelta
from collections import defaultdict

GENESYS_REGION = os.getenv("GENESYS_REGION", "mypurecloud.com")
CLIENT_ID = os.getenv("GENESYS_CLIENT_ID")
CLIENT_SECRET = os.getenv("GENESYS_CLIENT_SECRET")

LOGIN_URL = f"https://login.{GENESYS_REGION}"
API_URL = f"https://api.{GENESYS_REGION}"

INTERVAL = "2026-03-31T05:00:00.000Z/2026-04-01T05:00:00.000Z"
PAGE_SIZE = 50
PERU_TZ = timezone(timedelta(hours=-5))


def get_access_token():
    resp = requests.post(
        f"{LOGIN_URL}/oauth/token",
        data={"grant_type": "client_credentials"},
        auth=(CLIENT_ID, CLIENT_SECRET),
        timeout=30
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def build_body(page_number: int):
    return {
        "order": "desc",
        "orderBy": "conversationStart",
        "paging": {
            "pageSize": PAGE_SIZE,
            "pageNumber": page_number
        },
        "interval": INTERVAL,
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


def query_page(token, page_number):
    resp = requests.post(
        f"{API_URL}/api/v2/analytics/conversations/details/query",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        json=build_body(page_number),
        timeout=60
    )
    resp.raise_for_status()
    return resp.json()


def parse_iso_z(ts: str):
    return datetime.fromisoformat(ts.replace("Z", "+00:00"))


def fetch_all_conversations():
    token = get_access_token()

    first = query_page(token, 1)
    total_hits = first.get("totalHits", 0)
    conversations = first.get("conversations", [])

    total_pages = math.ceil(total_hits / PAGE_SIZE) if total_hits else 0

    print(f"totalHits API: {total_hits}")
    print(f"Página 1: {len(conversations)}")
    print(f"Total páginas esperadas: {total_pages}")

    for page in range(2, total_pages + 1):
        data = query_page(token, page)
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


def count_by_tnotresponding_emitdate(conversations):
    counts = defaultdict(int)
    debug_rows = []

    interval_start, interval_end = INTERVAL.split("/")
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

        # Tomamos el primer emitDate dentro del intervalo si existe
        in_range = []
        out_of_range = []

        for row in metrics_found:
            dt = parse_iso_z(row["emitDate"])
            if interval_start_dt <= dt < interval_end_dt:
                in_range.append((dt, row))
            else:
                out_of_range.append((dt, row))

        if in_range:
            # si hay varios, toma el primero cronológicamente dentro del rango
            dt, row = sorted(in_range, key=lambda x: x[0])[0]
            bucket = dt.astimezone(PERU_TZ).strftime("%Y-%m-%d %H:00")
            counts[bucket] += 1
            found_total += 1

        else:
            # si todos están fuera de rango, igual lo registramos en inconsistencias
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


def export_csv(hourly, filename):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["hora_peru", "cantidad"])
        for hour, count in hourly.items():
            writer.writerow([hour, count])

    print(f"CSV generado: {filename}")


def export_debug_csv(rows, filename="not_responding_inconsistencias.csv"):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
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


def main():
    conversations = fetch_all_conversations()

    by_start = count_by_conversation_start_hour(conversations)
    by_event, debug_rows, found_total = count_by_tnotresponding_emitdate(conversations)

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

    export_csv(by_start, "not_responding_por_hora_conversation_start.csv")
    export_csv(by_event, "not_responding_por_hora_evento.csv")
    export_debug_csv(debug_rows, "not_responding_inconsistencias.csv")


if __name__ == "__main__":
    main()