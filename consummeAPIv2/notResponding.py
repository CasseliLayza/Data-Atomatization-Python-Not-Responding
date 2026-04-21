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


def count_by_tnotresponding_emitdate(conversations):
    counts = defaultdict(int)

    for conv in conversations:
        found = False

        for participant in conv.get("participants", []):
            if participant.get("purpose") != "agent":
                continue

            for session in participant.get("sessions", []):
                for metric in session.get("metrics", []):
                    if metric.get("name") == "tNotResponding":
                        ts = metric.get("emitDate")
                        if ts:
                            dt_peru = parse_iso_z(ts).astimezone(PERU_TZ)
                            bucket = dt_peru.strftime("%Y-%m-%d %H:00")
                            counts[bucket] += 1
                            found = True
                            break
                if found:
                    break
            if found:
                break

    return dict(sorted(counts.items()))


def export_csv(hourly, filename):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["hora_peru", "cantidad"])
        for hour, count in hourly.items():
            writer.writerow([hour, count])

    print(f"CSV generado: {filename}")


def main():
    conversations = fetch_all_conversations()

    by_start = count_by_conversation_start_hour(conversations)
    by_event = count_by_tnotresponding_emitdate(conversations)

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

    export_csv(by_start, "not_responding_por_hora_conversation_start.csv")
    export_csv(by_event, "not_responding_por_hora_evento.csv")


if __name__ == "__main__":
    main()