import requests

TOKEN = "WYLIxEWurB7qEtO6EfbNJc5GQaWJFZHmp1ge50T5TMk8CmvMyOUJmfwXP4KYbGoM8NP5FbzUVj69pX-r2rn0UA"


url = "https://api.mypurecloud.com/api/v2/analytics/conversations/details/query"

body = {"order":"desc","orderBy":"conversationStart","paging":{"pageSize":50,"pageNumber":1},"interval":"2026-03-31T05:00:00.000Z/2026-04-01T05:00:00.000Z","segmentFilters":[{"type":"or","predicates":[{"dimension":"direction","value":"inbound"},{"dimension":"direction","value":"outbound"}]}],"conversationFilters":[{"type":"and","predicates":[{"type":"metric","metric":"tNotResponding","operator":"exists"}]}],"evaluationFilters":[],"surveyFilters":[]}

resp = requests.post(
    url,
    headers={
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json"
    },
    json=body,
    timeout=60
)

print(resp.status_code)
print(resp.json().get("totalHits"))