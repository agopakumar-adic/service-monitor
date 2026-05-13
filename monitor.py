"""
Service Outage Monitor
Checks status of 9 services and posts results to Microsoft Teams via Power Automate webhook.
"""

import os
import sys
import json
import urllib.request
import urllib.error
from datetime import datetime, timezone

# Ensure stdout handles Unicode (needed on Windows with cp1252 console)
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TEAMS_WEBHOOK_URL = os.environ.get("TEAMS_WEBHOOK_URL", "")

# Services with Atlassian Statuspage JSON APIs
STATUSPAGE_SERVICES = [
    {"name": "Azure",            "url": "https://azure.status.microsoft/api/v2/status.json"},
    {"name": "ChatGPT (OpenAI)", "url": "https://status.openai.com/api/v2/status.json"},
    {"name": "Claude (Anthropic)","url": "https://status.anthropic.com/api/v2/status.json"},
    {"name": "MS Copilot Studio", "url": "https://status.microsoft365.com/api/v2/status.json"},
    {"name": "FactSet",           "url": "https://status.factset.com/api/v2/status.json"},
    {"name": "OCI (Oracle Cloud)","url": "https://ocistatus.oraclecloud.com/api/v2/status.json"},
    {"name": "DocuSign",          "url": "https://status.docusign.com/api/v2/status.json"},
    {"name": "Snowflake",         "url": "https://status.snowflake.com/api/v2/status.json"},
]

BLOOMBERG_URL = "https://www.bloomberg.com/company/bloomberg-status/"
BLOOMBERG_OUTAGE_KEYWORDS = ["outage", "incident", "degraded", "disruption", "down", "investigating"]

INDICATOR_EMOJI = {
    "none":     "✅",
    "minor":    "⚠️",
    "major":    "🔴",
    "critical": "🔴",
    "unknown":  "⚠️",
}


def fetch_json(url: str, timeout: int = 10) -> dict | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "ServiceMonitor/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode())
    except Exception:
        return None


def fetch_html(url: str, timeout: int = 10) -> str | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "ServiceMonitor/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.read().decode(errors="ignore")
    except Exception:
        return None


def check_statuspage(service: dict) -> dict:
    data = fetch_json(service["url"])
    if data is None:
        return {"name": service["name"], "indicator": "unknown", "description": "Unreachable / No response"}
    try:
        indicator = data["status"]["indicator"]
        description = data["status"]["description"]
    except (KeyError, TypeError):
        indicator = "unknown"
        description = "Unexpected response format"
    return {"name": service["name"], "indicator": indicator, "description": description}


def check_bloomberg() -> dict:
    html = fetch_html(BLOOMBERG_URL)
    if html is None:
        return {"name": "Bloomberg", "indicator": "unknown", "description": "Unreachable / No response"}
    html_lower = html.lower()
    found = [kw for kw in BLOOMBERG_OUTAGE_KEYWORDS if kw in html_lower]
    if found:
        return {"name": "Bloomberg", "indicator": "minor", "description": f"Possible issue detected (keywords: {', '.join(found)})"}
    return {"name": "Bloomberg", "indicator": "none", "description": "No outage keywords detected"}


def build_teams_message(results: list[dict], has_issues: bool) -> dict:
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    if has_issues:
        title = f"🚨 Service Outage Alert — {timestamp}"
        title_color = "attention"
    else:
        title = f"✅ Hourly Service Check: All Systems Operational — {timestamp}"
        title_color = "good"

    rows = []
    for r in results:
        emoji = INDICATOR_EMOJI.get(r["indicator"], "⚠️")
        rows.append({
            "type": "ColumnSet",
            "columns": [
                {
                    "type": "Column",
                    "width": "auto",
                    "items": [{"type": "TextBlock", "text": emoji, "wrap": False}]
                },
                {
                    "type": "Column",
                    "width": "stretch",
                    "items": [{"type": "TextBlock", "text": f"**{r['name']}**: {r['description']}", "wrap": True}]
                }
            ]
        })

    body = [
        {"type": "TextBlock", "text": title, "weight": "Bolder", "size": "Medium", "color": title_color, "wrap": True},
        {"type": "TextBlock", "text": " ", "spacing": "Small"},
    ] + rows

    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.2",
        "body": body
    }


def post_to_teams(payload: dict) -> bool:
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        TEAMS_WEBHOOK_URL,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            status = resp.status
            print(f"Teams webhook response: HTTP {status}")
            return status < 300
    except urllib.error.HTTPError as e:
        print(f"Teams webhook HTTP error: {e.code} {e.reason}")
        body_text = e.read().decode(errors="ignore")
        print(f"Response body: {body_text}")
        return False
    except Exception as e:
        print(f"Teams webhook error: {e}")
        return False


def run():
    if not TEAMS_WEBHOOK_URL:
        print("ERROR: TEAMS_WEBHOOK_URL environment variable is not set.")
        sys.exit(1)

    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Running service status checks...")

    results = []
    for service in STATUSPAGE_SERVICES:
        result = check_statuspage(service)
        emoji = INDICATOR_EMOJI.get(result["indicator"], "⚠️")
        print(f"  {emoji} {result['name']}: {result['description']}")
        results.append(result)

    bloomberg = check_bloomberg()
    emoji = INDICATOR_EMOJI.get(bloomberg["indicator"], "⚠️")
    print(f"  {emoji} {bloomberg['name']}: {bloomberg['description']}")
    results.append(bloomberg)

    has_issues = any(r["indicator"] != "none" for r in results)

    payload = build_teams_message(results, has_issues)
    print("\nPosting to Teams...")
    success = post_to_teams(payload)

    if success:
        print("Teams notification sent successfully.")
    else:
        print("Failed to send Teams notification.")


if __name__ == "__main__":
    run()
