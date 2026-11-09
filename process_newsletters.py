"""
Vastgoed Nieuwsarchief — Verwerkingsscript
==========================================
Leest onverwerkte e-mails uit Outlook (Microsoft Graph), classificeert ze
met Claude en slaat de artikelen op in Supabase.

Vereiste omgevingsvariabelen (.env bestand):
  ANTHROPIC_API_KEY        — Anthropic API sleutel
  SUPABASE_URL             — https://ojjhwbqjphjllqhetndz.supabase.co
  SUPABASE_SERVICE_ROLE_KEY — Supabase service role key (Settings > API)
  MS_TENANT_ID             — Azure tenant ID (zie SETUP.md)
  MS_CLIENT_ID             — Azure app client ID
  MS_CLIENT_SECRET         — Azure app client secret

Installatie:
  pip install anthropic supabase msal python-dotenv --break-system-packages

Gebruik:
  python process_newsletters.py
  python process_newsletters.py --days 7    # verwerk laatste 7 dagen
"""

import os
import sys
import json
import logging
import argparse
from datetime import datetime, timedelta, timezone

from dotenv import load_dotenv
import msal
import requests
import anthropic
from supabase import create_client

load_dotenv()
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

# ── Constanten ──────────────────────────────────────────────────────────────

SUPABASE_URL  = os.getenv("SUPABASE_URL", "https://ojjhwbqjphjllqhetndz.supabase.co")
SUPABASE_KEY  = os.getenv("SUPABASE_SERVICE_ROLE_KEY")
ANTHROPIC_KEY = os.getenv("ANTHROPIC_API_KEY")
MS_TENANT     = os.getenv("MS_TENANT_ID")
MS_CLIENT_ID  = os.getenv("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
MS_SCOPE      = ["https://graph.microsoft.com/.default"]
GRAPH_BASE    = "https://graph.microsoft.com/v1.0"

# Nieuwsbrief-afzenders: pas aan op je eigen inbox
NEWSLETTER_SENDERS = [
    # Voeg hier e-mailadressen of domeinen toe, bijv:
    # "nieuwsbrief@vastgoedmarkt.nl",
    # "noreply@jll.com",
    # "research@cbre.com",
]

ASSET_CLASSES = ["kantoren", "retail", "logistiek", "woningen", "hotels_leisure", "parkeren", "overig"]

CLASSIFY_PROMPT = """Je bent een vastgoed-analist. Analyseer de onderstaande nieuwsbrief-tekst en extraheer alle afzonderlijke artikelen of nieuwsitems.

Voor elk artikel geef je terug:
- title: bondige Nederlandse titel (max 100 tekens)
- summary: samenvatting van 2-3 zinnen in het Nederlands
- full_text: de volledige tekst van het artikel (behoud zoveel mogelijk van de originele tekst)
- url: de directe URL van het artikel (alleen als expliciet aanwezig in de tekst, anders null)
- asset_classes: lijst van toepasselijke asset classes uit: kantoren, retail, logistiek, woningen, hotels_leisure, parkeren, overig
- topics: lijst van 2-5 specifieke onderwerpen (bijv. "huurprijzen", "ESG", "transactie", "nieuwbouw", "regulering")
- published_date: publicatiedatum in YYYY-MM-DD formaat (als vermeld, anders de datum van de e-mail)

Geef je antwoord terug als een JSON array: [{"title": ..., "summary": ..., ...}, ...]
Als de e-mail geen relevant vastgoednieuws bevat, geef dan een lege array [].

Nieuwsbrief onderwerp: {subject}
Datum: {date}

Tekst:
{body}
"""

# ── Microsoft Graph Auth ────────────────────────────────────────────────────

def get_graph_token():
    app = msal.ConfidentialClientApplication(
        MS_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{MS_TENANT}",
        client_credential=MS_CLIENT_SECRET
    )
    result = app.acquire_token_silent(MS_SCOPE, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=MS_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(f"Kon geen Graph token verkrijgen: {result.get('error_description')}")
    return result["access_token"]

def graph_get(token, path, params=None):
    resp = requests.get(
        f"{GRAPH_BASE}{path}",
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        timeout=30
    )
    resp.raise_for_status()
    return resp.json()

def graph_patch(token, path, body):
    resp = requests.patch(
        f"{GRAPH_BASE}{path}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body,
        timeout=30
    )
    resp.raise_for_status()

def get_archive_folder_id(token):
    """Haal het ID op van de Archive map (aanmaken als die niet bestaat)."""
    try:
        data = graph_get(token, "/me/mailFolders/archive")
        return data["id"]
    except Exception:
        # Maak Archive map aan
        resp = requests.post(
            f"{GRAPH_BASE}/me/mailFolders",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"displayName": "Nieuwsbrieven Archief"},
            timeout=30
        )
        resp.raise_for_status()
        return resp.json()["id"]

def move_to_archive(token, message_id, folder_id):
    requests.post(
        f"{GRAPH_BASE}/me/messages/{message_id}/move",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"destinationId": folder_id},
        timeout=30
    ).raise_for_status()

# ── Supabase helpers ────────────────────────────────────────────────────────

def get_processed_ids(sb):
    result = sb.table("newsletters").select("outlook_message_id").execute()
    return {row["outlook_message_id"] for row in (result.data or [])}

def save_newsletter(sb, msg_id, subject, sender_name, sender_email, received_at, body_text, body_html):
    result = sb.table("newsletters").insert({
        "outlook_message_id": msg_id,
        "subject": subject,
        "sender_name": sender_name,
        "sender_email": sender_email,
        "received_at": received_at,
        "body_text": body_text,
        "body_html": body_html,
    }).execute()
    return result.data[0]["id"]

def save_articles(sb, newsletter_id, articles):
    for art in articles:
        sb.table("articles").insert({
            "newsletter_id": newsletter_id,
            "title": art.get("title", "Zonder titel"),
            "summary": art.get("summary"),
            "full_text": art.get("full_text"),
            "url": art.get("url"),
            "asset_classes": art.get("asset_classes", ["overig"]),
            "topics": art.get("topics", []),
            "published_date": art.get("published_date"),
        }).execute()

def mark_archived(sb, newsletter_id):
    sb.table("newsletters").update({"archived_in_outlook": True, "processed_at": datetime.now(timezone.utc).isoformat()}).eq("id", newsletter_id).execute()

# ── Claude classificatie ────────────────────────────────────────────────────

def classify_email(client, subject, date_str, body_text):
    prompt = CLASSIFY_PROMPT.format(
        subject=subject,
        date=date_str,
        body=body_text[:15000]  # max tokens bewaken
    )
    message = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = message.content[0].text.strip()
    # Extract JSON array uit antwoord
    start = raw.find("[")
    end = raw.rfind("]") + 1
    if start == -1 or end == 0:
        return []
    return json.loads(raw[start:end])

# ── Hoofdfunctie ────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Verwerk nieuwsbrieven naar Supabase")
    parser.add_argument("--days", type=int, default=1, help="Aantal dagen terug om te verwerken (default: 1)")
    parser.add_argument("--dry-run", action="store_true", help="Verwerk maar sla niet op")
    args = parser.parse_args()

    # Valideer env vars
    missing = [v for v in ["SUPABASE_SERVICE_ROLE_KEY", "ANTHROPIC_API_KEY", "MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET"] if not os.getenv(v)]
    if missing:
        log.error(f"Ontbrekende omgevingsvariabelen: {', '.join(missing)}")
        log.error("Maak een .env bestand op basis van de instructies bovenaan dit script.")
        sys.exit(1)

    sb = create_client(SUPABASE_URL, SUPABASE_KEY)
    claude = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
    token = get_graph_token()
    archive_folder = get_archive_folder_id(token)

    # Haal al verwerkte IDs op
    processed_ids = get_processed_ids(sb)
    log.info(f"Al {len(processed_ids)} nieuwsbrieven verwerkt in database")

    # Haal e-mails op uit Outlook
    since = (datetime.now(timezone.utc) - timedelta(days=args.days)).isoformat()
    params = {
        "$filter": f"receivedDateTime ge {since} and isRead eq false",
        "$select": "id,subject,from,receivedDateTime,body",
        "$top": 50,
        "$orderby": "receivedDateTime asc"
    }

    # Filter op specifieke afzenders als geconfigureerd
    if NEWSLETTER_SENDERS:
        sender_filter = " or ".join([f"from/emailAddress/address eq '{s}'" for s in NEWSLETTER_SENDERS])
        params["$filter"] += f" and ({sender_filter})"

    messages_data = graph_get(token, "/me/messages", params)
    emails = messages_data.get("value", [])
    log.info(f"Gevonden: {len(emails)} e-mails in de laatste {args.days} dag(en)")

    new_emails = [e for e in emails if e["id"] not in processed_ids]
    log.info(f"Nieuw te verwerken: {len(new_emails)}")

    stats = {"processed": 0, "articles_saved": 0, "skipped": 0, "errors": 0}

    for email in new_emails:
        msg_id = email["id"]
        subject = email.get("subject", "(geen onderwerp)")
        sender = email.get("from", {}).get("emailAddress", {})
        sender_name = sender.get("name", "")
        sender_email = sender.get("address", "")
        received_at = email.get("receivedDateTime")
        body = email.get("body", {})
        body_text = body.get("content", "") if body.get("contentType") == "text" else ""
        body_html = body.get("content", "") if body.get("contentType") == "html" else ""

        # Gebruik platte tekst voor classificatie (verwijder HTML tags)
        if not body_text and body_html:
            import re
            body_text = re.sub(r'<[^>]+>', ' ', body_html)
            body_text = re.sub(r'\s+', ' ', body_text).strip()

        if not body_text or len(body_text) < 100:
            log.info(f"Overgeslagen (te kort): {subject}")
            stats["skipped"] += 1
            continue

        log.info(f"Verwerken: {subject} ({sender_email})")

        try:
            articles = classify_email(claude, subject, received_at[:10] if received_at else "", body_text)
            log.info(f"  → {len(articles)} artikel(en) gevonden")

            if not args.dry_run:
                newsletter_id = save_newsletter(sb, msg_id, subject, sender_name, sender_email, received_at, body_text, body_html)
                save_articles(sb, newsletter_id, articles)
                mark_archived(sb, newsletter_id)
                move_to_archive(token, msg_id, archive_folder)
                stats["articles_saved"] += len(articles)

            stats["processed"] += 1

        except Exception as e:
            log.error(f"Fout bij verwerken {subject}: {e}")
            stats["errors"] += 1

    log.info(f"\n✅ Klaar: {stats['processed']} verwerkt, {stats['articles_saved']} artikelen opgeslagen, {stats['skipped']} overgeslagen, {stats['errors']} fouten")


if __name__ == "__main__":
    main()
