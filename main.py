import os
from pathlib import Path
from dotenv import dotenv_values
from vkusvill.imap_client import ImapClient
from vkusvill.parsers import parse_message
from vkusvill.repository import CsvRepository
from vkusvill.increment import IncrementHandler
from vkusvill.gs_uploader import GoogleSheetsUploader

ROOT = Path(__file__).parent
config = dotenv_values(ROOT / ".env")

repo = CsvRepository(ROOT / "data")
uploader = GoogleSheetsUploader(ROOT / "credentials.json")

def sync_once(limit=100):
    # --- existing IMAP logic, but:
    # 1. returns raw message (bytes or str)
    # 2. calls parse_message(raw) -> Check
    # 3. repo.append_check(id, check)
    pass

if __name__ == "__main__":
    sync_once()
    uploader.upload_csv("check_info", repo.check_info)
    uploader.upload_csv("detailed_data", repo.items)


ROOT = Path(__file__).parent
config = dotenv_values(ROOT / ".env")
repo = CsvRepository(ROOT / "data")
inc = IncrementHandler(ROOT / "data" / "increment.txt")

with ImapClient(config["GMAIL_USERNAME"], config["GMAIL_PASSWORD"], config["MAILBOX"]) as client:
    for raw_msg in client.fetch_new_raw(since_uid=inc.get()):
        check = parse_message(raw_msg.raw_bytes, raw_msg.mail_from)
        repo.append_check(check_id=int(raw_msg.uid), check=check)
        inc.set(int(raw_msg.uid))

