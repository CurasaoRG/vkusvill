from pathlib import Path
from dotenv import dotenv_values
from src.imap_client import ImapClient
from src.parsers import parse_message
from src.repository.csv import CsvRepository
from src.repository.increment import IncrementHandler
from src.loaders.gs_uploader import GoogleSheetsUploader


ROOT = Path(__file__).parent.parent
config = dotenv_values(ROOT / ".env")
repo = CsvRepository(ROOT / "data")
inc = IncrementHandler(ROOT / "data" / "increment.txt")
uploader = GoogleSheetsUploader(ROOT / "credentials.json")
error_file = ROOT / "logs" /"error.log"

if __name__ == "__main__":
    with ImapClient(config["GMAIL_USERNAME"], config["GMAIL_PASSWORD"], config["MAILBOX"]) as client:
        for raw_msg in client.fetch_new_raw(since_uid=inc.get()):
            try:
                check = parse_message(raw_msg.raw_bytes, raw_msg.mail_from)
                repo.append_check(check_id=int(raw_msg.uid), check=check)
                inc.set(int(raw_msg.uid))
                print(f"PARSED: check_type = {raw_msg.mail_from}, uid = {raw_msg.uid}")
            except Exception as e:
                with error_file.open('a')as f:
                    f.write(f"PARSING ERROR: check_type = {raw_msg.mail_from}, uid = {raw_msg.uid}\n")
                    f.write(f"ERROR: {e}\n")
    uploader.upload_csv("check_info", repo.check_info)
    uploader.upload_csv("detailed_data", repo.items)