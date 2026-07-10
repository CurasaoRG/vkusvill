from __future__ import annotations

import imaplib
import email
from typing import Iterator, List
import re
import bisect
from src.models import RawMessage



class ImapClient:
    """
    Context-менеджер для работы с Gmail-IMAP.
    Пример:
        with ImapClient("user@gmail.com", "app_password", "INBOX") as client:
            for msg in client.fetch_new_raw(since_uid=42):
                ...
    """

    def __init__(self, username: str, password: str, mailbox: str):
        self.username = username
        self.password = password
        self.mailbox = mailbox
        self._imap: imaplib.IMAP4_SSL | None = None
        self._allowed_cache: set[int] | None = None   # кэш UID-ов
        # imaplib.Debug = 4 # debug option, shows mailbox requests

    def _ensure_cache(self) -> set[int]:
        if self._allowed_cache is not None:
            return self._allowed_cache

        self._imap.literal = u"ВКУСВИЛЛ".encode("utf-8")
        typ, data = self._imap.uid('SEARCH', 'CHARSET UTF-8', 'OR (FROM "noreply-cloudkassir@cp.ru") SUBJECT')
        if typ != "OK":
            raise RuntimeError("IMAP search failed")
        self._allowed_cache  = sorted(set(map(int, data[0].decode().split())))
        return self._allowed_cache

    def refresh_cache(self) -> None:
        self._allowed_cache = None

    def _find_next_uid(self, N):
        # allowed = self._ensure_cache()
        index = bisect.bisect_right(self._allowed_cache, N)
        if index < len(self._allowed_cache):
            return index
        return None

    # ---------- context-manager ----------
    def __enter__(self) -> ImapClient:
        self._imap = imaplib.IMAP4_SSL("imap.gmail.com")
        self._imap.login(self.username, self.password)
        self._imap.select(self.mailbox)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._imap:
            try:
                self._imap.close()
                self._imap.logout()
            except imaplib.IMAP4.error:
                pass  # уже закрыто или ошибка сети
            self._imap = None

    # ---------- public API ----------
    def fetch_new_raw(self, since_uid: int = 0) -> Iterator[RawMessage]:
        allowed = self._ensure_cache()
        index = self._find_next_uid(since_uid)
        for uid in sorted(allowed[index:]):
            msg_body = None
            if uid < since_uid:
                continue
            typ, msg_data = self._imap.uid("FETCH", str(uid), "(RFC822)")
            if typ != "OK":
                continue
            for part in msg_data:
                if isinstance(part, tuple):
                    msg = email.message_from_bytes(part[1])
            mail_from = re.search(r"[\w.-]+@[\w.-]+", msg['from']).group(0)
            if msg.is_multipart():
                for part in msg.walk():
                    if "text/" in part.get_content_type() and not msg_body:
                        msg_body = part.get_payload(decode=True)
                    if "application/pdf" in part.get_content_type():
                        msg_body = part.get_payload(decode=True)
            else:
                msg_body = msg.get_payload(decode=True).decode()
            yield RawMessage(uid=str(uid), mail_from=mail_from, raw_bytes=msg_body)