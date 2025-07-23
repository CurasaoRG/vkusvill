from __future__ import annotations

import imaplib
import email
from dataclasses import dataclass
from typing import Iterator, List


@dataclass
class RawMessage:
    uid: str
    mail_from: str
    raw_bytes: bytes


class ImapClient:
    """
    Context-менеджер для работы с Gmail-IMAP.
    Пример:
        with ImapClient("user@gmail.com", "app_password", "INBOX") as client:
            for msg in client.fetch_new_raw(since_uid=42):
                ...
    """

    def __init__(self, username: str, password: str, mailbox: str = "INBOX"):
        self.username = username
        self.password = password
        self.mailbox = mailbox
        self._imap: imaplib.IMAP4_SSL | None = None

    # ---------- context-manager ----------
    def __enter__(self) -> "ImapClient":
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

    # ---------- helpers ----------
    def _search_uids(self, criteria: str = "ALL") -> List[str]:
        """Возвращает список UID'ов по критерию."""
        if not self._imap:
            raise RuntimeError("Client not connected")

        typ, data = self._imap.uid("SEARCH", None, criteria)
        if typ != "OK":
            raise RuntimeError("IMAP search failed")
        # data[0] -> b'1 2 3 ...'
        return data[0].decode().split()

    # ---------- public API ----------
    def fetch_new_raw(self, since_uid: int = 0) -> Iterator[RawMessage]:
        """
        Генератор «сырых» писем с UID > since_uid.
        Отдаёт RawMessage в порядке возрастания UID.
        """
        criteria = f"UID {since_uid + 1}:*"
        for uid in sorted(map(int, self._search_uids(criteria))):
            typ, msg_data = self._imap.uid("FETCH", str(uid), "(RFC822)")
            if typ != "OK":
                continue  # пропускаем ошибки
            for part in msg_data:
                if isinstance(part, tuple):
                    raw_bytes = part[1]
                    parsed = email.message_from_bytes(raw_bytes)
                    mail_from = str(parsed.get("From", ""))
                    yield RawMessage(uid=str(uid), mail_from=mail_from, raw_bytes=raw_bytes)