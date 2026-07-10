from dataclasses import dataclass

@dataclass
class RawMessage:
    uid: str
    mail_from: str
    raw_bytes: bytes