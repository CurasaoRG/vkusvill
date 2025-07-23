from .ofd import OfdParser
from .ofd1 import Ofd1Parser
from .pdf import PdfParser
from ..models import Check

def parse_message(raw_mail: bytes, mail_from: str) -> Check:
    mapping = {
        "noreply@ofd.ru": OfdParser,
        "echeck@1-ofd.ru": Ofd1Parser,
        "noreply-cloudkassir@cp.ru": PdfParser,
    }
    cls = mapping[mail_from]
    return cls().parse(raw_mail)