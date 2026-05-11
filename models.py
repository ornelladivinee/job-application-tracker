from dataclasses import dataclass
from typing import Optional

@dataclass
class Application:
    company: str
    country: str
    position: str
    date_applied: str
    platform: str
    status: str
    notes: Optional[str] = ""