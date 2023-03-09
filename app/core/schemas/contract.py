from typing import Optional

from pydantic import BaseModel


class ContractCreate(BaseModel):
    name: str
    ca_id: Optional[str]
