from sqlalchemy import select

from app.core.db import AsyncSessionLocal
from app.core.models import Contract
from app.core.schemas.contract import ContractCreate


async def create_contract(
    new_contract: ContractCreate
) -> Contract:
    new_contract_data = new_contract.dict()
    db_contract = Contract(**new_contract_data)
    async with AsyncSessionLocal() as session:
        session.add(db_contract)
        await session.commit()
        await session.refresh(db_contract)
    return db_contract


async def get_contract_id_by_name(contract_name):
    async with AsyncSessionLocal() as session:
        db_contract_id = await session.execute(
            select(Contract.id).where(
                Contract.name == contract_name
            )
        )
        db_contract_id = db_contract_id.scalars().first()
    return db_contract_id
