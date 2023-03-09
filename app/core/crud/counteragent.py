from typing import Optional
from sqlalchemy import select

from app.core.db import AsyncSessionLocal
from app.core.models import CounterAgent


async def create_counteragent(new_ca):
    new_ca_data = new_ca.dict()
    db_ca = CounterAgent(**new_ca_data)
    async with AsyncSessionLocal() as session:
        session.add(db_ca)
        await session.commit()
        await session.refresh(db_ca)
    return db_ca


async def get_ca_id_by_name(ca_name: str) -> Optional[int]:
    async with AsyncSessionLocal() as session:
        db_ca_id = await session.execute(
            select(CounterAgent.id).where(
                CounterAgent.name == ca_name
            )
        )
        db_ca_id = db_ca_id.scalars().first()
    return db_ca_id
