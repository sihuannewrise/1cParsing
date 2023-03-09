# from sqlalchemy import select

from core.db import AsyncSessionLocal
from core.models import Expense


async def create_expense(new_expense):
    new_expense_data = new_expense.dict()
    db_expense = Expense(**new_expense_data)

    async with AsyncSessionLocal() as session:
        session.add(db_expense)
        await session.commit()
        await session.refresh(db_expense)
    return db_expense
