from datetime import datetime
from typing import List
from sqlalchemy import ForeignKey
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.core.db import Base


class Contract(Base):
    name: Mapped[str]
    ca_id: Mapped[int] = mapped_column(ForeignKey('counteragent.id'))
    expenses: Mapped[List['Expense']] = relationship(backref='contract')


class CounterAgent(Base):
    name: Mapped[str]
    contracts: Mapped[List[Contract]] = relationship(backref='counteragent')
    expenses: Mapped[List['Expense']] = relationship(backref='counteragent')


class Expense(Base):
    period: Mapped[datetime]
    ca_id: Mapped[int] = mapped_column(ForeignKey('counteragent.id'))
    contract_id: Mapped[int] = mapped_column(ForeignKey('contract.id'))
    amount: Mapped[float]
