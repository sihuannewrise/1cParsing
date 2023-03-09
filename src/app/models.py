from datetime import datetime
from typing import List
from sqlalchemy import ForeignKey
from sqlalchemy.orm import Mapped, mapped_column, relationship

from src.app.db import Base


class Contract(Base):
    name: Mapped[str]
    ca_id: Mapped[int] = mapped_column(ForeignKey('counteragent.id'))
    expenses: Mapped[List['Expenses']] = relationship(backref='contract')


class CounterAgent(Base):
    name: Mapped[str]
    contracts: Mapped[List[Contract]] = relationship(backref='counteragent')
    expenses: Mapped[List['Expenses']] = relationship(backref='counteragent')


class Expenses(Base):
    period: Mapped[datetime]
    ca_id: Mapped[int] = mapped_column(ForeignKey('counteragent.id'))
    contract_id: Mapped[int] = mapped_column(ForeignKey('contract.id'))
    amount: Mapped[float]
