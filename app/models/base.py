from datetime import datetime

from sqlalchemy import Boolean, CheckConstraint, Column, DateTime, Integer

from app.core.db import Base


class BaseCharityDonationModel(Base):
    __abstract__ = True

    full_amount = Column(Integer, nullable=False)
    invested_amount = Column(Integer, default=0)
    fully_invested = Column(Boolean, default=False)
    create_date = Column(DateTime, default=datetime.now)
    close_date = Column(DateTime)

    __table_args__ = (
        CheckConstraint(
            'full_amount > 0',
            name='check_full_amount_and_invested_amount_positive',
        ),
        CheckConstraint(
            '0 <= invested_amount <= full_amount',
            name='check_invested_amount_less_equal_full_amount',
        ),
    )
