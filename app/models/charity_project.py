from sqlalchemy import Column, String, Text

from app.models.base import BaseCharityDonationModel


class CharityProject(BaseCharityDonationModel):
    name = Column(String(100), unique=True, nullable=False)
    description = Column(Text, nullable=False)

    def __repr__(self):
        return (
            f'Имя проекта: {self.name}, '
            f'Описание проекта: {self.description}, '
            f'{super().__repr__()}'
        )
