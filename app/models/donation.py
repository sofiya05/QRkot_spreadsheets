from sqlalchemy import Column, ForeignKey, Integer, Text

from app.models.base import BaseCharityDonationModel


class Donation(BaseCharityDonationModel):
    user_id = Column(Integer, ForeignKey('user.id'))
    comment = Column(Text)

    def __repr__(self):
        return (
            f'Id пользователя: {self.user_id}, '
            f'Комментарий: {self.comment}, '
            f'{super().__repr__()}'
        )
