from datetime import datetime

from app.models.base import BaseCharityDonationModel


def investing_process(
    target: BaseCharityDonationModel,
    sources: list[BaseCharityDonationModel],
) -> list[BaseCharityDonationModel]:

    modified_sources = []

    for source in sources:
        transfer_amount = min(
            target.full_amount - target.invested_amount,
            source.full_amount - source.invested_amount,
        )

        for obj in [target, source]:
            obj.invested_amount += transfer_amount
            if obj.full_amount == obj.invested_amount:
                obj.fully_invested = True
                obj.close_date = datetime.now()

        modified_sources.append(source)

        if target.fully_invested:
            break

    return modified_sources
