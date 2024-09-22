from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.db import get_async_session
from app.core.user import current_superuser, current_user
from app.crud.charity_project import charity_project_crud
from app.crud.donation import donation_crud
from app.models import User
from app.schemas.donation import AllDotationsDB, DonationCreate, DonationDB
from app.services.investing import investing_process

router = APIRouter()


@router.get(
    '/',
    response_model=list[AllDotationsDB],
    response_model_exclude_none=True,
    dependencies=(Depends(current_superuser),),
)
async def get_all_donations(
    session: AsyncSession = Depends(get_async_session),
):
    """Только для суперюзеров.

    Возвращает список всех пожертвований."""
    all_donations = await donation_crud.get_multi(session)
    return all_donations


@router.post('/', response_model=DonationDB)
async def create_new_donation(
    obj_in: DonationCreate,
    session: AsyncSession = Depends(get_async_session),
    user: User = Depends(current_user),
):
    """Сделать пожертвование."""
    new_donation = await donation_crud.create(
        obj_in, session, user, commit=False
    )
    fill_models = await charity_project_crud.get_not_full_invested(session)
    sources = investing_process(new_donation, fill_models)
    session.add_all(sources)
    await session.commit()
    await session.refresh(new_donation)
    return new_donation


@router.get('/my', response_model=list[DonationDB])
async def get_my_donations(
    session: AsyncSession = Depends(get_async_session),
    user: User = Depends(current_user),
):
    """Вернуть список пожертвований пользователя, выполняющего запрос."""
    donations = await donation_crud.get_by_user(session, user)
    return donations
