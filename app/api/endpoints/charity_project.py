from datetime import datetime

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.api.validators import (
    check_alredy_invested,
    check_charity_project_exists,
    check_invested_sum,
    check_name_dublicate,
    check_project_closed,
)
from app.core.db import get_async_session
from app.core.user import current_superuser
from app.crud.charity_project import charity_project_crud
from app.crud.donation import donation_crud
from app.schemas.charity_project import (
    CharityProjectCreate,
    CharityProjectDB,
    CharityProjectUpdate,
)
from app.services.investing import investing_process

router = APIRouter()


@router.get(
    '/',
    response_model=list[CharityProjectDB],
    response_model_exclude_none=True,
)
async def get_all_charity_projects(
    session: AsyncSession = Depends(get_async_session),
):
    """Возвращает список всех проектов."""
    all_charity_projects = await charity_project_crud.get_multi(session)
    return all_charity_projects


@router.post(
    '/',
    response_model=CharityProjectDB,
    response_model_exclude_none=True,
    dependencies=(Depends(current_superuser),),
)
async def create_new_charity_project(
    obj_in: CharityProjectCreate,
    session: AsyncSession = Depends(get_async_session),
):
    """Только для суперюзеров.

    Создаёт благотворительный проект."""
    await check_name_dublicate(obj_in.name, session)
    new_charity_project = await charity_project_crud.create(
        obj_in, session, commit=False
    )
    fill_models = await donation_crud.get_not_full_invested(session)
    sources = investing_process(new_charity_project, fill_models)
    session.add_all(sources)
    await session.commit()
    await session.refresh(new_charity_project)
    return new_charity_project


@router.patch(
    '/{project_id}',
    response_model=CharityProjectDB,
    response_model_exclude_none=True,
    dependencies=(Depends(current_superuser),),
)
async def update_charity_project(
    project_id: int,
    obj_in: CharityProjectUpdate,
    session: AsyncSession = Depends(get_async_session),
):
    """Только для суперюзеров.

    Закрытый проект нельзя редактировать;
    нельзя установить требуемую сумму меньше уже вложенной.
    """

    charity_project = await check_charity_project_exists(project_id, session)
    check_project_closed(charity_project.fully_invested)
    if obj_in.name:
        await check_name_dublicate(obj_in.name, session)
    if obj_in.full_amount:
        check_invested_sum(charity_project.invested_amount, obj_in.full_amount)

    if obj_in.full_amount == charity_project.invested_amount:
        charity_project.fully_invested = True
        charity_project.close_date = datetime.now()

    charity_project = await charity_project_crud.update(
        charity_project, obj_in, session
    )

    return charity_project


@router.delete(
    '/{project_id}',
    response_model=CharityProjectDB,
    dependencies=(Depends(current_superuser),),
)
async def delete_charity_project(
    project_id: int, session: AsyncSession = Depends(get_async_session)
):
    """Только для суперюзеров.

    Удаляет проект.
    Нельзя удалить проект, в который уже были инвестированы средства,
    его можно только закрыть.
    """
    charity_project = await check_charity_project_exists(project_id, session)
    check_alredy_invested(charity_project.invested_amount)
    charity_project = await charity_project_crud.remove(
        charity_project, session
    )
    return charity_project
