from http import HTTPStatus

from fastapi import HTTPException
from sqlalchemy.ext.asyncio import AsyncSession

from app.crud.charity_project import charity_project_crud


async def check_name_dublicate(project_name: str, session: AsyncSession):
    project_id = await charity_project_crud.get_charity_project_id_by_name(
        project_name, session
    )
    if project_id:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='Проект с таким именем уже существует!',
        )


async def check_charity_project_exists(project_id: int, session: AsyncSession):
    charity_project = await charity_project_crud.get_by_id(project_id, session)
    if not charity_project:
        raise HTTPException(
            status_code=HTTPStatus.NOT_FOUND,
            detail='Такого проекта не существует!',
        )
    return charity_project


def check_invested_sum(invested_amount: int, new_full_amount: int):
    if invested_amount > new_full_amount:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail=(
                'Нелья установить значение full_amount'
                'меньше уже вложенной суммы.'
            ),
        )


def check_project_closed(fully_invested: bool):
    if fully_invested:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='Закрытый проект нельзя редактировать!',
        )


def check_alredy_invested(invested: bool):
    if invested > 0:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='В проект были внесены средства, не подлежит удалению!',
        )
