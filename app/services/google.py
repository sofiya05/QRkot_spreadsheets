from copy import deepcopy
from datetime import datetime

from aiogoogle import Aiogoogle

from app.core.config import settings
from app.models import CharityProject

DATE_TIME_FORMAT = '%Y/%m/%d %H:%M:%S'
COLUMNS = 5
ROWS = 100
SPREADSHEET_TEMPLATE = {
    'properties': {
        'title': 'Отчёт от {now_date_time}',
        'locale': 'ru_RU',
    },
    'sheets': [
        {
            'properties': {
                'sheetType': 'GRID',
                'sheetId': 0,
                'title': 'Лист1',
                'gridProperties': {
                    'rowCount': ROWS,
                    'columnCount': COLUMNS,
                },
            }
        }
    ],
}
TABLE_VALUES_TEMPLATE = [
    ['Отчёт от', '{now_date_time}'],
    ['Топ проектов по скорости закрытия.'],
    ['Название проекта', 'Время сбора', 'Описание'],
]


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    """Функция создания таблицы."""
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheet_body = deepcopy(SPREADSHEET_TEMPLATE)
    spreadsheet_body['properties']['title'] = spreadsheet_body['properties'][
        'title'
    ].format(now_date_time=datetime.now().strftime(DATE_TIME_FORMAT))
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )

    return response['spreadsheetId'], response['spreadsheetUrl']


async def set_user_permissions(
    spreadsheet_id: str, wrapper_services: Aiogoogle
):
    """
    Функция для предоставления прав доступа
    вашему личному аккаунту к созданному документу.
    """
    permissions_body = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': settings.email,
    }
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheet_id, json=permissions_body, fields='id'
        )
    )


async def spreadsheets_update_value(
    spreadsheet_id: str,
    projects: list[CharityProject],
    wrapper_services: Aiogoogle,
):
    service = await wrapper_services.discover('sheets', 'v4')
    table_values = deepcopy(TABLE_VALUES_TEMPLATE)
    table_values[0][1] = table_values[0][1].format(
        now_date_time=datetime.now().strftime(DATE_TIME_FORMAT)
    )

    table_values.extend(
        [
            str(project.name),
            str(project.close_date - project.create_date),
            str(project.description),
        ]
        for project in projects
    )

    total_rows = len(table_values)
    if total_rows > ROWS:
        raise ValueError(
            (
                'Таблица не поместится: '
                f'требуется {total_rows} строк, доступно {ROWS}.'
            )
        )

    total_columns = max(map(len, table_values))
    if total_columns > COLUMNS:
        raise ValueError(
            (
                'Таблица не поместится: '
                f'требуется {total_columns} колонок, доступно {COLUMNS}.'
            )
        )

    update_body = {'majorDimension': 'ROWS', 'values': table_values}
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheet_id,
            range=f'A1C1:A{total_rows}C{total_columns}',
            valueInputOption='USER_ENTERED',
            json=update_body,
        )
    )
