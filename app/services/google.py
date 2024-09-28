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
                    'rowCount': 0,
                    'columnCount': 0,
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
    spreadsheet_body = SPREADSHEET_TEMPLATE.copy()
    spreadsheet_body['properties']['title'] = spreadsheet_body['properties'][
        'title'
    ].format(now_date_time=datetime.now().strftime(DATE_TIME_FORMAT))
    spreadsheet_body['sheets'][0]['properties']['gridProperties'][
        'rowCount'
    ] = ROWS
    spreadsheet_body['sheets'][0]['properties']['gridProperties'][
        'columnCount'
    ] = COLUMNS
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )

    return (
        response['spreadsheetId'],
        f'https://docs.google.com/spreadsheets/d/{response["spreadsheetId"]}/edit',
    )


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
    table_values = [row.copy() for row in TABLE_VALUES_TEMPLATE]
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

    update_body = {'majorDimension': 'ROWS', 'values': table_values}
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheet_id,
            range=f'A1C1:A{len(table_values)}C{max(map(len, table_values))}',
            valueInputOption='USER_ENTERED',
            json=update_body,
        )
    )
