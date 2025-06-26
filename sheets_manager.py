from google.oauth2 import service_account
from googleapiclient.discovery import build
import httplib2
from oauth2client.service_account import ServiceAccountCredentials

class sheetAPI():
    # Укажите путь к вашему service_account.json
    SERVICE_ACCOUNT_FILE = 'kwork-buisness-data-06abc0966cd4.json'

    # Укажите нужные scope (права доступа)
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    # Создаем credentials из файла
    credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE,
                                                                   ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])

    httpAuth = credentials.authorize(httplib2.Http())  # Авторизуемся в системе

    # Создаем сервис для работы с Google Sheets
    service = build('sheets', 'v4', credentials=credentials)

    def create_table(self, title, list):
        spreadsheet = self.service.spreadsheets().create(body={
            'properties': {'title': title, 'locale': 'ru_RU'},
            'sheets': [{'properties': {'sheetType': 'GRID',
                                       'sheetId': 0,
                                       'title': list,
                                       'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
        }).execute()

        spreadsheetId = spreadsheet['spreadsheetId']  # сохраняем идентификатор файла

        respounse = ['https://docs.google.com/spreadsheets/d/' + spreadsheetId, spreadsheetId]
        return respounse

    def access(self, spreadsheetId, email):
        driveService = build('drive', 'v3', http=self.httpAuth)  # Выбираем работу с Google Drive и 3 версию API
        access = driveService.permissions().create(
            fileId=spreadsheetId,
            body={'type': 'user', 'role': 'writer', 'emailAddress': email},
            # Открываем доступ на редактирование
            fields='id'
        ).execute()
        respounse = f"Add user with email {email} to spreadsheetsId: '{spreadsheetId}'"
        return respounse

    def insert_datas(self, spreadsheetId, list_name, range, values):
        results = self.service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
            "valueInputOption": "USER_ENTERED",
            # Данные воспринимаются, как вводимые пользователем (считается значение формул)
            "data": [
                {"range": f"{list_name}!{range}",
                 "majorDimension": "ROWS",  # Сначала заполнять строки, затем столбцы
                 "values": values}
            ]
            #[
            #    ["Ячейка B2", "Ячейка C2", "Ячейка D2"],  # Заполняем первую строку
            #    ['25', "=6*6", "=sin(3,14/2)"]  # Заполняем вторую строку
            #] # пример для внесения данных
        }).execute()

    def take_datas(self, spreadsheetId, list_name, range):
        ranges = [f"{list_name}!{range}"]
        results = self.service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId,
                                                           ranges=ranges,
                                                           valueRenderOption='FORMATTED_VALUE',
                                                           dateTimeRenderOption='FORMATTED_STRING').execute()
        sheet_values = results['valueRanges'][0]['values']
        return sheet_values

    def take_settings(self):
        file = open("settings.txt", "r").read().strip()
        rows = file.split(",")
        data = []
        for i in rows:
            if i.replace(" ", "") != "":
                row = i.split(":")
                data.append((row[0].replace(" ", ""), row[-1].replace(" ", "")))
        return data