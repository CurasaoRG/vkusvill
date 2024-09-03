from oauth2client.service_account import ServiceAccountCredentials
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import csv

CREDENTIALS_FILE = '/home/rg/Documents/Study/PET_projects/Vkusvill/credentials.json'  
CHECK_INFO_LOCATION = '/home/rg/Documents/Study/PET_projects/Vkusvill/check_info.csv'
DATA_LOCATION = '/home/rg/Documents/Study/PET_projects/Vkusvill/data.csv'

def get_spreadsheet(spreadsheet_name="Вкусвилл"):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    try:
        spreadsheet = client.open(spreadsheet_name)
        print(f"Файл открыт: {spreadsheet.url}")
    except gspread.SpreadsheetNotFound:
        spreadsheet = client.create(spreadsheet_name)
        # Печатаем ссылку на созданный файл
        print(f"Файл создан: {spreadsheet.url}")
    return spreadsheet
    
def get_sheet(spreadsheet, worksheet_name):
    if worksheet_name not in [ws.title for ws in spreadsheet.worksheets()]:
        spreadsheet.add_worksheet(title=worksheet_name,rows=1,cols=1)        
    sheet = spreadsheet.worksheet(worksheet_name)
    # Шаг 5: Предоставление доступа на чтение всем пользователям
    spreadsheet.share('', perm_type='anyone', role='reader')
    print(f"выдал доступ")
    return sheet

def write_to_sheet(sheet, file_location):
    # Шаг 3: Чтение данных из CSV файла
    with open(file_location, mode='r') as file:
        reader = csv.reader(file, delimiter=';')
        csv_data = list(reader)
    print('прочитано')
    # Шаг 4: Очистка листа и загрузка данных из CSV
    sheet.clear()  # Очистка листа перед загрузкой новых данных
    sheet.update(range_name='A1', values=csv_data)  # Загрузка всех данных начиная с ячейки A1
    print('записано')

def main():
    gs = get_spreadsheet()
    write_to_sheet(sheet=get_sheet(gs, worksheet_name='Vkusvill_detailed_data'), file_location=DATA_LOCATION)
    write_to_sheet(sheet=get_sheet(gs, worksheet_name='Vkusvill_check_info'), file_location=CHECK_INFO_LOCATION)

if __name__=='__main__':
    main()