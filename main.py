import os
import shutil

from dotenv import load_dotenv

from custom import reformat_xlsx, transform_excel_column_to_num
from os_func import get_download_folder, get_files_re, contains_file, get_files, unzip_file, create_temp_dir, \
    get_basename_without_extension, open_file, get_max_filename
from xls_func import import_csv_to_xlsx, get_list_name_from_file_name, replace_values_from_another_sheet

load_dotenv()

ACCOUNT_NUMBER = os.getenv("ACCOUNT_NUMBER")
OUTPUT_FILE = os.path.join(os.path.expanduser('~'), os.getenv("OUTPUT_FILE"))

if __name__ == '__main__':
    path = get_download_folder()
    files = get_files(path, f'{ACCOUNT_NUMBER}*.zip')
    # files = get_files_re(path, r'^\d{6,7}\.zip$')
    last_file_name = get_max_filename(files)
    tmp_dir = create_temp_dir()
    if not contains_file(last_file_name, f'{ACCOUNT_NUMBER}_exchange.txt'):
        print(f'File {last_file_name} not contains txt file needed')
        exit(0)
    unzip_file(last_file_name, tmp_dir)
    txt_files = get_files(tmp_dir, f'{ACCOUNT_NUMBER}_*.txt')
    xlsx_file = f'{tmp_dir}/test.xlsx'
    for txt_file in txt_files:
        sheet_name = get_list_name_from_file_name(get_basename_without_extension(txt_file))
        import_csv_to_xlsx(txt_file, xlsx_file, sheet_name)
    print(tmp_dir)
    replace_values_from_another_sheet(xlsx_file, 'expense', 'Счёт', 'account_category', 'id', 'Название')
    replace_values_from_another_sheet(xlsx_file, 'expense', 'Валюта', 'currency', 'id', 'Название')
    replace_values_from_another_sheet(xlsx_file, 'expense', '№ Категории', 'expense_category', 'id', 'Название')
    transform_excel_column_to_num(xlsx_file, 'expense', 'Сумма')
    # set_auto_filter(xlsx_file, 'expense')
    reformat_xlsx(xlsx_file, OUTPUT_FILE)
    shutil.rmtree(tmp_dir)
    open_file(OUTPUT_FILE)

