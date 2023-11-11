import os
import subprocess
import tempfile
import zipfile
import fnmatch
import glob
import re


def get_files_re(path: str, mask: str, recursive: bool = False) -> list[str]:
    """
    Ищет файлы по маске регулярного выражения в папке path. Может рекурсивно
    :param path:
    :param mask:
    :param recursive:
    :return:
    """
    # Определение параметра glob
    recursive_option = '**' if recursive else ''

    # Поиск всех файлов и папок по заданному пути с указанным рекурсивным параметром
    all_files_and_folders = glob.glob(os.path.join(path, recursive_option, '*'), recursive=recursive)

    # Проверка соответствия маске для каждого файла
    regex = re.compile(mask)
    files = [f for f in all_files_and_folders if os.path.isfile(f) and regex.fullmatch(os.path.basename(f))]

    return files


def get_files(path: str, mask: str, recursive: bool = False) -> list[str]:
    """
    Ищет файлы по маске mask в path. Может искать рекурсивно
    :param path:
    :param mask:
    :param recursive:
    :return:
    """
    matches = []
    for root, dirnames, filenames in os.walk(path):
        if not recursive and root != path:
            continue
        for filename in fnmatch.filter(filenames, mask):
            matches.append(os.path.join(root, filename))
    return matches


def get_download_folder() -> str:
    """
    :return: path to download folder
    """
    # Для Windows
    if os.name == 'nt':
        import ctypes
        from ctypes import wintypes, windll

        CSIDL_DOWNLOADS = 0x19  # индентификатор папки загрузок для Windows

        # Функция для вызова WinAPI и получения пути к папке
        def _get_path(csidl):
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            windll.shell32.SHGetFolderPathW(None, csidl, None, 0, buf)
            return buf.value

        return _get_path(CSIDL_DOWNLOADS)

    # Для macOS, Linux и других Unix-подобных OS
    else:
        home = os.path.expanduser('~')
        download_folder = os.path.join(home, 'Downloads')
        # Проверка существования папки для Unix систем, если пользователь изменил настройки
        if not os.path.isdir(download_folder):
            # Fallback на использование стандартного XDG для Linux
            download_folder = os.getenv('XDG_DOWNLOAD_DIR',
                                        os.path.join(home, 'Downloads'))
        return download_folder


def contains_file(zip_file: str, search_file: str) -> bool:
    """
    Проверяет, содержит ли zip_file search_file
    :param zip_file:
    :param search_file:
    :return:
    """
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        if search_file in zip_ref.namelist():
            return True
    return False


def create_temp_dir():
    """
    Создаёт временный каталог и возвращает его путь
    :return:
    """
    temp_dir = tempfile.mkdtemp()
    return temp_dir


def unzip_file(zip_filepath: str, directory_to_extract_to: str):
    with zipfile.ZipFile(zip_filepath, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)


def get_basename_without_extension(path):
    return os.path.splitext(os.path.basename(path))[0]


def open_file(filename):
    try:
        # для Windows
        if os.name == 'nt':
            os.startfile(filename)

        # для MacOS
        elif os.name == 'mac':
            subprocess.call(('open', filename))

        # для Linux
        elif os.name == 'posix':
            subprocess.call(('xdg-open', filename))
    except:
        print(f'Could not open file {filename}')


def get_max_filename(file_paths):
    # Функция для извлечения имени файла без расширения
    def file_name_without_extension(path):
        return os.path.splitext(os.path.basename(path))[0]

    # Сортируем список файлов по имени файла без учёта расширения
    sorted_files = sorted(file_paths, key=file_name_without_extension)

    # Возвращаем файл с наибольшим именем (последний в отсортированном списке)
    return sorted_files[-1]
