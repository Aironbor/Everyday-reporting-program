from setuptools import setup
import platform
from glob import glob


SETUP_DICT = {

    'name': 'Программа для управления отчетами',
    'version': '2.0',
    'description': 'Программа для управления отчетами',
    'author': 'Ivan Metliaev',
    'author_email': 'ivan.metliaev.helper@gmail.com',

    'data_files': (
        ('', glob(r'C:\Windows\SYSTEM32\msvcp100.dll')),
        ('', glob(r'C:\Windows\SYSTEM32\msvcr100.dll')),
        ('platforms', glob(r'C:\Program Files (x86)\Python38-32\Lib\site-packages\PyQt5\Qt\plugins\platforms\qwindows.dll')),
        ('images', ['images/report.png']),
        ('sqldrivers', glob('C:\Program Files (x86)\Python38-32\Lib\site-packages\PyQt5\Qt\plugins\sqldrivers\qsqlite.dll')),
    ),
    'windows': [{'script': 'main_script.py'}],
    'options': {
        'py2exe': {
            'includes': ["lxml._elementpath","PyQt5.QtCore", "PyQt5.QtGui","PyQt5.QtWidgets","db_connect", "config", "images_store", "work_with_excel"],
        },
    }
}

if platform.system() == 'Windows':
    import py2exe
    SETUP_DICT['windows'] = [{
        'Name': 'Ivan Metliaev',
        'product_name': 'Программа для управления отчетами',
        'version': '3.1',
        'description': 'Программа cоздана Метляевым Иваном специально для ООО "Тентовые Конструкции"',
        'copyright': '© 2022, ivan.metliaev.helper@gmail.com. All Rights Reserved',
        'script': 'main_script.py',
        'icon_resources': [(0, r'report_ico.ico')]
    }]
    SETUP_DICT['zipfile'] = None


setup(**SETUP_DICT)
