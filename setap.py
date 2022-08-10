from setuptools import setup
import platform
from glob import glob


from setuptools import setup
import platform
from glob import glob

SETUP_DICT = {

    'name': 'Программа формирования производственных заданий',
    'version': '1.0',
    'description': 'Программа формирования производственных заданий',
    'author': 'Ivan Metliaev',
    'author_email': 'ivan.metliaev.helper@gmail.com',

    'data_files': (
        ('', glob(r'C:\Windows\SYSTEM32\msvcp100.dll')),
        ('', glob(r'C:\Windows\SYSTEM32\msvcr100.dll')),
        ('platforms', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\plugins\platforms\qwindows.dll')),
        ('sqldrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\plugins\sqldrivers\qsqlite.dll')),
        ('qtcoredrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\bin\Qt5Core.dll')),
        ('qtguidrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\bin\Qt5Gui.dll')),
        ('qtwidgetdrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\bin\Qt5Widgets.dll')),
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
        'version': '3.3',
        'description': 'Программа cоздана Метляевым Иваном специально для ООО "Тентовые Конструкции"',
        'copyright': '© 2022, ivan.metliaev.helper@gmail.com. All Rights Reserved',
        'script': 'main_script.py',
        'icon_resources': [(0, r'report_ico.ico')]
    }]
    SETUP_DICT['zipfile'] = None


setup(**SETUP_DICT)
