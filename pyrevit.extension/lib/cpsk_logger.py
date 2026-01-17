# -*- coding: utf-8 -*-
"""
CPSK Logger - Модуль логирования для CPSK Tools.

Совместим с IronPython 2.7 (pyRevit).
Логи сохраняются в папку logs корня проекта.

Использование:
    from cpsk_logger import Logger

    # В начале скрипта - очистить лог и записать заголовок
    Logger.init("FOPtoProject")

    Logger.info("FOPtoProject", "Загружен IDS файл")
    Logger.warning("FOPtoProject", "Параметр не найден в ФОП")
    Logger.error("FOPtoProject", "Ошибка парсинга")

В лог автоматически добавляется номер строки откуда был вызван логгер.
"""

import os
import sys
import codecs
import traceback
from datetime import datetime

# Python 2/3 совместимость
try:
    unicode
except NameError:
    unicode = str

# Определяем пути
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.dirname(_THIS_DIR)
PROJECT_DIR = os.path.dirname(EXTENSION_DIR)  # Корень проекта
LOG_DIR = os.path.join(PROJECT_DIR, "logs")
DEFAULT_LOG_FILE = os.path.join(LOG_DIR, "cpsk.log")


class Logger:
    """
    Простой логгер для CPSK Tools, совместимый с IronPython 2.7.

    Использует файловое логирование без зависимости от модуля logging,
    который может работать некорректно в IronPython.
    """

    _log_file = DEFAULT_LOG_FILE
    _level = "DEBUG"
    _levels = {"DEBUG": 0, "INFO": 1, "WARNING": 2, "ERROR": 3, "CRITICAL": 4}
    _max_file_size = 5 * 1024 * 1024  # 5 MB
    _initialized = False

    @classmethod
    def init(cls, name):
        """
        Инициализировать логгер для скрипта.
        Очищает лог и записывает заголовок сессии.
        Вызывать в начале каждого скрипта.

        :param name: Имя скрипта/модуля
        """
        cls._ensure_log_dir()
        cls.clear()
        cls._initialized = True

        # Записываем заголовок сессии
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        header = [
            u"=" * 70,
            u"СКРИПТ: {}".format(name),
            u"ЗАПУСК: {}".format(timestamp),
            u"=" * 70,
            u""
        ]
        try:
            with codecs.open(cls._log_file, 'a', 'utf-8') as f:
                f.write(u"\n".join(header) + u"\n")
        except (IOError, OSError):
            pass

    @classmethod
    def _get_caller_info(cls, stack_level=3):
        """
        Получить информацию о вызывающем коде (файл и номер строки).

        :param stack_level: Уровень стека (3 = вызов из info/debug/etc)
        :return: Строка вида "filename.py:123"
        """
        try:
            frame = sys._getframe(stack_level)
            filename = os.path.basename(frame.f_code.co_filename)
            lineno = frame.f_lineno
            return u"{}:{}".format(filename, lineno)
        except:
            return u"?:?"

    @classmethod
    def configure(cls, log_file=None, level="INFO"):
        """
        Настроить логгер.

        :param log_file: Путь к файлу лога (по умолчанию logs/cpsk.log)
        :param level: Уровень логирования (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        """
        if log_file:
            cls._log_file = os.path.abspath(log_file)
        if level in cls._levels:
            cls._level = level

    @classmethod
    def _ensure_log_dir(cls):
        """Создать директорию для логов если не существует."""
        log_dir = os.path.dirname(cls._log_file)
        if log_dir and not os.path.exists(log_dir):
            try:
                os.makedirs(log_dir)
            except OSError:
                pass  # Директория уже существует или нет прав

    @classmethod
    def _rotate_if_needed(cls):
        """Ротация лог-файла если превышен размер."""
        try:
            if os.path.exists(cls._log_file):
                size = os.path.getsize(cls._log_file)
                if size > cls._max_file_size:
                    # Переименовываем старый файл
                    backup = cls._log_file + ".old"
                    if os.path.exists(backup):
                        os.remove(backup)
                    os.rename(cls._log_file, backup)
        except (OSError, IOError):
            pass

    @classmethod
    def _write(cls, level, name, message, stack_level=3):
        """
        Записать сообщение в лог.

        :param level: Уровень (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        :param name: Имя модуля/команды
        :param message: Текст сообщения
        :param stack_level: Уровень стека для определения вызывающего кода
        """
        # Проверка уровня
        if cls._levels.get(level, 0) < cls._levels.get(cls._level, 0):
            return

        cls._ensure_log_dir()
        cls._rotate_if_needed()

        # Получаем информацию о вызывающем коде
        caller = cls._get_caller_info(stack_level)

        # Форматируем запись
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]  # Короткий формат с миллисекундами
        # Формат: 10:30:45.123 | INFO     | FOPtoProject | script.py:123 | Сообщение
        log_entry = u"{} | {:<8} | {:<15} | {:<20} | {}\n".format(
            timestamp, level, name, caller, message
        )

        try:
            with codecs.open(cls._log_file, 'a', 'utf-8') as f:
                f.write(log_entry)
        except (IOError, OSError):
            pass  # Не можем записать - пропускаем молча

    @classmethod
    def debug(cls, name, message):
        """Записать DEBUG сообщение."""
        cls._write("DEBUG", name, message)

    @classmethod
    def info(cls, name, message):
        """Записать INFO сообщение."""
        cls._write("INFO", name, message)

    @classmethod
    def warning(cls, name, message):
        """Записать WARNING сообщение."""
        cls._write("WARNING", name, message)

    @classmethod
    def error(cls, name, message, exc_info=False):
        """
        Записать ERROR сообщение.

        :param exc_info: Если True, добавить traceback текущего исключения
        """
        cls._write("ERROR", name, message)
        if exc_info:
            cls._log_exception(name)

    @classmethod
    def critical(cls, name, message, exc_info=True):
        """
        Записать CRITICAL сообщение.
        По умолчанию добавляет traceback.
        """
        cls._write("CRITICAL", name, message)
        if exc_info:
            cls._log_exception(name)

    @classmethod
    def _log_exception(cls, name):
        """Записать traceback текущего исключения."""
        try:
            exc_type, exc_value, exc_tb = sys.exc_info()
            if exc_type:
                tb_lines = traceback.format_exception(exc_type, exc_value, exc_tb)
                for line in tb_lines:
                    for subline in line.rstrip().split('\n'):
                        cls._write("ERROR", name, u"  " + subline, stack_level=4)
        except:
            pass

    @classmethod
    def exception(cls, name, message):
        """Записать ERROR с полным traceback (аналог logging.exception)."""
        cls._write("ERROR", name, message)
        cls._log_exception(name)

    @classmethod
    def clear(cls):
        """Очистить лог-файл."""
        cls._ensure_log_dir()
        try:
            with codecs.open(cls._log_file, 'w', 'utf-8') as f:
                f.write(u"")
        except (IOError, OSError):
            pass

    @classmethod
    def get_log_path(cls):
        """Получить путь к файлу лога."""
        return cls._log_file

    @classmethod
    def log_separator(cls, name, title=""):
        """Записать разделитель в лог (для визуального разделения секций)."""
        separator = u"-" * 50
        if title:
            cls._write("INFO", name, separator, stack_level=4)
            cls._write("INFO", name, title, stack_level=4)
            cls._write("INFO", name, separator, stack_level=4)
        else:
            cls._write("INFO", name, separator, stack_level=4)

    @classmethod
    def file_opened(cls, name, file_path, description=""):
        """Логировать открытие файла с информацией о нём."""
        msg = u"ФАЙЛ ОТКРЫТ: {}".format(file_path)
        if description:
            msg += u" ({})".format(description)
        cls._write("INFO", name, msg)

        # Дополнительная информация о файле
        if os.path.exists(file_path):
            size = os.path.getsize(file_path)
            cls._write("DEBUG", name, u"  Размер: {} байт".format(size))
        else:
            cls._write("WARNING", name, u"  Файл не существует!")

    @classmethod
    def file_saved(cls, name, file_path, description=""):
        """Логировать сохранение файла."""
        msg = u"ФАЙЛ СОХРАНЁН: {}".format(file_path)
        if description:
            msg += u" ({})".format(description)
        cls._write("INFO", name, msg)

    @classmethod
    def data(cls, name, label, data, max_items=20):
        """
        Логировать структуру данных для отладки.

        :param name: Имя модуля
        :param label: Описание данных
        :param data: Данные (dict, list, или любой объект)
        :param max_items: Максимум элементов для вывода
        """
        cls._write("DEBUG", name, u"ДАННЫЕ [{}]:".format(label))

        if isinstance(data, dict):
            cls._write("DEBUG", name, u"  Тип: dict, Кол-во: {}".format(len(data)))
            for i, (k, v) in enumerate(data.items()):
                if i >= max_items:
                    cls._write("DEBUG", name, u"  ... и ещё {}".format(len(data) - max_items))
                    break
                v_str = unicode(v) if not isinstance(v, (str, unicode)) else v
                if len(v_str) > 80:
                    v_str = v_str[:80] + u"..."
                cls._write("DEBUG", name, u"  [{}] = {}".format(k, v_str))

        elif isinstance(data, (list, tuple)):
            cls._write("DEBUG", name, u"  Тип: {}, Кол-во: {}".format(type(data).__name__, len(data)))
            for i, item in enumerate(data):
                if i >= max_items:
                    cls._write("DEBUG", name, u"  ... и ещё {}".format(len(data) - max_items))
                    break
                item_str = unicode(item) if not isinstance(item, (str, unicode)) else item
                if len(item_str) > 80:
                    item_str = item_str[:80] + u"..."
                cls._write("DEBUG", name, u"  [{}] {}".format(i, item_str))

        else:
            cls._write("DEBUG", name, u"  Тип: {}".format(type(data).__name__))
            cls._write("DEBUG", name, u"  Значение: {}".format(data))

    @classmethod
    def result(cls, name, success, message, details=None):
        """Логировать результат операции."""
        level = "INFO" if success else "ERROR"
        status = u"УСПЕХ" if success else u"ОШИБКА"
        cls._write(level, name, u"РЕЗУЛЬТАТ [{}]: {}".format(status, message))
        if details:
            if isinstance(details, (list, tuple)):
                for d in details:
                    cls._write(level, name, u"  - {}".format(d))
            else:
                cls._write(level, name, u"  {}".format(details))


# Для удобства - алиасы на уровне модуля
def debug(name, message):
    """Записать DEBUG сообщение."""
    Logger.debug(name, message)

def info(name, message):
    """Записать INFO сообщение."""
    Logger.info(name, message)

def warning(name, message):
    """Записать WARNING сообщение."""
    Logger.warning(name, message)

def error(name, message):
    """Записать ERROR сообщение."""
    Logger.error(name, message)

def critical(name, message):
    """Записать CRITICAL сообщение."""
    Logger.critical(name, message)
