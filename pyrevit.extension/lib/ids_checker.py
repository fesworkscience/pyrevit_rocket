# -*- coding: utf-8 -*-
"""
IDS Checker - внешний скрипт для проверки IFC по IDS
Запускается через CPython (не IronPython!)

Использование:
    python ids_checker.py <ids_file> <ifc_file> <report_file>

Требования:
    pip install ifcopenshell ifctester
"""

import sys
import os
import json
import traceback
import ctypes
from ctypes import wintypes
import xml.etree.ElementTree as ET
import re
import codecs

# Windows API для проверки существования файлов
# (os.path.exists() иногда не работает с определёнными путями)
INVALID_FILE_ATTRIBUTES = 0xFFFFFFFF
GetFileAttributesW = ctypes.windll.kernel32.GetFileAttributesW
GetFileAttributesW.argtypes = [wintypes.LPCWSTR]
GetFileAttributesW.restype = wintypes.DWORD


def win_path_exists(path):
    """Проверить существование пути через Windows API."""
    attrs = GetFileAttributesW(path)
    return attrs != INVALID_FILE_ATTRIBUTES


# === ЛОГИРОВАНИЕ ===
LOG_FILE = os.path.join(os.path.dirname(__file__), "debug.log")


def log(msg):
    """Записать сообщение в лог-файл."""
    try:
        with codecs.open(LOG_FILE, 'a', 'utf-8') as f:
            f.write(msg + "\n")
    except:
        pass


def clear_log():
    """Очистить лог-файл."""
    try:
        with codecs.open(LOG_FILE, 'w', 'utf-8') as f:
            f.write("=== IDS Checker Log ===\n")
    except:
        pass


def fix_entity_restrictions(specs):
    """
    Исправление бага ifctester 0.8.4 с xs:restriction/xs:enumeration.

    Проблема: ifctester некорректно обрабатывает Restriction объекты созданные
    парсером при загрузке IDS файла - валидация возвращает 0 applicable entities.

    Решение: Пересоздаём Entity с новыми Restriction объектами, сохраняя
    все значения enumeration. Это соответствует рекомендации из GitHub issue #4661:
    restriction = ifctester.ids.Restriction(options={"enumeration": [...]})
    spec.applicability.append(ifctester.ids.Entity(name=restriction))

    GitHub issue: https://github.com/IfcOpenShell/IfcOpenShell/issues/4661
    """
    log("fix_entity_restrictions: START")
    try:
        from ifctester import ids as ids_module
        from ifctester.facet import Restriction
    except ImportError:
        log("fix_entity_restrictions: ImportError")
        return

    fixed_count = 0
    for spec in specs.specifications:
        if not hasattr(spec, 'applicability') or not spec.applicability:
            continue

        new_applicability = []
        for facet in spec.applicability:
            # Проверяем нужно ли пересоздавать Entity
            if type(facet).__name__ != 'Entity':
                new_applicability.append(facet)
                continue

            name_val = getattr(facet, 'name', None)
            pred_val = getattr(facet, 'predefinedType', None)

            needs_fix = (isinstance(name_val, Restriction) or
                         isinstance(pred_val, Restriction))

            if not needs_fix:
                new_applicability.append(facet)
                continue

            # Пересоздаём Entity с новыми Restriction объектами
            new_name = name_val
            new_pred = pred_val

            if isinstance(name_val, Restriction):
                options = getattr(name_val, 'options', {})
                enum_values = options.get('enumeration', [])
                if enum_values:
                    # Исключаем *TYPE классы - они не имеют PropertySets как instances
                    filtered_names = [v for v in enum_values if not v.endswith('TYPE')]
                    if filtered_names:
                        new_name = ids_module.Restriction(options={'enumeration': filtered_names})
                    elif enum_values:
                        # Если только TYPE классы - оставляем как есть
                        new_name = ids_module.Restriction(options={'enumeration': enum_values})

            if isinstance(pred_val, Restriction):
                options = getattr(pred_val, 'options', {})
                enum_values = options.get('enumeration', [])
                if enum_values:
                    # Исключаем USERDEFINED и NOTDEFINED - они ломают валидацию ifctester
                    filtered_values = [v for v in enum_values
                                       if v not in ('USERDEFINED', 'NOTDEFINED')]
                    if filtered_values:
                        new_pred = ids_module.Restriction(options={'enumeration': filtered_values})
                    else:
                        new_pred = None  # Если остались только USERDEFINED/NOTDEFINED

            # Создаём новый Entity
            new_entity = ids_module.Entity(name=new_name, predefinedType=new_pred)
            new_applicability.append(new_entity)
            fixed_count += 1

        spec.applicability = new_applicability

    log("fix_entity_restrictions: DONE, fixed {} entities".format(fixed_count))


def format_restriction(value):
    """Форматировать Restriction объект в читаемый текст."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    # Если это Restriction с enumeration
    if hasattr(value, 'options'):
        options = value.options
        if 'enumeration' in options:
            values = options['enumeration']
            if len(values) == 1:
                return values[0]
            return ", ".join(values)
    # Если это dict
    if isinstance(value, dict) and 'enumeration' in value:
        values = value['enumeration']
        if len(values) == 1:
            return values[0]
        return ", ".join(values)
    return str(value)


def generate_html_report(specs, report_path):
    """
    Генерировать HTML отчет о результатах IDS проверки.
    Свой reporter без багов ifctester.
    """
    html_parts = []

    # Header
    html_parts.append("""<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>IDS Validation Report</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; }
        h1 { color: #333; border-bottom: 2px solid #007bff; padding-bottom: 10px; }
        .summary { background: #fff; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .summary-stats { display: flex; gap: 20px; flex-wrap: wrap; }
        .stat { padding: 15px 25px; border-radius: 6px; text-align: center; }
        .stat-passed { background: #d4edda; color: #155724; }
        .stat-failed { background: #f8d7da; color: #721c24; }
        .stat-total { background: #e2e3e5; color: #383d41; }
        .stat-value { font-size: 24px; font-weight: bold; }
        .stat-label { font-size: 12px; text-transform: uppercase; }
        .spec { background: #fff; margin-bottom: 15px; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .spec-header { padding: 15px 20px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; }
        .spec-header.passed { background: #d4edda; }
        .spec-header.failed { background: #f8d7da; }
        .spec-header.na { background: #fff3cd; }
        .spec-title { font-weight: 600; flex: 1; }
        .spec-status { padding: 4px 12px; border-radius: 4px; font-size: 12px; font-weight: bold; }
        .status-pass { background: #28a745; color: white; }
        .status-fail { background: #dc3545; color: white; }
        .status-na { background: #ffc107; color: #333; }
        .spec-body { padding: 20px; border-top: 1px solid #eee; display: none; }
        .spec.open .spec-body { display: block; }
        .applicability { background: #f8f9fa; padding: 10px 15px; border-radius: 4px; margin-bottom: 15px; font-size: 14px; }
        .applicability-label { font-weight: 600; color: #666; margin-bottom: 5px; }
        .req-table { width: 100%; border-collapse: collapse; }
        .req-table th, .req-table td { padding: 10px; text-align: left; border-bottom: 1px solid #eee; }
        .req-table th { background: #f8f9fa; font-weight: 600; }
        .req-status { width: 80px; text-align: center; }
        .req-pass { color: #28a745; }
        .req-fail { color: #dc3545; }
        .toggle-all { margin-bottom: 15px; }
        .toggle-all button { padding: 8px 16px; margin-right: 10px; border: 1px solid #ddd; background: #fff; border-radius: 4px; cursor: pointer; }
        .toggle-all button:hover { background: #f0f0f0; }
        .elements-info { font-size: 13px; color: #666; margin-top: 5px; }
    </style>
</head>
<body>
<div class="container">
    <h1>Отчёт проверки IDS</h1>
""")

    # Summary
    total_specs = len(specs.specifications) if hasattr(specs, 'specifications') else 0
    passed_specs = sum(1 for s in specs.specifications if getattr(s, 'status', None) is True)
    failed_specs = total_specs - passed_specs

    total_reqs = 0
    passed_reqs = 0
    for spec in specs.specifications:
        if hasattr(spec, 'requirements') and spec.requirements:
            for req in spec.requirements:
                total_reqs += 1
                if getattr(req, 'status', None) is True:
                    passed_reqs += 1

    html_parts.append("""
    <div class="summary">
        <div class="summary-stats">
            <div class="stat stat-total">
                <div class="stat-value">{}</div>
                <div class="stat-label">Спецификаций</div>
            </div>
            <div class="stat stat-passed">
                <div class="stat-value">{}</div>
                <div class="stat-label">Пройдено</div>
            </div>
            <div class="stat stat-failed">
                <div class="stat-value">{}</div>
                <div class="stat-label">Не пройдено</div>
            </div>
            <div class="stat stat-total">
                <div class="stat-value">{}/{}</div>
                <div class="stat-label">Требований</div>
            </div>
        </div>
    </div>
    <div class="toggle-all">
        <button onclick="toggleAll(true)">Развернуть все</button>
        <button onclick="toggleAll(false)">Свернуть все</button>
        <button onclick="showFailed()">Только ошибки</button>
    </div>
""".format(total_specs, passed_specs, failed_specs, passed_reqs, total_reqs))

    # Specifications
    for i, spec in enumerate(specs.specifications):
        spec_name = getattr(spec, 'name', 'Specification {}'.format(i + 1))
        spec_status = getattr(spec, 'status', None)
        applicable = spec.applicable_entities if hasattr(spec, 'applicable_entities') else []
        applicable_count = len(applicable)

        if spec_status is True:
            status_class = "passed"
            status_text = "PASS"
            status_badge = "status-pass"
        elif applicable_count == 0:
            status_class = "na"
            status_text = "N/A"
            status_badge = "status-na"
        else:
            status_class = "failed"
            status_text = "FAIL"
            status_badge = "status-fail"

        # Applicability text
        applicability_parts = []
        if hasattr(spec, 'applicability') and spec.applicability:
            for facet in spec.applicability:
                if hasattr(facet, 'name'):
                    name_str = format_restriction(facet.name)
                    pred_str = format_restriction(getattr(facet, 'predefinedType', None))
                    if pred_str:
                        applicability_parts.append("{} ({})".format(name_str, pred_str))
                    else:
                        applicability_parts.append(name_str)

        applicability_text = ", ".join(applicability_parts) if applicability_parts else "Все элементы"

        html_parts.append("""
    <div class="spec" data-status="{}">
        <div class="spec-header {}" onclick="this.parentElement.classList.toggle('open')">
            <span class="spec-title">{}</span>
            <span class="spec-status {}">{}  ({} элементов)</span>
        </div>
        <div class="spec-body">
            <div class="applicability">
                <div class="applicability-label">Применимость:</div>
                {}
            </div>
""".format(status_class, status_class, spec_name, status_badge, status_text, applicable_count, applicability_text))

        # Requirements table
        if hasattr(spec, 'requirements') and spec.requirements:
            html_parts.append("""
            <table class="req-table">
                <thead>
                    <tr>
                        <th class="req-status">Статус</th>
                        <th>Требование</th>
                        <th>PropertySet</th>
                    </tr>
                </thead>
                <tbody>
""")
            for req in spec.requirements:
                req_status = getattr(req, 'status', None)
                req_class = "req-pass" if req_status is True else "req-fail"
                req_icon = "✓" if req_status is True else "✗"

                # Property name
                base_name = getattr(req, 'baseName', None)
                prop_name = format_restriction(base_name) if base_name else "—"

                # PropertySet name
                pset = getattr(req, 'propertySet', None)
                pset_name = format_restriction(pset) if pset else "—"

                html_parts.append("""
                    <tr>
                        <td class="req-status {}"><strong>{}</strong></td>
                        <td>{}</td>
                        <td>{}</td>
                    </tr>
""".format(req_class, req_icon, prop_name, pset_name))

            html_parts.append("""
                </tbody>
            </table>
""")

        html_parts.append("""
        </div>
    </div>
""")

    # Footer with JS
    html_parts.append("""
</div>
<script>
function toggleAll(open) {
    document.querySelectorAll('.spec').forEach(el => {
        if (open) el.classList.add('open');
        else el.classList.remove('open');
    });
}
function showFailed() {
    document.querySelectorAll('.spec').forEach(el => {
        if (el.dataset.status === 'failed' || el.dataset.status === 'na') {
            el.classList.add('open');
        } else {
            el.classList.remove('open');
        }
    });
}
</script>
</body>
</html>
""")

    # Write file
    with codecs.open(report_path, 'w', 'utf-8') as f:
        f.write("".join(html_parts))


def check_dependencies():
    """Проверить наличие зависимостей."""
    missing = []
    try:
        import ifcopenshell
    except ImportError:
        missing.append("ifcopenshell")

    try:
        from ifctester import ids, reporter
    except ImportError:
        missing.append("ifctester")

    return missing


def validate_ids(ids_path, ifc_path, report_path):
    """
    Проверить IFC модель на соответствие IDS требованиям.

    Args:
        ids_path: Путь к IDS файлу (.ids)
        ifc_path: Путь к IFC файлу (.ifc)
        report_path: Путь для сохранения HTML отчета

    Returns:
        dict: Результат проверки
    """
    clear_log()
    log("validate_ids: START")
    log("  ids_path: {}".format(ids_path))
    log("  ifc_path: {}".format(ifc_path))
    log("  report_path: {}".format(report_path))

    import ifcopenshell
    from ifctester import ids, reporter
    import shutil
    import tempfile

    result = {
        "success": False,
        "ids_file": ids_path,
        "ifc_file": ifc_path,
        "report_file": report_path,
        "total_specs": 0,
        "passed_specs": 0,
        "failed_specs": 0,
        "total_requirements": 0,
        "passed_requirements": 0,
        "failed_requirements": 0,
        "errors": []
    }

    # Копируем IDS во временный файл с уникальным ASCII именем
    # (ifctester/ifcopenshell кэшируют данные по имени файла)
    import uuid
    unique_id = str(uuid.uuid4())[:8]
    temp_ids_path = None
    temp_dir = tempfile.gettempdir()
    try:
        temp_ids_path = os.path.join(temp_dir, "ids_check_{}.ids".format(unique_id))
        shutil.copy2(ids_path, temp_ids_path)
    except Exception as e:
        result["errors"].append("Ошибка копирования IDS: {}".format(str(e)))
        return result

    # Загрузить IDS
    try:
        specs = ids.open(temp_ids_path)
        result["total_specs"] = len(specs.specifications) if hasattr(specs, 'specifications') else 0

        # Исправляем баг ifctester с Restriction в Entity (GitHub issue #4661)
        # DEBUG: проверим до и после
        spec10 = specs.specifications[10] if len(specs.specifications) > 10 else None
        if spec10:
            before_name = str(spec10.applicability[0].name) if spec10.applicability else "N/A"

        fix_entity_restrictions(specs)

        if spec10:
            after_name = str(spec10.applicability[0].name) if spec10.applicability else "N/A"
            result["debug_before"] = before_name
            result["debug_after"] = after_name
    except Exception as e:
        result["errors"].append("Ошибка загрузки IDS: {}".format(str(e)))
        return result

    # Копируем IFC во временный файл с уникальным ASCII именем
    # (ifcopenshell кэширует данные по имени файла)
    temp_ifc_path = None
    try:
        temp_ifc_path = os.path.join(temp_dir, "ifc_check_{}.ifc".format(unique_id))
        shutil.copy2(ifc_path, temp_ifc_path)
    except Exception as e:
        result["errors"].append("Ошибка копирования IFC: {}".format(str(e)))
        return result

    # Открыть IFC
    try:
        ifc_file = ifcopenshell.open(temp_ifc_path)
    except Exception as e:
        result["errors"].append("Ошибка открытия IFC: {}".format(str(e)))
        return result

    # Валидация
    log("validate_ids: Running specs.validate()")
    try:
        specs.validate(ifc_file)
    except Exception as e:
        log("validate_ids: ERROR - {}".format(str(e)))
        result["errors"].append("Ошибка валидации: {}".format(str(e)))
        return result

    # Логируем результаты spec #10 (фундаменты)
    if len(specs.specifications) > 10:
        spec10 = specs.specifications[10]
        log("Spec #10 (Фундаменты):")
        log("  status: {}".format(spec10.status))
        log("  applicable_entities: {}".format(len(spec10.applicable_entities) if spec10.applicable_entities else 0))
        if hasattr(spec10, 'requirements') and spec10.requirements:
            for i, req in enumerate(spec10.requirements[:5]):
                log("  Req #{} {}: {}".format(i, getattr(req, 'baseName', '?'), getattr(req, 'status', None)))

    # Подсчет результатов
    passed_specs = 0
    failed_specs = 0
    total_reqs = 0
    passed_reqs = 0
    failed_reqs = 0

    if hasattr(specs, 'specifications'):
        for spec in specs.specifications:
            # Используем статус спецификации напрямую
            spec_status = getattr(spec, 'status', None)
            if spec_status is True:
                passed_specs += 1
            else:
                failed_specs += 1

            # Подсчёт требований
            if hasattr(spec, 'requirements') and spec.requirements:
                for req in spec.requirements:
                    total_reqs += 1
                    req_status = getattr(req, 'status', None)
                    if req_status is True:
                        passed_reqs += 1
                    else:
                        failed_reqs += 1

    result["passed_specs"] = passed_specs
    result["failed_specs"] = failed_specs
    result["total_requirements"] = total_reqs
    result["passed_requirements"] = passed_reqs
    result["failed_requirements"] = failed_reqs

    # Создать HTML отчет (свой, без багов ifctester reporter)
    try:
        generate_html_report(specs, report_path)
        result["success"] = True
    except Exception as e:
        result["errors"].append("Ошибка создания отчета: {}".format(str(e)))

    # Удалить временные файлы
    for temp_path in [temp_ids_path, temp_ifc_path]:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

    return result


def validate_ids_standard(ids_path, ifc_path, report_path):
    """
    Валидация IDS СТАНДАРТНЫМ ifctester БЕЗ исправлений.
    Генерирует отчёт стандартным reporter.Html для сравнения.
    """
    log("validate_ids_standard: START (without fixes)")

    import ifcopenshell
    from ifctester import ids, reporter
    import shutil
    import tempfile
    import uuid

    result = {"success": False, "errors": []}

    # Копируем файлы во временные
    unique_id = str(uuid.uuid4())[:8]
    temp_dir = tempfile.gettempdir()

    try:
        temp_ids_path = os.path.join(temp_dir, "std_ids_{}.ids".format(unique_id))
        temp_ifc_path = os.path.join(temp_dir, "std_ifc_{}.ifc".format(unique_id))
        shutil.copy2(ids_path, temp_ids_path)
        shutil.copy2(ifc_path, temp_ifc_path)
    except Exception as e:
        result["errors"].append("Ошибка копирования: {}".format(str(e)))
        return result

    try:
        # Загружаем БЕЗ fix_entity_restrictions!
        specs = ids.open(temp_ids_path)
        log("validate_ids_standard: IDS loaded, {} specs".format(len(specs.specifications)))

        ifc_file = ifcopenshell.open(temp_ifc_path)
        log("validate_ids_standard: IFC loaded")

        # Валидация БЕЗ исправлений
        specs.validate(ifc_file)
        log("validate_ids_standard: Validation done")

        # Генерируем отчёт СТАНДАРТНЫМ reporter.Html
        html_reporter = reporter.Html(specs)
        html_reporter.report()
        html_reporter.to_file(report_path)
        log("validate_ids_standard: Standard HTML report saved to {}".format(report_path))

        result["success"] = True

    except Exception as e:
        log("validate_ids_standard: ERROR - {}".format(str(e)))
        result["errors"].append("Ошибка: {}".format(str(e)))

    # Удалить временные файлы
    for temp_path in [temp_ids_path, temp_ifc_path]:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

    return result


def save_result(result, result_file=None):
    """Сохранить результат в файл и вывести в stdout."""
    result_json = json.dumps(result, ensure_ascii=False)
    print(result_json)

    if result_file:
        with codecs.open(result_file, 'w', 'utf-8') as f:
            f.write(result_json)


def main():
    """Главная функция."""
    # Лог для отладки
    debug_log = os.path.join(os.path.dirname(__file__), "ids_checker_debug.log")

    def log(msg):
        with codecs.open(debug_log, 'a', 'utf-8') as f:
            f.write(msg + "\n")

    log("=" * 50)
    log("ids_checker.py started")
    log("sys.argv: {}".format(sys.argv))

    # Флаг для стандартного отчёта (без исправлений)
    use_standard = "--standard" in sys.argv

    # Режим JSON (для решения проблемы кириллицы в путях)
    if len(sys.argv) >= 4 and sys.argv[1] == "--json":
        params_file = sys.argv[2]
        result_file = sys.argv[3]

        log("params_file: {}".format(params_file))
        log("params_file exists: {}".format(os.path.exists(params_file)))

        try:
            # Читаем JSON с правильной кодировкой
            with codecs.open(params_file, 'r', 'utf-8-sig') as f:
                content = f.read()
                log("params content: {}".format(content))

            params = json.loads(content)

            # Нормализуем пути
            ids_path = os.path.normpath(params["ids_path"])
            ifc_path = os.path.normpath(params["ifc_path"])
            report_path = os.path.normpath(params["report_path"])

            # Опциональный путь для стандартного отчета
            standard_report_path = params.get("standard_report_path")
            if standard_report_path:
                standard_report_path = os.path.normpath(standard_report_path)
                log("standard_report_path: {}".format(standard_report_path))

            log("ids_path: {}".format(ids_path))
            log("ids_path os.path.exists: {}".format(os.path.exists(ids_path)))
            log("ids_path win_path_exists: {}".format(win_path_exists(ids_path)))
            log("ifc_path: {}".format(ifc_path))
            log("ifc_path win_path_exists: {}".format(win_path_exists(ifc_path)))
        except Exception as e:
            log("ERROR reading params: {}".format(str(e)))
            save_result({
                "success": False,
                "errors": ["Ошибка чтения параметров: {}".format(str(e))]
            }, result_file)
            sys.exit(1)

    # Обычный режим (аргументы командной строки)
    elif len(sys.argv) >= 4:
        ids_path = sys.argv[1]
        ifc_path = sys.argv[2]
        report_path = sys.argv[3]
        result_file = None

        # Проверяем флаг --standard для стандартного отчёта
        use_standard = "--standard" in sys.argv
    else:
        print(json.dumps({
            "success": False,
            "errors": ["Использование: python ids_checker.py <ids_file> <ifc_file> <report_file>"]
        }))
        sys.exit(1)

    # Проверить зависимости
    missing = check_dependencies()
    if missing:
        save_result({
            "success": False,
            "errors": ["Отсутствуют библиотеки: {}. Установите: pip install {}".format(
                ", ".join(missing), " ".join(missing)
            )]
        }, result_file)
        sys.exit(1)

    # Проверить файлы (используем Windows API вместо os.path.exists)
    if not win_path_exists(ids_path):
        save_result({
            "success": False,
            "errors": ["IDS файл не найден: {}".format(ids_path)]
        }, result_file)
        sys.exit(1)

    if not win_path_exists(ifc_path):
        save_result({
            "success": False,
            "errors": ["IFC файл не найден: {}".format(ifc_path)]
        }, result_file)
        sys.exit(1)

    # Выполнить проверку
    try:
        # Если флаг --standard - используем стандартный ifctester без исправлений
        if 'use_standard' in dir() and use_standard:
            log("Using STANDARD ifctester (no fixes)")
            result = validate_ids_standard(ids_path, ifc_path, report_path)
        else:
            # CPSK версия с исправлениями
            result = validate_ids(ids_path, ifc_path, report_path)

            # Если запрошен стандартный отчет - генерируем его дополнительно
            if 'standard_report_path' in dir() and standard_report_path:
                log("Also generating STANDARD report to: {}".format(standard_report_path))
                try:
                    validate_ids_standard(ids_path, ifc_path, standard_report_path)
                except Exception as std_err:
                    log("ERROR generating standard report: {}".format(str(std_err)))
                    # Не прерываем - основной отчет уже создан

        save_result(result, result_file)
    except Exception as e:
        save_result({
            "success": False,
            "errors": ["Неожиданная ошибка: {}".format(traceback.format_exc())]
        }, result_file)
        sys.exit(1)


if __name__ == "__main__":
    main()
