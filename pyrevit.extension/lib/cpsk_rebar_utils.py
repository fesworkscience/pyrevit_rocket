# -*- coding: utf-8 -*-
"""
Модуль для работы с арматурой и геометрией для Smart Openings.

Функции для:
- Получения геометрии стержней
- Поиска пересечений с отверстиями
- Разрезания и склейки стержней
- Проверки коллинеарности
"""

import math
import os
import codecs

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    XYZ,
    Line,
    Options,
    Transaction,
    ElementId,
    Solid,
    GeometryInstance,
    BoundingBoxXYZ,
    BoundingBoxIntersectsFilter,
    Outline,
    CurveLoop,
    GeometryCreationUtilities,
    SolidOptions,
    SolidUtils,
    Transform
)

from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarStyle,
    RebarHookOrientation
)

# Tolerance для геометрических проверок (в футах, ~3мм)
TOLERANCE = 0.01


def mm_to_feet(mm):
    """Конвертировать миллиметры в футы."""
    return mm / 304.8


def feet_to_mm(feet):
    """Конвертировать футы в миллиметры."""
    return feet * 304.8


def get_rebar_centerline(rebar, return_method=False):
    """
    Получить centerline кривую стержня.

    Args:
        rebar: Rebar element
        return_method: Если True, возвращает tuple (Curve, method_name)

    Returns:
        Curve или None (или tuple если return_method=True)
    """
    # Метод 1: GetCenterlineCurves
    try:
        curves = rebar.GetCenterlineCurves(False, False, False, None, 0)
        if curves and curves.Count > 0:
            if return_method:
                return curves[0], "GetCenterlineCurves(1)"
            return curves[0]
    except Exception:
        pass

    # Метод 2: GetCenterlineCurves с другими параметрами
    try:
        curves = rebar.GetCenterlineCurves(True, True, True, None, 0)
        if curves and curves.Count > 0:
            if return_method:
                return curves[0], "GetCenterlineCurves(2)"
            return curves[0]
    except Exception:
        pass

    # Метод 3: Через геометрию
    try:
        opt = Options()
        opt.ComputeReferences = False
        geom = rebar.get_Geometry(opt)
        if geom:
            for g in geom:
                if hasattr(g, 'GetEndPoint'):
                    if return_method:
                        return g, "Geometry"
                    return g
                if isinstance(g, Line):
                    if return_method:
                        return g, "Geometry"
                    return g
    except Exception:
        pass

    # Метод 4: Создать линию из BoundingBox
    try:
        bbox = rebar.get_BoundingBox(None)
        if bbox:
            dx = bbox.Max.X - bbox.Min.X
            dy = bbox.Max.Y - bbox.Min.Y
            dz = bbox.Max.Z - bbox.Min.Z

            # Найти главную ось
            if dx >= dy and dx >= dz:
                # Вдоль X
                mid_y = (bbox.Min.Y + bbox.Max.Y) / 2
                mid_z = (bbox.Min.Z + bbox.Max.Z) / 2
                line = Line.CreateBound(
                    XYZ(bbox.Min.X, mid_y, mid_z),
                    XYZ(bbox.Max.X, mid_y, mid_z)
                )
                if return_method:
                    return line, "BBox-X"
                return line
            elif dy >= dx and dy >= dz:
                # Вдоль Y
                mid_x = (bbox.Min.X + bbox.Max.X) / 2
                mid_z = (bbox.Min.Z + bbox.Max.Z) / 2
                line = Line.CreateBound(
                    XYZ(mid_x, bbox.Min.Y, mid_z),
                    XYZ(mid_x, bbox.Max.Y, mid_z)
                )
                if return_method:
                    return line, "BBox-Y"
                return line
            else:
                # Вдоль Z
                mid_x = (bbox.Min.X + bbox.Max.X) / 2
                mid_y = (bbox.Min.Y + bbox.Max.Y) / 2
                line = Line.CreateBound(
                    XYZ(mid_x, mid_y, bbox.Min.Z),
                    XYZ(mid_x, mid_y, bbox.Max.Z)
                )
                if return_method:
                    return line, "BBox-Z"
                return line
    except Exception:
        pass

    if return_method:
        return None, "None"
    return None


def get_rebar_endpoints(rebar):
    """
    Получить конечные точки стержня.

    Args:
        rebar: Rebar element

    Returns:
        tuple (start_point, end_point) или (None, None)
    """
    curve = get_rebar_centerline(rebar)
    if curve is None:
        return None, None

    return curve.GetEndPoint(0), curve.GetEndPoint(1)


def get_opening_solid(opening):
    """
    Получить Solid геометрию отверстия.

    Args:
        opening: Opening element

    Returns:
        Solid или None
    """
    opt = Options()
    opt.ComputeReferences = False
    opt.DetailLevel = 1  # Medium

    geom = opening.get_Geometry(opt)
    if geom is None:
        return None

    for geom_obj in geom:
        if isinstance(geom_obj, Solid):
            if geom_obj.Volume > 0:
                return geom_obj
        elif isinstance(geom_obj, GeometryInstance):
            inst_geom = geom_obj.GetInstanceGeometry()
            for inst_obj in inst_geom:
                if isinstance(inst_obj, Solid):
                    if inst_obj.Volume > 0:
                        return inst_obj
    return None


def get_opening_bounding_box(opening):
    """
    Получить BoundingBox отверстия.

    Args:
        opening: Opening element

    Returns:
        BoundingBoxXYZ или None
    """
    return opening.get_BoundingBox(None)


def line_intersects_solid(line, solid):
    """
    Проверить пересекает ли линия Solid и найти точки пересечения.

    Args:
        line: Line (Curve)
        solid: Solid

    Returns:
        list of XYZ points (точки пересечения) или пустой список
    """
    intersections = []

    # Проходим по граням solid
    for face in solid.Faces:
        try:
            result = face.Intersect(line)
            if result is not None:
                # result - это IntersectionResultArray
                for i in range(result.Size):
                    pt = result.get_Item(i).XYZPoint
                    intersections.append(pt)
        except Exception:
            continue

    # Сортировать по расстоянию от начала линии
    if intersections:
        start = line.GetEndPoint(0)
        intersections.sort(key=lambda p: p.DistanceTo(start))

    return intersections


def line_intersects_bbox(line, bbox, buffer_feet=0):
    """
    Проверить пересекает ли линия BoundingBox (с буфером).

    Args:
        line: Line (Curve)
        bbox: BoundingBoxXYZ
        buffer_feet: Буфер в футах

    Returns:
        bool
    """
    if bbox is None:
        return False

    # Расширить bbox на буфер
    min_pt = XYZ(
        bbox.Min.X - buffer_feet,
        bbox.Min.Y - buffer_feet,
        bbox.Min.Z - buffer_feet
    )
    max_pt = XYZ(
        bbox.Max.X + buffer_feet,
        bbox.Max.Y + buffer_feet,
        bbox.Max.Z + buffer_feet
    )

    # Получить концы линии
    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)

    # Простая проверка: если оба конца за пределами bbox с одной стороны - нет пересечения
    if p0.X < min_pt.X and p1.X < min_pt.X:
        return False
    if p0.X > max_pt.X and p1.X > max_pt.X:
        return False
    if p0.Y < min_pt.Y and p1.Y < min_pt.Y:
        return False
    if p0.Y > max_pt.Y and p1.Y > max_pt.Y:
        return False
    if p0.Z < min_pt.Z and p1.Z < min_pt.Z:
        return False
    if p0.Z > max_pt.Z and p1.Z > max_pt.Z:
        return False

    return True


def are_lines_collinear(line1, line2, angle_tolerance=0.01):
    """
    Проверить коллинеарность двух линий.

    Args:
        line1: Line (Curve)
        line2: Line (Curve)
        angle_tolerance: Допуск угла в радианах

    Returns:
        bool
    """
    dir1 = (line1.GetEndPoint(1) - line1.GetEndPoint(0)).Normalize()
    dir2 = (line2.GetEndPoint(1) - line2.GetEndPoint(0)).Normalize()

    # Проверить параллельность (dot product близок к 1 или -1)
    dot = abs(dir1.DotProduct(dir2))
    if dot < 1.0 - angle_tolerance:
        return False

    # Проверить что точки лежат на одной прямой
    # Расстояние от точки line2.start до линии line1
    p = line2.GetEndPoint(0)
    closest = line1.Project(p)
    if closest is None:
        return False

    distance = p.DistanceTo(closest.XYZPoint)
    return distance < TOLERANCE


def find_closest_endpoints(rebar1, rebar2):
    """
    Найти ближайшие концы двух стержней.

    Args:
        rebar1: Rebar element
        rebar2: Rebar element

    Returns:
        tuple (dist, end1_idx, end2_idx) или None
        end_idx: 0 = start, 1 = end
    """
    p1_start, p1_end = get_rebar_endpoints(rebar1)
    p2_start, p2_end = get_rebar_endpoints(rebar2)

    if p1_start is None or p2_start is None:
        return None

    # Все комбинации
    distances = [
        (p1_start.DistanceTo(p2_start), 0, 0),
        (p1_start.DistanceTo(p2_end), 0, 1),
        (p1_end.DistanceTo(p2_start), 1, 0),
        (p1_end.DistanceTo(p2_end), 1, 1),
    ]

    # Найти минимальное расстояние
    distances.sort(key=lambda x: x[0])
    return distances[0]


def can_merge_rebars(rebar1, rebar2, tolerance_feet=None):
    """
    Проверить можно ли склеить два стержня.

    Args:
        rebar1: Rebar element
        rebar2: Rebar element
        tolerance_feet: Допуск расстояния между концами

    Returns:
        tuple (can_merge: bool, info: dict)
    """
    if tolerance_feet is None:
        tolerance_feet = mm_to_feet(50)  # 50mm по умолчанию

    info = {
        "closest_distance": None,
        "end1_idx": None,
        "end2_idx": None,
        "collinear": False
    }

    # Получить centerlines
    line1 = get_rebar_centerline(rebar1)
    line2 = get_rebar_centerline(rebar2)

    if line1 is None or line2 is None:
        return False, info

    # Проверить коллинеарность
    info["collinear"] = are_lines_collinear(line1, line2)
    if not info["collinear"]:
        return False, info

    # Найти ближайшие концы
    closest = find_closest_endpoints(rebar1, rebar2)
    if closest is None:
        return False, info

    dist, end1_idx, end2_idx = closest
    info["closest_distance"] = dist
    info["end1_idx"] = end1_idx
    info["end2_idx"] = end2_idx

    # Проверить расстояние
    if dist > tolerance_feet:
        return False, info

    return True, info


def get_rebars_in_host(doc, host):
    """
    Получить все стержни, размещённые в данном хосте (плите).

    Args:
        doc: Revit Document
        host: Host element (Floor, Wall, etc.)

    Returns:
        list of Rebar elements
    """
    result = []
    host_id = host.Id

    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_Rebar
    ).WhereElementIsNotElementType()

    for rebar in collector:
        try:
            rebar_host_id = rebar.GetHostId()
            if rebar_host_id == host_id:
                result.append(rebar)
        except Exception:
            continue

    return result


def get_openings_in_floor(doc, floor):
    """
    Получить все отверстия в плите.

    Args:
        doc: Revit Document
        floor: Floor element

    Returns:
        list of Opening elements
    """
    result = []
    floor_id = floor.Id

    # Отверстия в перекрытиях
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_FloorOpening
    ).WhereElementIsNotElementType()

    for opening in collector:
        try:
            host = opening.Host
            if host is not None and host.Id == floor_id:
                result.append(opening)
        except Exception:
            continue

    return result


def find_rebars_intersecting_opening(doc, opening, host):
    """
    Найти стержни, пересекающие отверстие.

    Args:
        doc: Revit Document
        opening: Opening element
        host: Host element (плита)

    Returns:
        list of Rebar elements
    """
    result = []

    # Получить bbox отверстия для быстрой фильтрации
    bbox = get_opening_bounding_box(opening)
    if bbox is None:
        return result

    # Получить solid для точной проверки
    solid = get_opening_solid(opening)

    # Получить все стержни в хосте
    rebars = get_rebars_in_host(doc, host)

    for rebar in rebars:
        line = get_rebar_centerline(rebar)
        if line is None:
            continue

        # Быстрая проверка по bbox
        if not line_intersects_bbox(line, bbox, mm_to_feet(100)):
            continue

        # Точная проверка по solid (если есть)
        if solid is not None:
            intersections = line_intersects_solid(line, solid)
            if len(intersections) >= 2:
                result.append(rebar)
        else:
            # Если solid не получен - используем bbox
            result.append(rebar)

    return result


def split_line_by_solid(line, solid, offset_feet):
    """
    Разделить линию на части по границам solid с отступом.

    Args:
        line: Line (Curve)
        solid: Solid
        offset_feet: Отступ от границы solid в футах

    Returns:
        list of Line или None если нет пересечения
    """
    intersections = line_intersects_solid(line, solid)

    if len(intersections) < 2:
        return None

    # Берём первую и последнюю точку пересечения
    entry_pt = intersections[0]
    exit_pt = intersections[-1]

    start_pt = line.GetEndPoint(0)
    end_pt = line.GetEndPoint(1)

    # Направление линии
    direction = (end_pt - start_pt).Normalize()

    # Вычислить новые точки с отступом
    cut_start = entry_pt - direction.Multiply(offset_feet)
    cut_end = exit_pt + direction.Multiply(offset_feet)

    result = []

    # Левая часть: от start до cut_start
    if start_pt.DistanceTo(cut_start) > TOLERANCE:
        try:
            left_line = Line.CreateBound(start_pt, cut_start)
            result.append(("left", left_line))
        except Exception:
            pass

    # Правая часть: от cut_end до end
    if cut_end.DistanceTo(end_pt) > TOLERANCE:
        try:
            right_line = Line.CreateBound(cut_end, end_pt)
            result.append(("right", right_line))
        except Exception:
            pass

    return result


def split_line_by_rect_2d(line, rect_min_x, rect_min_y, rect_max_x, rect_max_y, offset_feet):
    """
    Разделить линию на части по границам прямоугольника в 2D с отступом.

    Использует 2D проверку (игнорирует Z), затем создаёт 3D линии
    на том же Z уровне что и исходная линия.

    Args:
        line: Line (Curve)
        rect_min_x, rect_min_y, rect_max_x, rect_max_y: границы прямоугольника
        offset_feet: Отступ от границы в футах

    Returns:
        list of (name, Line) или None если нет пересечения
    """
    start_pt = line.GetEndPoint(0)
    end_pt = line.GetEndPoint(1)

    x1, y1 = start_pt.X, start_pt.Y
    x2, y2 = end_pt.X, end_pt.Y
    z_level = (start_pt.Z + end_pt.Z) / 2  # Средний Z уровень

    # Проверить пересекает ли линия прямоугольник
    if not line_intersects_rect_2d((x1, y1), (x2, y2), rect_min_x, rect_min_y, rect_max_x, rect_max_y):
        return None

    # Найти точки входа и выхода линии через прямоугольник
    # Используем параметрическое представление линии: P = P1 + t*(P2-P1), t in [0,1]

    dx = x2 - x1
    dy = y2 - y1

    # Параметры t для пересечения с каждой стороной
    t_values = []

    # Пересечение с левой стороной (x = rect_min_x)
    if abs(dx) > 1e-10:
        t = (rect_min_x - x1) / dx
        if 0 <= t <= 1:
            y_at_t = y1 + t * dy
            if rect_min_y <= y_at_t <= rect_max_y:
                t_values.append(t)

    # Пересечение с правой стороной (x = rect_max_x)
    if abs(dx) > 1e-10:
        t = (rect_max_x - x1) / dx
        if 0 <= t <= 1:
            y_at_t = y1 + t * dy
            if rect_min_y <= y_at_t <= rect_max_y:
                t_values.append(t)

    # Пересечение с нижней стороной (y = rect_min_y)
    if abs(dy) > 1e-10:
        t = (rect_min_y - y1) / dy
        if 0 <= t <= 1:
            x_at_t = x1 + t * dx
            if rect_min_x <= x_at_t <= rect_max_x:
                t_values.append(t)

    # Пересечение с верхней стороной (y = rect_max_y)
    if abs(dy) > 1e-10:
        t = (rect_max_y - y1) / dy
        if 0 <= t <= 1:
            x_at_t = x1 + t * dx
            if rect_min_x <= x_at_t <= rect_max_x:
                t_values.append(t)

    # Убрать дубликаты и отсортировать
    t_values = sorted(set(t_values))

    if len(t_values) < 2:
        # Линия касается прямоугольника только в одной точке или не пересекает
        return None

    # Взять первую и последнюю точку пересечения
    t_entry = t_values[0]
    t_exit = t_values[-1]

    # Вычислить точки входа и выхода
    entry_x = x1 + t_entry * dx
    entry_y = y1 + t_entry * dy
    exit_x = x1 + t_exit * dx
    exit_y = y1 + t_exit * dy

    # Направление линии (нормализованное)
    line_length = math.sqrt(dx * dx + dy * dy)
    if line_length < 1e-10:
        return None

    dir_x = dx / line_length
    dir_y = dy / line_length

    # Применить отступ
    cut_start_x = entry_x - dir_x * offset_feet
    cut_start_y = entry_y - dir_y * offset_feet
    cut_end_x = exit_x + dir_x * offset_feet
    cut_end_y = exit_y + dir_y * offset_feet

    result = []

    # Левая часть: от start до cut_start
    dist_to_cut_start = math.sqrt((cut_start_x - x1)**2 + (cut_start_y - y1)**2)
    if dist_to_cut_start > TOLERANCE:
        try:
            left_line = Line.CreateBound(
                XYZ(x1, y1, z_level),
                XYZ(cut_start_x, cut_start_y, z_level)
            )
            result.append(("left", left_line))
        except Exception:
            pass

    # Правая часть: от cut_end до end
    dist_from_cut_end = math.sqrt((x2 - cut_end_x)**2 + (y2 - cut_end_y)**2)
    if dist_from_cut_end > TOLERANCE:
        try:
            right_line = Line.CreateBound(
                XYZ(cut_end_x, cut_end_y, z_level),
                XYZ(x2, y2, z_level)
            )
            result.append(("right", right_line))
        except Exception:
            pass

    return result if result else None


def create_rebar_from_curve(doc, original_rebar, new_curve, host):
    """
    Создать новый Rebar из кривой на основе оригинального.

    Args:
        doc: Revit Document
        original_rebar: Оригинальный Rebar для копирования параметров
        new_curve: Новая Curve для стержня
        host: Host element

    Returns:
        Rebar или None
    """
    try:
        # Получить параметры оригинального стержня
        rebar_type = doc.GetElement(original_rebar.GetTypeId())

        # Нормаль для размещения (обычно Z-up для плит)
        normal = XYZ.BasisZ

        # Создать новый стержень
        # useExistingShapeIfPossible=False чтобы не объединялся с существующими
        new_rebar = Rebar.CreateFromCurves(
            doc,
            RebarStyle.Standard,
            rebar_type,
            None,  # startHook
            None,  # endHook
            host,
            normal,
            [new_curve],
            RebarHookOrientation.Left,
            RebarHookOrientation.Left,
            False,  # useExistingShapeIfPossible - НЕ использовать существующие формы
            True    # createNewShape - создать новую форму
        )

        # ВАЖНО: Установить layout как одиночный стержень через accessor
        # Это предотвращает автоматическое объединение в Rebar Set
        layout_set = False
        layout_error = None
        if new_rebar is not None:
            try:
                accessor = new_rebar.GetShapeDrivenAccessor()
                if accessor is not None:
                    accessor.SetLayoutAsSingle()
                    layout_set = True
                else:
                    layout_error = "GetShapeDrivenAccessor returned None"
            except Exception as ex:
                layout_error = str(ex)

        # Логируем создание стержня
        if new_rebar is not None:
            log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log_rebar.log")
            try:
                with codecs.open(log_path, 'a', 'utf-8') as f:
                    f.write("\n=== Created rebar ID={} from original ID={} ===\n".format(
                        new_rebar.Id.IntegerValue, original_rebar.Id.IntegerValue))

                    # Логируем переданную кривую
                    f.write("  INPUT curve:\n")
                    f.write("    Length: {:.2f} ft ({:.0f} mm)\n".format(
                        new_curve.Length, new_curve.Length * 304.8))
                    p0 = new_curve.GetEndPoint(0)
                    p1 = new_curve.GetEndPoint(1)
                    f.write("    Start: ({:.4f}, {:.4f}, {:.4f})\n".format(p0.X, p0.Y, p0.Z))
                    f.write("    End: ({:.4f}, {:.4f}, {:.4f})\n".format(p1.X, p1.Y, p1.Z))

                    # Проверяем РЕАЛЬНУЮ геометрию созданного стержня
                    f.write("  ACTUAL rebar geometry:\n")
                    actual_curve = get_rebar_centerline(new_rebar)
                    if actual_curve:
                        ap0 = actual_curve.GetEndPoint(0)
                        ap1 = actual_curve.GetEndPoint(1)
                        f.write("    Length: {:.2f} ft ({:.0f} mm)\n".format(
                            actual_curve.Length, actual_curve.Length * 304.8))
                        f.write("    Start: ({:.4f}, {:.4f}, {:.4f})\n".format(ap0.X, ap0.Y, ap0.Z))
                        f.write("    End: ({:.4f}, {:.4f}, {:.4f})\n".format(ap1.X, ap1.Y, ap1.Z))

                        # Проверка совпадения
                        length_diff = abs(actual_curve.Length - new_curve.Length) * 304.8
                        if length_diff > 10:  # больше 10мм разницы
                            f.write("    !!! WARNING: Length differs by {:.0f} mm !!!\n".format(length_diff))
                    else:
                        f.write("    ERROR: Could not get centerline!\n")

                    if layout_set:
                        f.write("  SetLayoutAsSingle: OK\n")
                    else:
                        f.write("  SetLayoutAsSingle: FAILED - {}\n".format(layout_error))
            except Exception as log_ex:
                try:
                    with codecs.open(log_path, 'a', 'utf-8') as f:
                        f.write("  LOG ERROR: {}\n".format(str(log_ex)))
                except Exception:
                    pass

        return new_rebar
    except Exception as e:
        # Логируем ошибку
        import codecs
        log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log_rebar.log")
        try:
            with codecs.open(log_path, 'a', 'utf-8') as f:
                f.write("ERROR creating rebar from original ID={}: {}\n".format(
                    original_rebar.Id.IntegerValue, str(e)))
        except Exception:
            pass
        return None


def merge_rebars(doc, rebar1, rebar2, host):
    """
    Склеить два стержня в один.

    Args:
        doc: Revit Document
        rebar1: Первый Rebar
        rebar2: Второй Rebar
        host: Host element

    Returns:
        Rebar (новый) или None
    """
    # Проверить что можно склеить
    can_merge, info = can_merge_rebars(rebar1, rebar2)
    if not can_merge:
        return None

    # Получить концы
    p1_start, p1_end = get_rebar_endpoints(rebar1)
    p2_start, p2_end = get_rebar_endpoints(rebar2)

    # Определить внешние концы (дальние друг от друга)
    end1_idx = info["end1_idx"]
    end2_idx = info["end2_idx"]

    # Внешний конец rebar1
    if end1_idx == 0:
        outer1 = p1_end
    else:
        outer1 = p1_start

    # Внешний конец rebar2
    if end2_idx == 0:
        outer2 = p2_end
    else:
        outer2 = p2_start

    # Создать линию от outer1 до outer2
    try:
        merged_line = Line.CreateBound(outer1, outer2)
    except Exception:
        return None

    # Создать новый стержень
    # Используем более длинный стержень как основу для параметров
    line1 = get_rebar_centerline(rebar1)
    line2 = get_rebar_centerline(rebar2)

    if line1.Length >= line2.Length:
        base_rebar = rebar1
    else:
        base_rebar = rebar2

    new_rebar = create_rebar_from_curve(doc, base_rebar, merged_line, host)

    return new_rebar


def get_area_reinforcement_in_host(doc, host):
    """
    Получить Area Reinforcement в хосте.

    Args:
        doc: Revit Document
        host: Host element

    Returns:
        list of AreaReinforcement elements
    """
    result = []
    host_id = host.Id

    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_AreaRein
    ).WhereElementIsNotElementType()

    for ar in collector:
        try:
            ar_host_id = ar.GetHostId()
            if ar_host_id == host_id:
                result.append(ar)
        except Exception:
            continue

    return result


def convert_area_reinforcement_to_rebars(doc, area_reinf, host):
    """
    Конвертировать Area Reinforcement в отдельные Rebar.

    Args:
        doc: Revit Document
        area_reinf: AreaReinforcement element
        host: Host element

    Returns:
        list of created Rebar elements
    """
    created_rebars = []

    try:
        # Получить все стержни в зоне
        rebar_ids = area_reinf.GetRebarInSystemIds()

        for rid in rebar_ids:
            rebar = doc.GetElement(rid)
            if rebar is None:
                continue

            # Получить геометрию
            curve = get_rebar_centerline(rebar)
            if curve is None:
                continue

            # Создать независимый Rebar
            new_rebar = create_rebar_from_curve(doc, rebar, curve, host)
            if new_rebar is not None:
                created_rebars.append(new_rebar)

        # Удалить Area Reinforcement
        doc.Delete(area_reinf.Id)

    except Exception:
        pass

    return created_rebars


def get_floor_sketch_openings(floor):
    """
    Получить отверстия из скетча плиты (inner loops).

    Отверстия созданные через "Edit Boundary" появляются как inner loops
    в скетче плиты, а не как отдельные Opening элементы.

    Args:
        floor: Floor element

    Returns:
        list of dict with keys:
            - 'loop': CurveLoop
            - 'solid': Solid (для проверки пересечений)
            - 'center': XYZ (центр отверстия)
            - 'area': float (площадь в кв.м)
            - 'index': int (индекс inner loop)
    """
    result = []

    try:
        # Получить sketch плиты
        sketch = floor.GetDependentElements(None)
        sketch_elem = None

        # Поиск Sketch элемента
        from Autodesk.Revit.DB import Sketch
        doc = floor.Document

        for elem_id in sketch:
            elem = doc.GetElement(elem_id)
            if isinstance(elem, Sketch):
                sketch_elem = elem
                break

        if sketch_elem is None:
            return result

        # Получить профили скетча
        profile = sketch_elem.Profile

        # Первый loop - внешний контур, остальные - отверстия
        loop_index = 0
        for curve_array in profile:
            if loop_index == 0:
                # Пропускаем внешний контур
                loop_index += 1
                continue

            # Это inner loop (отверстие)
            curves = []
            for curve in curve_array:
                curves.append(curve)

            if not curves:
                loop_index += 1
                continue

            # Создать CurveLoop
            try:
                curve_loop = CurveLoop.Create(list(curves))
            except Exception:
                loop_index += 1
                continue

            # Вычислить центр и площадь
            center, area = get_curve_loop_center_and_area(curves)

            # Создать solid для проверки пересечений
            solid = create_solid_from_curve_loop(curve_loop, floor)

            opening_info = {
                'loop': curve_loop,
                'curves': curves,  # Сохраняем curves для 2D проверки
                'solid': solid,
                'center': center,
                'area': area,
                'index': loop_index - 1  # 0-based для inner loops
            }
            result.append(opening_info)

            loop_index += 1

    except Exception:
        pass

    return result


def get_curve_loop_center_and_area(curves):
    """
    Вычислить центр и площадь контура.

    Args:
        curves: list of Curve

    Returns:
        tuple (center: XYZ, area: float в кв.м)
    """
    # Собираем все точки
    points = []
    for curve in curves:
        points.append(curve.GetEndPoint(0))

    if not points:
        return XYZ(0, 0, 0), 0

    # Центр как среднее арифметическое
    sum_x = sum(p.X for p in points)
    sum_y = sum(p.Y for p in points)
    sum_z = sum(p.Z for p in points)
    n = len(points)
    center = XYZ(sum_x / n, sum_y / n, sum_z / n)

    # Площадь по формуле Shoelace (для плоского контура)
    area = 0
    for i in range(n):
        j = (i + 1) % n
        area += points[i].X * points[j].Y
        area -= points[j].X * points[i].Y
    area = abs(area) / 2.0

    # Конвертировать кв.футы в кв.м
    area_sqm = area * 0.0929  # 1 sq.ft = 0.0929 sq.m

    return center, area_sqm


def create_solid_from_curve_loop(curve_loop, floor):
    """
    Создать Solid из CurveLoop для разрезания стержней.
    Solid создаётся на всю высоту проекта (-100м до +100м).
    """
    try:
        # Создаём solid на огромную высоту чтобы гарантировано покрыть любую отметку
        # От -100м до +100м (в футах: ~-328 до +328)
        extrusion_height = mm_to_feet(200000)  # 200 метров

        # Экструзия вверх
        solid = GeometryCreationUtilities.CreateExtrusionGeometry(
            [curve_loop],
            XYZ(0, 0, 1),
            extrusion_height / 2
        )

        # Сдвинуть вниз на половину высоты
        move_down = Transform.CreateTranslation(XYZ(0, 0, -extrusion_height / 4))
        solid = SolidUtils.CreateTransformed(solid, move_down)

        return solid

    except Exception:
        return None


def create_solid_from_curves(curves, floor):
    """
    Создать Solid из списка кривых для разрезания стержней.
    """
    try:
        # Создать CurveLoop из curves
        curve_loop = CurveLoop.Create(list(curves))
        return create_solid_from_curve_loop(curve_loop, floor)
    except Exception:
        return None


def get_opening_2d_bounds(curves):
    """
    Получить 2D границы отверстия (min/max X, Y).

    Args:
        curves: list of Curve

    Returns:
        tuple (min_x, min_y, max_x, max_y)
    """
    min_x = float('inf')
    min_y = float('inf')
    max_x = float('-inf')
    max_y = float('-inf')

    for curve in curves:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)

        min_x = min(min_x, p0.X, p1.X)
        min_y = min(min_y, p0.Y, p1.Y)
        max_x = max(max_x, p0.X, p1.X)
        max_y = max(max_y, p0.Y, p1.Y)

    return (min_x, min_y, max_x, max_y)


def point_in_polygon_2d(px, py, polygon_points):
    """
    Проверить находится ли точка внутри полигона (2D, ray casting).

    Args:
        px, py: координаты точки
        polygon_points: list of (x, y) tuples

    Returns:
        bool
    """
    n = len(polygon_points)
    inside = False

    j = n - 1
    for i in range(n):
        xi, yi = polygon_points[i]
        xj, yj = polygon_points[j]

        if ((yi > py) != (yj > py)) and (px < (xj - xi) * (py - yi) / (yj - yi) + xi):
            inside = not inside

        j = i

    return inside


def line_intersects_polygon_2d(line_start, line_end, polygon_points):
    """
    Проверить пересекает ли линия полигон в 2D.

    Args:
        line_start: (x, y)
        line_end: (x, y)
        polygon_points: list of (x, y)

    Returns:
        bool
    """
    # Проверить концы линии
    if point_in_polygon_2d(line_start[0], line_start[1], polygon_points):
        return True
    if point_in_polygon_2d(line_end[0], line_end[1], polygon_points):
        return True

    # Проверить пересечение линии с рёбрами полигона
    n = len(polygon_points)
    for i in range(n):
        j = (i + 1) % n
        if segments_intersect_2d(
            line_start, line_end,
            polygon_points[i], polygon_points[j]
        ):
            return True

    return False


def segments_intersect_2d(p1, p2, p3, p4):
    """
    Проверить пересекаются ли два отрезка в 2D.
    """
    def ccw(A, B, C):
        return (C[1] - A[1]) * (B[0] - A[0]) > (B[1] - A[1]) * (C[0] - A[0])

    return ccw(p1, p3, p4) != ccw(p2, p3, p4) and ccw(p1, p2, p3) != ccw(p1, p2, p4)


def check_rebar_intersects_opening_2d(rebar, opening_curves):
    """
    Проверить пересекает ли стержень отверстие в 2D (план).

    Использует простую проверку пересечения bounding boxes + проверку
    пересечения линии стержня с прямоугольником отверстия.

    Args:
        rebar: Rebar element
        opening_curves: list of Curve (контур отверстия)

    Returns:
        bool
    """
    # Получить границы отверстия
    op_bounds = get_opening_2d_bounds(opening_curves)
    if op_bounds[0] == float('inf'):
        return False

    op_min_x, op_min_y, op_max_x, op_max_y = op_bounds

    # Получить линию стержня в 2D
    line_start = None
    line_end = None

    # Попробовать centerline
    curve = get_rebar_centerline(rebar)
    if curve is not None:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        line_start = (p0.X, p0.Y)
        line_end = (p1.X, p1.Y)
    else:
        # Fallback: использовать BoundingBox
        bbox = rebar.get_BoundingBox(None)
        if bbox is not None:
            # Определить направление стержня по bbox
            dx = bbox.Max.X - bbox.Min.X
            dy = bbox.Max.Y - bbox.Min.Y

            if dx >= dy:
                # Горизонтальный стержень
                mid_y = (bbox.Min.Y + bbox.Max.Y) / 2
                line_start = (bbox.Min.X, mid_y)
                line_end = (bbox.Max.X, mid_y)
            else:
                # Вертикальный стержень
                mid_x = (bbox.Min.X + bbox.Max.X) / 2
                line_start = (mid_x, bbox.Min.Y)
                line_end = (mid_x, bbox.Max.Y)

    if line_start is None or line_end is None:
        return False

    # Простая проверка: пересекает ли линия прямоугольник отверстия
    return line_intersects_rect_2d(line_start, line_end, op_min_x, op_min_y, op_max_x, op_max_y)


def line_intersects_rect_2d(line_start, line_end, rect_min_x, rect_min_y, rect_max_x, rect_max_y):
    """
    Проверить пересекает ли линия прямоугольник в 2D.

    Использует алгоритм Cohen-Sutherland для клиппинга линии.

    Args:
        line_start: (x, y) начало линии
        line_end: (x, y) конец линии
        rect_min_x, rect_min_y, rect_max_x, rect_max_y: границы прямоугольника

    Returns:
        bool: True если линия пересекает или проходит через прямоугольник
    """
    x1, y1 = line_start
    x2, y2 = line_end

    # Коды региона для Cohen-Sutherland
    INSIDE = 0
    LEFT = 1
    RIGHT = 2
    BOTTOM = 4
    TOP = 8

    def compute_code(x, y):
        code = INSIDE
        if x < rect_min_x:
            code |= LEFT
        elif x > rect_max_x:
            code |= RIGHT
        if y < rect_min_y:
            code |= BOTTOM
        elif y > rect_max_y:
            code |= TOP
        return code

    code1 = compute_code(x1, y1)
    code2 = compute_code(x2, y2)

    while True:
        # Оба конца внутри прямоугольника
        if code1 == 0 and code2 == 0:
            return True

        # Оба конца снаружи с одной стороны - нет пересечения
        if (code1 & code2) != 0:
            return False

        # Выбрать точку снаружи
        code_out = code1 if code1 != 0 else code2

        # Найти точку пересечения с границей
        if code_out & TOP:
            x = x1 + (x2 - x1) * (rect_max_y - y1) / (y2 - y1) if y2 != y1 else x1
            y = rect_max_y
        elif code_out & BOTTOM:
            x = x1 + (x2 - x1) * (rect_min_y - y1) / (y2 - y1) if y2 != y1 else x1
            y = rect_min_y
        elif code_out & RIGHT:
            y = y1 + (y2 - y1) * (rect_max_x - x1) / (x2 - x1) if x2 != x1 else y1
            x = rect_max_x
        elif code_out & LEFT:
            y = y1 + (y2 - y1) * (rect_min_x - x1) / (x2 - x1) if x2 != x1 else y1
            x = rect_min_x
        else:
            break

        # Обновить точку и пересчитать код
        if code_out == code1:
            x1, y1 = x, y
            code1 = compute_code(x1, y1)
        else:
            x2, y2 = x, y
            code2 = compute_code(x2, y2)

    return False


def get_all_openings_in_floor(doc, floor):
    """
    Получить все отверстия в плите: Opening элементы, sketch-based и вложенные семейства.

    Args:
        doc: Revit Document
        floor: Floor element

    Returns:
        list of dict with keys:
            - 'type': 'element', 'sketch' или 'family'
            - 'element': Opening/FamilyInstance element
            - 'curves': list of Curve (для 2D проверки пересечений)
            - 'solid': Solid (для element type)
            - 'center': XYZ
            - 'area': float (кв.м)
            - 'name': str (для отображения)
            - 'id': уникальный идентификатор
    """
    result = []

    # 1. Получить Opening элементы
    openings = get_openings_in_floor(doc, floor)
    for i, opening in enumerate(openings):
        solid = get_opening_solid(opening)
        bbox = get_opening_bounding_box(opening)

        center = None
        curves = []

        if bbox:
            center = XYZ(
                (bbox.Min.X + bbox.Max.X) / 2,
                (bbox.Min.Y + bbox.Max.Y) / 2,
                (bbox.Min.Z + bbox.Max.Z) / 2
            )
            # Создать прямоугольный контур из bbox для 2D проверки
            curves = create_rect_curves_from_bbox(bbox)

        # Примерная площадь из bbox
        area = 0
        if bbox:
            area = abs(bbox.Max.X - bbox.Min.X) * abs(bbox.Max.Y - bbox.Min.Y) * 0.0929

        result.append({
            'type': 'element',
            'element': opening,
            'curves': curves,
            'solid': solid,
            'center': center,
            'area': area,
            'name': "Отверстие #{} (ID:{})".format(i + 1, opening.Id.IntegerValue),
            'id': "elem_{}".format(opening.Id.IntegerValue)
        })

    # 2. Получить sketch-based отверстия (inner loops)
    sketch_openings = get_floor_sketch_openings(floor)
    for i, op_info in enumerate(sketch_openings):
        result.append({
            'type': 'sketch',
            'element': None,
            'curves': op_info.get('curves', []),
            'solid': None,
            'center': op_info['center'],
            'area': op_info['area'],
            'name': "Контур #{} ({:.2f} м2)".format(i + 1, op_info['area']),
            'id': "sketch_{}".format(op_info['index'])
        })

    # 3. Получить вложенные семейства с отверстиями
    nested_openings = get_nested_family_openings(doc, floor)
    for i, fam_info in enumerate(nested_openings):
        result.append({
            'type': 'family',
            'element': fam_info['element'],
            'curves': fam_info.get('curves', []),
            'solid': fam_info.get('solid'),
            'center': fam_info['center'],
            'area': fam_info['area'],
            'name': "Семейство: {} (ID:{})".format(
                fam_info['family_name'], fam_info['element'].Id.IntegerValue
            ),
            'id': "family_{}".format(fam_info['element'].Id.IntegerValue),
            # Информация о bbox для отладки
            'bbox_source': fam_info.get('bbox_source', 'unknown'),
            'elem_size': fam_info.get('elem_size', (0, 0)),
            'curves_size': fam_info.get('curves_size', (0, 0))
        })

    return result


def get_solid_bbox(solid):
    """
    Получить BoundingBox из Solid.

    Args:
        solid: Revit Solid

    Returns:
        BoundingBoxXYZ или None
    """
    if solid is None:
        return None

    try:
        # Solid.GetBoundingBox(Transform) - передаём Identity чтобы получить в текущих координатах
        bbox = solid.GetBoundingBox(Transform.Identity)
        if bbox is not None:
            return bbox
    except Exception:
        pass

    # Fallback: вычислить bbox из вершин граней
    try:
        min_x = float('inf')
        min_y = float('inf')
        min_z = float('inf')
        max_x = float('-inf')
        max_y = float('-inf')
        max_z = float('-inf')

        found_points = False

        for face in solid.Faces:
            mesh = face.Triangulate()
            if mesh is None:
                continue
            for i in range(mesh.NumTriangles):
                triangle = mesh.get_Triangle(i)
                for j in range(3):
                    pt = triangle.get_Vertex(j)
                    found_points = True
                    min_x = min(min_x, pt.X)
                    min_y = min(min_y, pt.Y)
                    min_z = min(min_z, pt.Z)
                    max_x = max(max_x, pt.X)
                    max_y = max(max_y, pt.Y)
                    max_z = max(max_z, pt.Z)

        if found_points:
            bbox = BoundingBoxXYZ()
            bbox.Min = XYZ(min_x, min_y, min_z)
            bbox.Max = XYZ(max_x, max_y, max_z)
            return bbox
    except Exception:
        pass

    return None


def create_rect_curves_from_bbox(bbox):
    """
    Создать прямоугольный контур из BoundingBox (для 2D проверки).
    """
    curves = []
    try:
        min_pt = bbox.Min
        max_pt = bbox.Max

        # 4 угла прямоугольника
        p1 = XYZ(min_pt.X, min_pt.Y, 0)
        p2 = XYZ(max_pt.X, min_pt.Y, 0)
        p3 = XYZ(max_pt.X, max_pt.Y, 0)
        p4 = XYZ(min_pt.X, max_pt.Y, 0)

        curves.append(Line.CreateBound(p1, p2))
        curves.append(Line.CreateBound(p2, p3))
        curves.append(Line.CreateBound(p3, p4))
        curves.append(Line.CreateBound(p4, p1))
    except Exception:
        pass

    return curves


def get_nested_family_openings(doc, floor):
    """
    Получить вложенные семейства содержащие отверстия.

    Args:
        doc: Revit Document
        floor: Floor element

    Returns:
        list of dict with family info
    """
    result = []

    # Логирование для отладки
    import os
    log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log_family_openings.log")

    def log_debug(msg):
        try:
            with codecs.open(log_path, 'a', 'utf-8') as f:
                f.write(msg + "\n")
        except Exception:
            pass

    log_debug("\n=== get_nested_family_openings START ===")
    log_debug("Floor ID: {}".format(floor.Id.IntegerValue))

    # Получить bbox плиты для фильтрации
    floor_bbox = floor.get_BoundingBox(None)
    if floor_bbox is None:
        log_debug("ERROR: floor_bbox is None")
        return result

    # Ключевые слова для определения семейств отверстий
    opening_keywords = [
        'отверстие', 'проем', 'opening', 'hole', 'void', 'cut',
        'проход', 'гильза', 'sleeve'
    ]

    # Категории которые могут содержать отверстия
    opening_categories = [
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_SpecialityEquipment,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_PlumbingFixtures,
    ]

    for cat in opening_categories:
        try:
            collector = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()

            for elem in collector:
                try:
                    elem_bbox = elem.get_BoundingBox(None)
                    if elem_bbox is None:
                        continue

                    # Проверить что элемент пересекается с плитой по XY
                    if not bboxes_intersect_xy(floor_bbox, elem_bbox):
                        continue

                    # Получить имя семейства
                    family_name = ""
                    try:
                        elem_type = doc.GetElement(elem.GetTypeId())
                        if elem_type:
                            family_name = elem_type.FamilyName or ""
                    except Exception:
                        pass

                    # Проверить по имени семейства (содержит ключевые слова)
                    is_opening_by_name = any(
                        kw in family_name.lower() for kw in opening_keywords
                    )

                    log_debug("  Checking elem ID={}, family='{}'".format(
                        elem.Id.IntegerValue, family_name))
                    log_debug("    is_opening_by_name: {}".format(is_opening_by_name))

                    # Проверить что элемент содержит solid
                    solid = get_element_void_solid(elem)
                    log_debug("    solid found: {}".format(solid is not None))
                    if solid:
                        log_debug("    solid.Volume: {}".format(solid.Volume))

                    # Принять если имя содержит ключевое слово ИЛИ есть solid
                    if not is_opening_by_name and solid is None:
                        log_debug("    SKIP: not opening by name and no solid")
                        continue

                    # ВАЖНО: Использовать BoundingBox solid (фактическое отверстие),
                    # а не elem_bbox (весь элемент семейства)!
                    bbox_source = "element"
                    curves_bbox = elem_bbox

                    if solid is not None:
                        solid_bbox = get_solid_bbox(solid)
                        log_debug("    solid_bbox: {}".format(solid_bbox is not None))
                        if solid_bbox is not None:
                            curves_bbox = solid_bbox
                            bbox_source = "solid"
                            log_debug("    Using SOLID bbox")
                        else:
                            log_debug("    solid_bbox is None, trying params")
                    else:
                        log_debug("    No solid, trying to get dimensions from params")

                    # Если нет solid или не удалось получить его bbox,
                    # попробовать получить размеры из параметров семейства
                    if bbox_source == "element":
                        open_width, open_height = get_opening_dimensions_from_params(elem, doc)
                        log_debug("    Params: width={}, height={}".format(
                            open_width, open_height))

                        if open_width is not None and open_height is not None:
                            # Создать bbox из центра элемента и размеров отверстия
                            elem_center_x = (elem_bbox.Min.X + elem_bbox.Max.X) / 2
                            elem_center_y = (elem_bbox.Min.Y + elem_bbox.Max.Y) / 2

                            # Создать новый bbox с размерами отверстия
                            param_bbox = BoundingBoxXYZ()
                            param_bbox.Min = XYZ(
                                elem_center_x - open_width / 2,
                                elem_center_y - open_height / 2,
                                elem_bbox.Min.Z
                            )
                            param_bbox.Max = XYZ(
                                elem_center_x + open_width / 2,
                                elem_center_y + open_height / 2,
                                elem_bbox.Max.Z
                            )
                            curves_bbox = param_bbox
                            bbox_source = "params"
                            log_debug("    Using PARAMS bbox: {}x{} mm".format(
                                int(open_width * 304.8), int(open_height * 304.8)))

                    center = XYZ(
                        (curves_bbox.Min.X + curves_bbox.Max.X) / 2,
                        (curves_bbox.Min.Y + curves_bbox.Max.Y) / 2,
                        (curves_bbox.Min.Z + curves_bbox.Max.Z) / 2
                    )

                    area = abs(curves_bbox.Max.X - curves_bbox.Min.X) * \
                           abs(curves_bbox.Max.Y - curves_bbox.Min.Y) * 0.0929

                    # Логирование для отладки
                    elem_size_x = abs(elem_bbox.Max.X - elem_bbox.Min.X) * 304.8
                    elem_size_y = abs(elem_bbox.Max.Y - elem_bbox.Min.Y) * 304.8
                    curves_size_x = abs(curves_bbox.Max.X - curves_bbox.Min.X) * 304.8
                    curves_size_y = abs(curves_bbox.Max.Y - curves_bbox.Min.Y) * 304.8

                    log_debug("    ADDED: bbox_source={}, elem_size={}x{}, curves_size={}x{}".format(
                        bbox_source, int(elem_size_x), int(elem_size_y),
                        int(curves_size_x), int(curves_size_y)))

                    result.append({
                        'element': elem,
                        'solid': solid,
                        'curves': create_rect_curves_from_bbox(curves_bbox),
                        'center': center,
                        'area': area,
                        'family_name': family_name or "Unknown",
                        'bbox_source': bbox_source,
                        'elem_size': (elem_size_x, elem_size_y),
                        'curves_size': (curves_size_x, curves_size_y)
                    })

                except Exception as e:
                    log_debug("    EXCEPTION: {}".format(str(e)))
                    continue

        except Exception as e:
            log_debug("  Category exception: {}".format(str(e)))
            continue

    log_debug("=== get_nested_family_openings END, found {} families ===".format(len(result)))
    return result


def get_opening_dimensions_from_params(elem, doc):
    """
    Получить размеры отверстия из параметров семейства.

    Для ADSK семейств типа ADSK_ОбобщеннаяМодель_ОтверстиеПрямоугольное
    размеры хранятся в параметрах.

    Returns:
        tuple (width, height) в футах или (None, None)
    """
    width = None
    height = None

    # Список возможных имён параметров ширины (lowercase)
    width_param_names = [
        'adsk_размер_ширина', 'ширина', 'width', 'w', 'b',
        'adsk_отверстие_ширина', 'размер_ширина', 'ширинаотверстия',
        'ширина отверстия', 'opening width'
    ]

    # Список возможных имён параметров высоты (lowercase)
    height_param_names = [
        'adsk_размер_высота', 'высота', 'height', 'h', 'l', 'длина',
        'adsk_отверстие_высота', 'размер_высота', 'высотаотверстия',
        'высота отверстия', 'opening height', 'глубина', 'depth', 'length'
    ]

    try:
        # Сначала проверить параметры экземпляра
        for param in elem.Parameters:
            param_name = param.Definition.Name.lower()
            if param.StorageType.ToString() == "Double" and param.HasValue:
                value = param.AsDouble()
                if value > 0:
                    if param_name in width_param_names and width is None:
                        width = value
                    elif param_name in height_param_names and height is None:
                        height = value

        # Если не нашли, проверить параметры типа
        if width is None or height is None:
            elem_type = doc.GetElement(elem.GetTypeId())
            if elem_type:
                for param in elem_type.Parameters:
                    param_name = param.Definition.Name.lower()
                    if param.StorageType.ToString() == "Double" and param.HasValue:
                        value = param.AsDouble()
                        if value > 0:
                            if param_name in width_param_names and width is None:
                                width = value
                            elif param_name in height_param_names and height is None:
                                height = value
    except Exception:
        pass

    return (width, height)


def bboxes_intersect_xy(bbox1, bbox2):
    """Проверить пересекаются ли два BoundingBox в плане XY."""
    return not (bbox1.Max.X < bbox2.Min.X or bbox1.Min.X > bbox2.Max.X or
                bbox1.Max.Y < bbox2.Min.Y or bbox1.Min.Y > bbox2.Max.Y)


def get_element_void_solid(elem):
    """
    Получить void/cut solid из элемента если есть.
    """
    try:
        opt = Options()
        opt.ComputeReferences = False

        geom = elem.get_Geometry(opt)
        if geom is None:
            return None

        for geom_obj in geom:
            if isinstance(geom_obj, Solid):
                if geom_obj.Volume > 0:
                    return geom_obj
            elif isinstance(geom_obj, GeometryInstance):
                inst_geom = geom_obj.GetInstanceGeometry()
                for inst_obj in inst_geom:
                    if isinstance(inst_obj, Solid):
                        if inst_obj.Volume > 0:
                            return inst_obj
    except Exception:
        pass

    return None


def find_rebars_intersecting_opening_solid(doc, solid, host):
    """
    УСТАРЕВШИЙ - используйте find_rebars_intersecting_opening_2d
    """
    return []


def find_rebars_intersecting_opening_2d(doc, opening_curves, host, debug_log=False):
    """
    Найти стержни, пересекающие отверстие в 2D (план).

    Args:
        doc: Revit Document
        opening_curves: list of Curve (контур отверстия)
        host: Host element (плита)
        debug_log: Включить отладочное логирование

    Returns:
        list of Rebar elements
    """
    result = []

    if not opening_curves:
        return result

    # Получить границы отверстия для логирования
    op_bounds = get_opening_2d_bounds(opening_curves)

    # Получить все стержни в хосте
    rebars = get_rebars_in_host(doc, host)

    # Отладочный лог
    if debug_log:
        log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log_intersect.log")
        try:
            with codecs.open(log_path, 'a', 'utf-8') as f:
                f.write("\n=== find_rebars_intersecting_opening_2d ===\n")
                f.write("Opening bounds: X({:.2f}-{:.2f}) Y({:.2f}-{:.2f})\n".format(
                    op_bounds[0], op_bounds[2], op_bounds[1], op_bounds[3]))
                f.write("Total rebars in host: {}\n".format(len(rebars)))
        except Exception:
            pass

    for rebar in rebars:
        intersects = check_rebar_intersects_opening_2d(rebar, opening_curves)
        if intersects:
            result.append(rebar)

        # Подробный лог для первых нескольких стержней
        if debug_log and len(rebars) <= 20:
            import codecs
            log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log_intersect.log")
            try:
                curve = get_rebar_centerline(rebar)
                with codecs.open(log_path, 'a', 'utf-8') as f:
                    if curve:
                        p0 = curve.GetEndPoint(0)
                        p1 = curve.GetEndPoint(1)
                        f.write("  Rebar {}: ({:.2f},{:.2f})-({:.2f},{:.2f}) -> {}\n".format(
                            rebar.Id.IntegerValue, p0.X, p0.Y, p1.X, p1.Y,
                            "YES" if intersects else "NO"))
                    else:
                        f.write("  Rebar {}: NO CENTERLINE -> {}\n".format(
                            rebar.Id.IntegerValue, "YES" if intersects else "NO"))
            except Exception:
                pass

    if debug_log:
        log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log_intersect.log")
        try:
            with codecs.open(log_path, 'a', 'utf-8') as f:
                f.write("Result: {} intersecting rebars\n".format(len(result)))
        except Exception:
            pass

    return result


def line_intersects_bbox_transform(line, bbox_transform, buffer_feet=0):
    """
    Проверить пересечение линии с BoundingBox из Transform.

    Args:
        line: Line
        bbox_transform: BoundingBox от Solid.GetBoundingBox()
        buffer_feet: буфер

    Returns:
        bool
    """
    try:
        min_pt = bbox_transform.Min
        max_pt = bbox_transform.Max

        # Применить трансформ если есть
        transform = bbox_transform.Transform
        if transform is not None and not transform.IsIdentity:
            min_pt = transform.OfPoint(min_pt)
            max_pt = transform.OfPoint(max_pt)

        # Расширить на буфер
        min_x = min(min_pt.X, max_pt.X) - buffer_feet
        max_x = max(min_pt.X, max_pt.X) + buffer_feet
        min_y = min(min_pt.Y, max_pt.Y) - buffer_feet
        max_y = max(min_pt.Y, max_pt.Y) + buffer_feet
        min_z = min(min_pt.Z, max_pt.Z) - buffer_feet
        max_z = max(min_pt.Z, max_pt.Z) + buffer_feet

        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)

        if p0.X < min_x and p1.X < min_x:
            return False
        if p0.X > max_x and p1.X > max_x:
            return False
        if p0.Y < min_y and p1.Y < min_y:
            return False
        if p0.Y > max_y and p1.Y > max_y:
            return False
        if p0.Z < min_z and p1.Z < min_z:
            return False
        if p0.Z > max_z and p1.Z > max_z:
            return False

        return True
    except Exception:
        return True  # При ошибке - проверяем на всякий случай
