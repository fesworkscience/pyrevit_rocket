# -*- coding: utf-8 -*-
"""
PLY Visualization - визуализация точек с цветами.

Режимы:
- Раскраска по высоте (градиент от синего к красному)
- Использование оригинальных цветов из PLY
- Без цвета (серый по умолчанию)
"""


# Режимы раскраски
COLOR_MODE_NONE = "none"           # Без цвета (серый)
COLOR_MODE_HEIGHT = "height"       # По высоте (градиент)
COLOR_MODE_ORIGINAL = "original"   # Оригинальные цвета из PLY


def get_height_color(z, min_z, max_z):
    """
    Получить цвет по высоте (градиент синий -> зелёный -> красный).

    Args:
        z: высота точки
        min_z: минимальная высота
        max_z: максимальная высота

    Returns:
        tuple: (r, g, b) 0-255
    """
    if max_z <= min_z:
        return (128, 128, 128)

    # Нормализуем высоту в диапазон 0-1
    t = (z - min_z) / (max_z - min_z)
    t = max(0.0, min(1.0, t))

    # Градиент: синий (низ) -> голубой -> зелёный -> жёлтый -> красный (верх)
    if t < 0.25:
        # Синий -> Голубой
        ratio = t / 0.25
        r = 0
        g = int(255 * ratio)
        b = 255
    elif t < 0.5:
        # Голубой -> Зелёный
        ratio = (t - 0.25) / 0.25
        r = 0
        g = 255
        b = int(255 * (1 - ratio))
    elif t < 0.75:
        # Зелёный -> Жёлтый
        ratio = (t - 0.5) / 0.25
        r = int(255 * ratio)
        g = 255
        b = 0
    else:
        # Жёлтый -> Красный
        ratio = (t - 0.75) / 0.25
        r = 255
        g = int(255 * (1 - ratio))
        b = 0

    return (r, g, b)


def apply_colors(points, color_mode=COLOR_MODE_NONE, min_z=None, max_z=None):
    """
    Применить цвета к точкам.

    Args:
        points: список точек [(x, y, z, r, g, b), ...]
        color_mode: режим раскраски (none, height, original)
        min_z: минимальная высота для градиента (опционально)
        max_z: максимальная высота для градиента (опционально)

    Returns:
        list: точки с обновлёнными цветами [(x, y, z, r, g, b), ...]
    """
    if not points:
        return points

    if color_mode == COLOR_MODE_NONE:
        # Серый цвет для всех точек
        return [(p[0], p[1], p[2], 128, 128, 128) for p in points]

    elif color_mode == COLOR_MODE_ORIGINAL:
        # Используем оригинальные цвета, если нет - серый
        result = []
        for p in points:
            if p[3] is not None and p[4] is not None and p[5] is not None:
                result.append(p)
            else:
                result.append((p[0], p[1], p[2], 128, 128, 128))
        return result

    elif color_mode == COLOR_MODE_HEIGHT:
        # Раскраска по высоте
        if min_z is None or max_z is None:
            # Вычисляем границы автоматически
            z_values = [p[2] for p in points]
            min_z = min(z_values)
            max_z = max(z_values)

        result = []
        for p in points:
            r, g, b = get_height_color(p[2], min_z, max_z)
            result.append((p[0], p[1], p[2], r, g, b))
        return result

    else:
        return points


def get_color_for_revit(r, g, b):
    """
    Конвертирует RGB в формат для Revit (если нужно).

    Args:
        r, g, b: компоненты цвета 0-255

    Returns:
        tuple: (r, g, b) нормализованные
    """
    return (
        max(0, min(255, int(r))),
        max(0, min(255, int(g))),
        max(0, min(255, int(b)))
    )


def quantize_color(r, g, b, levels=8):
    """
    Квантизация цвета - уменьшение количества уникальных цветов.

    Args:
        r, g, b: компоненты цвета 0-255
        levels: количество уровней на канал (8 = 512 цветов, 16 = 4096 цветов)

    Returns:
        tuple: (r, g, b) квантизированный цвет
    """
    step = 256 // levels
    qr = ((r // step) * step) + step // 2
    qg = ((g // step) * step) + step // 2
    qb = ((b // step) * step) + step // 2
    return (min(255, qr), min(255, qg), min(255, qb))


def group_points_by_color(points, levels=8):
    """
    Группирует точки по квантизированным цветам.

    Args:
        points: список точек [(x, y, z, r, g, b), ...]
        levels: количество уровней квантизации на канал

    Returns:
        dict: {(r, g, b): [(x, y, z, r, g, b), ...], ...}
    """
    groups = {}

    for p in points:
        # Получаем цвет точки
        r = p[3] if p[3] is not None else 128
        g = p[4] if p[4] is not None else 128
        b = p[5] if p[5] is not None else 128

        # Квантизируем
        color_key = quantize_color(r, g, b, levels)

        if color_key not in groups:
            groups[color_key] = []
        groups[color_key].append(p)

    return groups


def has_colors(points):
    """
    Проверяет, есть ли цвета в точках.

    Args:
        points: список точек [(x, y, z, r, g, b), ...]

    Returns:
        bool: True если хотя бы часть точек имеет цвета
    """
    if not points:
        return False

    for p in points:
        if len(p) >= 6 and p[3] is not None:
            return True

    return False
