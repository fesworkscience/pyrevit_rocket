# -*- coding: utf-8 -*-
"""
PLY Filters - фильтрация точек облака.

Фильтры:
- Voxel Grid - равномерное прореживание по сетке вокселей
- Statistical Outlier Removal - удаление статистических выбросов
- Radius Outlier Removal - удаление изолированных точек
"""

import math


def voxel_grid_filter(points, voxel_size_m=0.05, progress_callback=None):
    """
    Voxel Grid фильтр - оставляет одну точку на воксель.
    Гарантирует равномерное распределение точек в пространстве.

    Args:
        points: список точек [(x, y, z, r, g, b), ...]
        voxel_size_m: размер воксела в метрах (по умолчанию 5 см)
        progress_callback: функция callback(percent)

    Returns:
        list: отфильтрованные точки
    """
    if not points or voxel_size_m <= 0:
        return points

    # Словарь: ключ воксела -> первая точка в этом вокселе
    voxel_map = {}

    total = len(points)
    update_interval = max(1, total // 100)

    for i, p in enumerate(points):
        # Вычисляем индекс воксела
        vx = int(math.floor(p[0] / voxel_size_m))
        vy = int(math.floor(p[1] / voxel_size_m))
        vz = int(math.floor(p[2] / voxel_size_m))

        key = (vx, vy, vz)

        # Сохраняем только первую точку в вокселе
        if key not in voxel_map:
            voxel_map[key] = p

        if progress_callback and i % update_interval == 0:
            progress_callback(int(100.0 * i / total))

    if progress_callback:
        progress_callback(100)

    return list(voxel_map.values())


def statistical_outlier_filter(points, k_neighbors=20, std_ratio=2.0, progress_callback=None):
    """
    Statistical Outlier Removal - удаляет точки с аномальным расстоянием до соседей.

    Алгоритм:
    1. Для каждой точки находим K ближайших соседей
    2. Вычисляем среднее расстояние до соседей
    3. Удаляем точки, где расстояние > mean + std_ratio * std

    Args:
        points: список точек [(x, y, z, r, g, b), ...]
        k_neighbors: количество соседей для анализа
        std_ratio: множитель стандартного отклонения
        progress_callback: функция callback(percent)

    Returns:
        list: отфильтрованные точки
    """
    if not points or len(points) < k_neighbors + 1:
        return points

    n = len(points)

    # Для оптимизации используем пространственное разбиение
    # Строим сетку для быстрого поиска соседей
    grid = _build_spatial_grid(points, cell_size=0.5)

    # Вычисляем среднее расстояние до K соседей для каждой точки
    mean_distances = []
    update_interval = max(1, n // 50)

    for i, p in enumerate(points):
        neighbors = _find_k_nearest(p, points, grid, k_neighbors)
        if neighbors:
            avg_dist = sum(neighbors) / len(neighbors)
            mean_distances.append(avg_dist)
        else:
            mean_distances.append(0)

        if progress_callback and i % update_interval == 0:
            progress_callback(int(50.0 * i / n))

    # Вычисляем глобальное среднее и стандартное отклонение
    if not mean_distances:
        return points

    global_mean = sum(mean_distances) / len(mean_distances)
    variance = sum((d - global_mean) ** 2 for d in mean_distances) / len(mean_distances)
    global_std = math.sqrt(variance)

    # Порог для фильтрации
    threshold = global_mean + std_ratio * global_std

    # Фильтруем точки
    filtered = []
    for i, p in enumerate(points):
        if mean_distances[i] <= threshold:
            filtered.append(p)

        if progress_callback and i % update_interval == 0:
            progress_callback(50 + int(50.0 * i / n))

    if progress_callback:
        progress_callback(100)

    return filtered


def radius_outlier_filter(points, radius_m=0.1, min_neighbors=5, progress_callback=None):
    """
    Radius Outlier Removal - удаляет точки с малым количеством соседей в радиусе.

    Args:
        points: список точек [(x, y, z, r, g, b), ...]
        radius_m: радиус поиска соседей в метрах
        min_neighbors: минимальное количество соседей для сохранения точки
        progress_callback: функция callback(percent)

    Returns:
        list: отфильтрованные точки
    """
    if not points:
        return points

    n = len(points)
    radius_sq = radius_m * radius_m

    # Строим пространственную сетку
    grid = _build_spatial_grid(points, cell_size=radius_m * 2)

    filtered = []
    update_interval = max(1, n // 100)

    for i, p in enumerate(points):
        # Считаем соседей в радиусе
        neighbor_count = _count_neighbors_in_radius(p, i, points, grid, radius_sq, radius_m * 2)

        if neighbor_count >= min_neighbors:
            filtered.append(p)

        if progress_callback and i % update_interval == 0:
            progress_callback(int(100.0 * i / n))

    if progress_callback:
        progress_callback(100)

    return filtered


def _build_spatial_grid(points, cell_size=0.5):
    """
    Строит пространственную сетку для быстрого поиска соседей.

    Returns:
        dict: {(cx, cy, cz): [индексы точек в ячейке]}
    """
    grid = {}

    for i, p in enumerate(points):
        cx = int(math.floor(p[0] / cell_size))
        cy = int(math.floor(p[1] / cell_size))
        cz = int(math.floor(p[2] / cell_size))

        key = (cx, cy, cz)
        if key not in grid:
            grid[key] = []
        grid[key].append(i)

    return grid


def _find_k_nearest(point, all_points, grid, k, cell_size=0.5):
    """
    Находит K ближайших соседей для точки.

    Returns:
        list: расстояния до K ближайших соседей
    """
    cx = int(math.floor(point[0] / cell_size))
    cy = int(math.floor(point[1] / cell_size))
    cz = int(math.floor(point[2] / cell_size))

    # Ищем в соседних ячейках
    distances = []

    for dx in range(-2, 3):
        for dy in range(-2, 3):
            for dz in range(-2, 3):
                key = (cx + dx, cy + dy, cz + dz)
                if key in grid:
                    for idx in grid[key]:
                        other = all_points[idx]
                        dist_sq = (point[0] - other[0]) ** 2 + \
                                  (point[1] - other[1]) ** 2 + \
                                  (point[2] - other[2]) ** 2
                        if dist_sq > 0:  # Исключаем саму точку
                            distances.append(math.sqrt(dist_sq))

    # Сортируем и берём K ближайших
    distances.sort()
    return distances[:k]


def _count_neighbors_in_radius(point, point_idx, all_points, grid, radius_sq, cell_size):
    """
    Считает количество соседей в заданном радиусе.
    """
    cx = int(math.floor(point[0] / cell_size))
    cy = int(math.floor(point[1] / cell_size))
    cz = int(math.floor(point[2] / cell_size))

    count = 0

    # Ищем в соседних ячейках
    for dx in range(-1, 2):
        for dy in range(-1, 2):
            for dz in range(-1, 2):
                key = (cx + dx, cy + dy, cz + dz)
                if key in grid:
                    for idx in grid[key]:
                        if idx != point_idx:
                            other = all_points[idx]
                            dist_sq = (point[0] - other[0]) ** 2 + \
                                      (point[1] - other[1]) ** 2 + \
                                      (point[2] - other[2]) ** 2
                            if dist_sq <= radius_sq:
                                count += 1

    return count
