# -*- coding: utf-8 -*-
"""
Alignment Utils - математика для выравнивания облаков точек.

Алгоритмы:
- Rigid transformation (rotation + translation) из точечных соответствий
- ICP (Iterative Closest Point) для fine-tuning
"""

import math


def calculate_centroid(points):
    """
    Вычислить центроид (центр масс) точек.

    Args:
        points: список точек [(x, y, z), ...]

    Returns:
        tuple: (cx, cy, cz)
    """
    if not points:
        return (0, 0, 0)

    n = len(points)
    cx = sum(p[0] for p in points) / n
    cy = sum(p[1] for p in points) / n
    cz = sum(p[2] for p in points) / n

    return (cx, cy, cz)


def subtract_centroid(points, centroid):
    """
    Вычесть центроид из точек (центрирование).

    Args:
        points: список точек
        centroid: центроид (cx, cy, cz)

    Returns:
        list: центрированные точки
    """
    cx, cy, cz = centroid
    return [(p[0] - cx, p[1] - cy, p[2] - cz) for p in points]


def matrix_multiply_3x3(A, B):
    """Умножение матриц 3x3."""
    result = [[0, 0, 0], [0, 0, 0], [0, 0, 0]]
    for i in range(3):
        for j in range(3):
            for k in range(3):
                result[i][j] += A[i][k] * B[k][j]
    return result


def matrix_transpose_3x3(M):
    """Транспонирование матрицы 3x3."""
    return [
        [M[0][0], M[1][0], M[2][0]],
        [M[0][1], M[1][1], M[2][1]],
        [M[0][2], M[1][2], M[2][2]]
    ]


def matrix_vector_multiply(M, v):
    """Умножение матрицы 3x3 на вектор."""
    return (
        M[0][0] * v[0] + M[0][1] * v[1] + M[0][2] * v[2],
        M[1][0] * v[0] + M[1][1] * v[1] + M[1][2] * v[2],
        M[2][0] * v[0] + M[2][1] * v[1] + M[2][2] * v[2]
    )


def cross_product(a, b):
    """Векторное произведение."""
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0]
    )


def dot_product(a, b):
    """Скалярное произведение."""
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def vector_length(v):
    """Длина вектора."""
    return math.sqrt(v[0] ** 2 + v[1] ** 2 + v[2] ** 2)


def normalize_vector(v):
    """Нормализация вектора."""
    length = vector_length(v)
    if length < 1e-10:
        return (0, 0, 0)
    return (v[0] / length, v[1] / length, v[2] / length)


def compute_rotation_matrix_from_points(source_points, target_points):
    """
    Вычислить матрицу поворота методом SVD-подобного алгоритма.
    Упрощённая версия алгоритма Кабша для IronPython.

    Args:
        source_points: исходные точки (центрированные)
        target_points: целевые точки (центрированные)

    Returns:
        list: матрица поворота 3x3
    """
    # Вычисляем ковариационную матрицу H = sum(source * target^T)
    H = [[0, 0, 0], [0, 0, 0], [0, 0, 0]]

    for s, t in zip(source_points, target_points):
        for i in range(3):
            for j in range(3):
                si = [s[0], s[1], s[2]][i]
                tj = [t[0], t[1], t[2]][j]
                H[i][j] += si * tj

    # Используем итеративный метод для нахождения поворота
    # (упрощённая версия - работает для малых углов и хорошо распределённых точек)
    R = compute_rotation_iterative(H)

    return R


def compute_rotation_iterative(H, iterations=50):
    """
    Итеративное вычисление матрицы поворота из ковариационной матрицы.
    Метод полярного разложения.
    """
    # Начальное приближение - единичная матрица
    R = [[1, 0, 0], [0, 1, 0], [0, 0, 1]]

    for _ in range(iterations):
        # R_new = (H * R^T + R * H^T) / 2 - упрощённая итерация
        Rt = matrix_transpose_3x3(R)
        HR = matrix_multiply_3x3(H, Rt)

        # Ортогонализация через Gram-Schmidt
        R = orthogonalize_matrix(HR)

    return R


def orthogonalize_matrix(M):
    """Ортогонализация матрицы методом Грама-Шмидта."""
    # Извлекаем столбцы
    c0 = [M[0][0], M[1][0], M[2][0]]
    c1 = [M[0][1], M[1][1], M[2][1]]
    c2 = [M[0][2], M[1][2], M[2][2]]

    # Ортогонализация
    u0 = normalize_vector(c0)

    # c1 - проекция на u0
    proj = dot_product(c1, u0)
    c1_orth = (c1[0] - proj * u0[0], c1[1] - proj * u0[1], c1[2] - proj * u0[2])
    u1 = normalize_vector(c1_orth)

    # u2 = u0 x u1
    u2 = cross_product(u0, u1)
    u2 = normalize_vector(u2)

    # Собираем матрицу
    return [
        [u0[0], u1[0], u2[0]],
        [u0[1], u1[1], u2[1]],
        [u0[2], u1[2], u2[2]]
    ]


def calculate_rigid_transform(source_points, target_points):
    """
    Вычислить жёсткое преобразование (rotation + translation).

    Формула: target = R * source + t

    Args:
        source_points: список исходных точек [(x, y, z), ...]
        target_points: список целевых точек [(x, y, z), ...]

    Returns:
        tuple: (R, t) где R - матрица поворота 3x3, t - вектор переноса
    """
    if len(source_points) < 3 or len(target_points) < 3:
        # Недостаточно точек - возвращаем единичное преобразование
        return ([[1, 0, 0], [0, 1, 0], [0, 0, 1]], (0, 0, 0))

    if len(source_points) != len(target_points):
        raise ValueError("Количество точек должно совпадать")

    # Вычисляем центроиды
    source_centroid = calculate_centroid(source_points)
    target_centroid = calculate_centroid(target_points)

    # Центрируем точки
    source_centered = subtract_centroid(source_points, source_centroid)
    target_centered = subtract_centroid(target_points, target_centroid)

    # Вычисляем матрицу поворота
    R = compute_rotation_matrix_from_points(source_centered, target_centered)

    # Вычисляем вектор переноса: t = target_centroid - R * source_centroid
    rotated_source_centroid = matrix_vector_multiply(R, source_centroid)
    t = (
        target_centroid[0] - rotated_source_centroid[0],
        target_centroid[1] - rotated_source_centroid[1],
        target_centroid[2] - rotated_source_centroid[2]
    )

    return (R, t)


def apply_transform(points, R, t):
    """
    Применить преобразование к точкам.

    Args:
        points: список точек [(x, y, z), ...]
        R: матрица поворота 3x3
        t: вектор переноса (tx, ty, tz)

    Returns:
        list: преобразованные точки
    """
    result = []
    for p in points:
        rotated = matrix_vector_multiply(R, p)
        transformed = (
            rotated[0] + t[0],
            rotated[1] + t[1],
            rotated[2] + t[2]
        )
        result.append(transformed)
    return result


def distance_squared(p1, p2):
    """Квадрат расстояния между точками."""
    return (p1[0] - p2[0]) ** 2 + (p1[1] - p2[1]) ** 2 + (p1[2] - p2[2]) ** 2


def find_closest_point(point, point_cloud, grid=None, cell_size=0.5):
    """
    Найти ближайшую точку в облаке.

    Args:
        point: точка (x, y, z)
        point_cloud: облако точек
        grid: пространственная сетка (опционально)
        cell_size: размер ячейки сетки

    Returns:
        tuple: (индекс, расстояние^2)
    """
    if grid is not None:
        # Используем сетку для ускорения
        return find_closest_point_grid(point, point_cloud, grid, cell_size)

    # Наивный поиск
    min_dist_sq = float('inf')
    min_idx = -1

    for i, p in enumerate(point_cloud):
        dist_sq = distance_squared(point, p)
        if dist_sq < min_dist_sq:
            min_dist_sq = dist_sq
            min_idx = i

    return (min_idx, min_dist_sq)


def build_spatial_grid(points, cell_size=0.5):
    """
    Построить пространственную сетку для быстрого поиска.

    Returns:
        dict: {(cx, cy, cz): [индексы точек]}
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


def find_closest_point_grid(point, point_cloud, grid, cell_size):
    """Найти ближайшую точку с использованием сетки."""
    cx = int(math.floor(point[0] / cell_size))
    cy = int(math.floor(point[1] / cell_size))
    cz = int(math.floor(point[2] / cell_size))

    min_dist_sq = float('inf')
    min_idx = -1

    # Ищем в соседних ячейках
    for dx in range(-2, 3):
        for dy in range(-2, 3):
            for dz in range(-2, 3):
                key = (cx + dx, cy + dy, cz + dz)
                if key in grid:
                    for idx in grid[key]:
                        dist_sq = distance_squared(point, point_cloud[idx])
                        if dist_sq < min_dist_sq:
                            min_dist_sq = dist_sq
                            min_idx = idx

    return (min_idx, min_dist_sq)


def icp_align(source_points, target_points, max_iterations=50, tolerance=1e-6,
              max_correspondence_dist=None, progress_callback=None):
    """
    ICP (Iterative Closest Point) алгоритм для выравнивания облаков точек.

    Args:
        source_points: исходное облако точек (будет трансформировано)
        target_points: целевое облако точек (остаётся на месте)
        max_iterations: максимум итераций
        tolerance: порог сходимости (изменение ошибки)
        max_correspondence_dist: максимальное расстояние соответствия (м)
        progress_callback: функция callback(iteration, error)

    Returns:
        tuple: (R, t, final_error, iterations)
    """
    if not source_points or not target_points:
        return ([[1, 0, 0], [0, 1, 0], [0, 0, 1]], (0, 0, 0), 0, 0)

    # Рабочая копия source
    current_source = list(source_points)

    # Накопленное преобразование
    total_R = [[1, 0, 0], [0, 1, 0], [0, 0, 1]]
    total_t = (0, 0, 0)

    # Строим сетку для target
    cell_size = 0.1  # 10 см
    target_grid = build_spatial_grid(target_points, cell_size)

    # Порог расстояния
    if max_correspondence_dist is None:
        # Автоматически определяем из размера облака
        target_centroid = calculate_centroid(target_points)
        max_dist = max(
            math.sqrt(distance_squared(p, target_centroid))
            for p in target_points
        )
        max_correspondence_dist = max_dist * 0.5

    max_dist_sq = max_correspondence_dist ** 2

    prev_error = float('inf')

    for iteration in range(max_iterations):
        # 1. Находим соответствия
        correspondences = []

        for i, sp in enumerate(current_source):
            idx, dist_sq = find_closest_point_grid(sp, target_points, target_grid, cell_size)

            if idx >= 0 and dist_sq <= max_dist_sq:
                correspondences.append((i, idx, dist_sq))

        if len(correspondences) < 3:
            break

        # 2. Вычисляем среднюю ошибку
        total_error = sum(c[2] for c in correspondences)
        mean_error = math.sqrt(total_error / len(correspondences))

        if progress_callback:
            progress_callback(iteration, mean_error)

        # 3. Проверяем сходимость
        if abs(prev_error - mean_error) < tolerance:
            break

        prev_error = mean_error

        # 4. Извлекаем соответствующие точки
        src_corr = [current_source[c[0]] for c in correspondences]
        tgt_corr = [target_points[c[1]] for c in correspondences]

        # 5. Вычисляем преобразование
        R, t = calculate_rigid_transform(src_corr, tgt_corr)

        # 6. Применяем преобразование к текущему source
        current_source = apply_transform(current_source, R, t)

        # 7. Накапливаем преобразование
        # total_R_new = R * total_R
        total_R = matrix_multiply_3x3(R, total_R)
        # total_t_new = R * total_t + t
        rotated_t = matrix_vector_multiply(R, total_t)
        total_t = (rotated_t[0] + t[0], rotated_t[1] + t[1], rotated_t[2] + t[2])

    # Финальная ошибка
    final_error = 0
    count = 0
    for sp in current_source:
        idx, dist_sq = find_closest_point_grid(sp, target_points, target_grid, cell_size)
        if idx >= 0:
            final_error += dist_sq
            count += 1

    if count > 0:
        final_error = math.sqrt(final_error / count)

    return (total_R, total_t, final_error, iteration + 1)


def extract_points_from_lines(curves):
    """
    Извлечь точки из набора линий (для DirectShape).

    Предполагаем, что каждая группа из 3 линий - это крестик точки.

    Args:
        curves: список Revit Line

    Returns:
        list: точки [(x, y, z), ...]
    """
    points = []

    # Каждые 3 линии = 1 точка (крестик)
    for i in range(0, len(curves), 3):
        if i < len(curves):
            line = curves[i]
            # Берём середину первой линии как центр точки
            start = line.GetEndPoint(0)
            end = line.GetEndPoint(1)
            center = (
                (start.X + end.X) / 2,
                (start.Y + end.Y) / 2,
                (start.Z + end.Z) / 2
            )
            points.append(center)

    return points
