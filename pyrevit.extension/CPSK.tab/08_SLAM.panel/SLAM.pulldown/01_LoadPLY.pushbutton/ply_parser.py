# -*- coding: utf-8 -*-
"""
PLY Parser - парсинг PLY файлов с оптимизацией производительности.

Особенности:
- Блочное чтение (не по одной точке)
- Потоковая обработка (генератор)
- Прогресс-бар при загрузке
- Поддержка binary и ASCII форматов
- Конвертация ARKit (Y-up) -> Revit (Z-up)
"""

import struct
import codecs


class PLYHeader:
    """Информация из заголовка PLY файла."""

    def __init__(self):
        self.vertex_count = 0
        self.has_colors = False
        self.is_binary = False
        self.header_size = 0
        self.format = "unknown"


def parse_header(file_path):
    """
    Парсит заголовок PLY файла.

    Args:
        file_path: путь к PLY файлу

    Returns:
        PLYHeader или None при ошибке
    """
    header = PLYHeader()

    try:
        with open(file_path, 'rb') as f:
            line = f.readline().decode('ascii', errors='ignore').strip()
            if line != 'ply':
                return None

            while True:
                line = f.readline().decode('ascii', errors='ignore').strip()
                header.header_size = f.tell()

                if line.startswith('format'):
                    if 'binary_little_endian' in line:
                        header.is_binary = True
                        header.format = "binary_little_endian"
                    elif 'binary_big_endian' in line:
                        header.is_binary = True
                        header.format = "binary_big_endian"
                    else:
                        header.format = "ascii"

                elif line.startswith('element vertex'):
                    header.vertex_count = int(line.split()[-1])

                elif line.startswith('property'):
                    if 'red' in line or 'green' in line or 'blue' in line:
                        header.has_colors = True

                elif line == 'end_header':
                    break

    except Exception:
        return None

    return header


def parse_ply_streaming(file_path, progress_callback=None, chunk_size=10000):
    """
    Потоковый парсинг PLY файла (генератор).
    Возвращает точки порциями для экономии памяти.

    Args:
        file_path: путь к PLY файлу
        progress_callback: функция callback(percent) для прогресса
        chunk_size: размер порции точек

    Yields:
        list of tuples: [(x, y, z, r, g, b), ...] в метрах, цвета 0-255 или None
    """
    header = parse_header(file_path)

    if header is None:
        return

    if header.is_binary:
        for chunk in _parse_binary_streaming(file_path, header, progress_callback, chunk_size):
            yield chunk
    else:
        for chunk in _parse_ascii_streaming(file_path, header, progress_callback, chunk_size):
            yield chunk


def _parse_binary_streaming(file_path, header, progress_callback, chunk_size):
    """Потоковый парсинг binary PLY."""

    # Размер одной вершины: 3 float (xyz) + опционально 3 uchar (rgb)
    vertex_size = 12 + (3 if header.has_colors else 0)

    # Размер блока для чтения (много вершин за раз)
    block_vertices = min(chunk_size, 50000)
    block_size = vertex_size * block_vertices

    with open(file_path, 'rb') as f:
        f.seek(header.header_size)

        points = []
        vertices_read = 0

        while vertices_read < header.vertex_count:
            # Сколько вершин осталось
            remaining = header.vertex_count - vertices_read
            to_read = min(block_vertices, remaining)

            # Читаем блок данных
            data = f.read(vertex_size * to_read)
            if not data:
                break

            # Парсим блок
            offset = 0
            for _ in range(to_read):
                if offset + 12 > len(data):
                    break

                # Координаты
                x, y, z = struct.unpack_from('<fff', data, offset)
                offset += 12

                # Конвертация ARKit (Y-up) -> Revit (Z-up)
                revit_x = x
                revit_y = -z  # ARKit Z backward -> Revit Y forward
                revit_z = y   # ARKit Y up -> Revit Z up

                # Цвета
                r, g, b = None, None, None
                if header.has_colors and offset + 3 <= len(data):
                    r, g, b = struct.unpack_from('<BBB', data, offset)
                    offset += 3

                points.append((revit_x, revit_y, revit_z, r, g, b))
                vertices_read += 1

                # Отдаём порцию
                if len(points) >= chunk_size:
                    if progress_callback:
                        progress_callback(int(100.0 * vertices_read / header.vertex_count))
                    yield points
                    points = []

        # Отдаём остаток
        if points:
            if progress_callback:
                progress_callback(100)
            yield points


def _parse_ascii_streaming(file_path, header, progress_callback, chunk_size):
    """Потоковый парсинг ASCII PLY."""

    with codecs.open(file_path, 'r', 'utf-8') as f:
        # Пропускаем заголовок
        for line in f:
            if line.strip() == 'end_header':
                break

        points = []
        vertices_read = 0

        for line in f:
            if vertices_read >= header.vertex_count:
                break

            parts = line.strip().split()
            if len(parts) < 3:
                continue

            # Координаты
            x, y, z = float(parts[0]), float(parts[1]), float(parts[2])

            # Конвертация ARKit -> Revit
            revit_x = x
            revit_y = -z
            revit_z = y

            # Цвета
            r, g, b = None, None, None
            if header.has_colors and len(parts) >= 6:
                r, g, b = int(parts[3]), int(parts[4]), int(parts[5])

            points.append((revit_x, revit_y, revit_z, r, g, b))
            vertices_read += 1

            # Отдаём порцию
            if len(points) >= chunk_size:
                if progress_callback:
                    progress_callback(int(100.0 * vertices_read / header.vertex_count))
                yield points
                points = []

        # Отдаём остаток
        if points:
            if progress_callback:
                progress_callback(100)
            yield points


def parse_ply_full(file_path, progress_callback=None):
    """
    Полный парсинг PLY файла в память.

    Args:
        file_path: путь к PLY файлу
        progress_callback: функция callback(percent) для прогресса

    Returns:
        list of tuples: [(x, y, z, r, g, b), ...] или None при ошибке
    """
    header = parse_header(file_path)

    if header is None:
        return None

    all_points = []

    for chunk in parse_ply_streaming(file_path, progress_callback):
        all_points.extend(chunk)

    return all_points


def get_bounds(points):
    """
    Получить границы облака точек.

    Args:
        points: список точек [(x, y, z, ...), ...]

    Returns:
        tuple: (min_x, max_x, min_y, max_y, min_z, max_z) в метрах
    """
    if not points:
        return (0, 0, 0, 0, 0, 0)

    min_x = min_y = min_z = float('inf')
    max_x = max_y = max_z = float('-inf')

    for p in points:
        if p[0] < min_x: min_x = p[0]
        if p[0] > max_x: max_x = p[0]
        if p[1] < min_y: min_y = p[1]
        if p[1] > max_y: max_y = p[1]
        if p[2] < min_z: min_z = p[2]
        if p[2] > max_z: max_z = p[2]

    return (min_x, max_x, min_y, max_y, min_z, max_z)
