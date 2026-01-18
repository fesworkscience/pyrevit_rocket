# -*- coding: utf-8 -*-
"""Generate GitHub Social Preview image (1280x640)"""

from PIL import Image, ImageDraw, ImageFont
import os

# Размеры для GitHub Social Preview
WIDTH = 1280
HEIGHT = 640

# Цвета
BG_COLOR = (24, 24, 27)  # Тёмный фон
ACCENT_COLOR = (59, 130, 246)  # Синий акцент
TEXT_COLOR = (255, 255, 255)  # Белый текст
SUBTEXT_COLOR = (161, 161, 170)  # Серый подтекст
GRADIENT_TOP = (30, 41, 59)
GRADIENT_BOTTOM = (15, 23, 42)

def create_gradient(width, height, color1, color2):
    """Создать вертикальный градиент"""
    img = Image.new('RGB', (width, height))
    for y in range(height):
        ratio = y / height
        r = int(color1[0] * (1 - ratio) + color2[0] * ratio)
        g = int(color1[1] * (1 - ratio) + color2[1] * ratio)
        b = int(color1[2] * (1 - ratio) + color2[2] * ratio)
        for x in range(width):
            img.putpixel((x, y), (r, g, b))
    return img

def get_font(size, bold=False):
    """Получить шрифт"""
    # Попробуем разные шрифты
    font_names = [
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/calibrib.ttf" if bold else "C:/Windows/Fonts/calibri.ttf",
    ]
    for font_path in font_names:
        if os.path.exists(font_path):
            try:
                return ImageFont.truetype(font_path, size)
            except:
                pass
    return ImageFont.load_default()

def draw_rounded_rect(draw, xy, radius, fill):
    """Нарисовать прямоугольник со скруглёнными углами"""
    x1, y1, x2, y2 = xy
    draw.rectangle([x1 + radius, y1, x2 - radius, y2], fill=fill)
    draw.rectangle([x1, y1 + radius, x2, y2 - radius], fill=fill)
    draw.ellipse([x1, y1, x1 + radius * 2, y1 + radius * 2], fill=fill)
    draw.ellipse([x2 - radius * 2, y1, x2, y1 + radius * 2], fill=fill)
    draw.ellipse([x1, y2 - radius * 2, x1 + radius * 2, y2], fill=fill)
    draw.ellipse([x2 - radius * 2, y2 - radius * 2, x2, y2], fill=fill)

def main():
    # Создаём градиентный фон
    img = create_gradient(WIDTH, HEIGHT, GRADIENT_TOP, GRADIENT_BOTTOM)
    draw = ImageDraw.Draw(img)

    # Добавляем декоративные элементы (сетка точек)
    for x in range(0, WIDTH, 40):
        for y in range(0, HEIGHT, 40):
            alpha = 0.03 + (y / HEIGHT) * 0.02
            dot_color = (int(59 * alpha + GRADIENT_TOP[0] * (1-alpha)),
                        int(130 * alpha + GRADIENT_TOP[1] * (1-alpha)),
                        int(246 * alpha + GRADIENT_TOP[2] * (1-alpha)))
            draw.ellipse([x-1, y-1, x+1, y+1], fill=dot_color)

    # Декоративные линии
    for i in range(3):
        y = 200 + i * 150
        line_color = (40, 50, 70)
        draw.line([(0, y), (WIDTH, y)], fill=line_color, width=1)

    # Акцентная полоса сверху
    draw.rectangle([0, 0, WIDTH, 4], fill=ACCENT_COLOR)

    # Иконка/логотип (стилизованная ракета из геометрических фигур)
    icon_x, icon_y = 540, 140
    # Корпус ракеты
    draw.polygon([(icon_x, icon_y), (icon_x + 100, icon_y + 50),
                  (icon_x + 100, icon_y + 120), (icon_x, icon_y + 170),
                  (icon_x - 100, icon_y + 120), (icon_x - 100, icon_y + 50)],
                 fill=ACCENT_COLOR)
    # Верхушка
    draw.polygon([(icon_x, icon_y - 60), (icon_x + 60, icon_y + 20),
                  (icon_x - 60, icon_y + 20)], fill=(99, 160, 255))
    # Окно
    draw.ellipse([icon_x - 25, icon_y + 40, icon_x + 25, icon_y + 90], fill=(30, 41, 59))
    draw.ellipse([icon_x - 18, icon_y + 47, icon_x + 18, icon_y + 83], fill=(120, 180, 255))
    # Пламя
    draw.polygon([(icon_x - 40, icon_y + 170), (icon_x, icon_y + 250),
                  (icon_x + 40, icon_y + 170)], fill=(251, 146, 60))
    draw.polygon([(icon_x - 20, icon_y + 170), (icon_x, icon_y + 220),
                  (icon_x + 20, icon_y + 170)], fill=(253, 224, 71))

    # Заголовок
    title_font = get_font(72, bold=True)
    title = "CPSK Tools"
    bbox = draw.textbbox((0, 0), title, font=title_font)
    title_width = bbox[2] - bbox[0]
    draw.text(((WIDTH - title_width) // 2, 340), title, font=title_font, fill=TEXT_COLOR)

    # Подзаголовок
    subtitle_font = get_font(32)
    subtitle = "Industrial Building Automation for Revit"
    bbox = draw.textbbox((0, 0), subtitle, font=subtitle_font)
    subtitle_width = bbox[2] - bbox[0]
    draw.text(((WIDTH - subtitle_width) // 2, 430), subtitle, font=subtitle_font, fill=SUBTEXT_COLOR)

    # Теги внизу
    tags = ["pyRevit", "BIM", "Dynamo", "IDS", "Python"]
    tag_font = get_font(20)
    tag_y = 520
    total_width = 0
    tag_padding = 20
    tag_spacing = 15

    # Вычисляем общую ширину тегов
    tag_widths = []
    for tag in tags:
        bbox = draw.textbbox((0, 0), tag, font=tag_font)
        w = bbox[2] - bbox[0] + tag_padding * 2
        tag_widths.append(w)
        total_width += w + tag_spacing
    total_width -= tag_spacing

    # Рисуем теги
    tag_x = (WIDTH - total_width) // 2
    for i, tag in enumerate(tags):
        w = tag_widths[i]
        draw_rounded_rect(draw, [tag_x, tag_y, tag_x + w, tag_y + 36], 18, (55, 65, 81))
        bbox = draw.textbbox((0, 0), tag, font=tag_font)
        text_w = bbox[2] - bbox[0]
        draw.text((tag_x + (w - text_w) // 2, tag_y + 7), tag, font=tag_font, fill=SUBTEXT_COLOR)
        tag_x += w + tag_spacing

    # Версия в углу
    version_font = get_font(18)
    draw.text((WIDTH - 100, HEIGHT - 40), "v1.0.53", font=version_font, fill=SUBTEXT_COLOR)

    # pyRevit badge
    badge_font = get_font(16)
    draw.text((40, HEIGHT - 40), "Built with pyRevit", font=badge_font, fill=SUBTEXT_COLOR)

    # Сохраняем
    output_path = os.path.join(os.path.dirname(__file__), "social_preview.png")
    img.save(output_path, "PNG", quality=95)
    print("Social preview saved to:", output_path)
    return output_path

if __name__ == "__main__":
    main()
