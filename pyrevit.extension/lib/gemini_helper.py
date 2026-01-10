# -*- coding: utf-8 -*-
"""
Gemini API Helper - внешний скрипт для работы с Google Gemini API.
Вызывается через subprocess из IronPython скриптов.

Использование:
    python gemini_helper.py --image "path/to/image.png" --prompt "prompt text" --output "path/to/output"

Возвращает через stdout JSON с результатом:
    {"success": true, "output_path": "path/to/generated.png"}
    или
    {"success": false, "error": "error message"}
"""

import os
import sys
import json
import argparse
import mimetypes
import base64
from datetime import datetime



def save_binary_file(file_name, data):
    """Сохранить бинарные данные в файл."""
    with open(file_name, "wb") as f:
        f.write(data)
    return file_name


def generate_image(image_path, prompt, output_dir, token):
    """
    Отправить изображение в Gemini и получить сгенерированное изображение.

    Args:
        image_path: путь к исходному изображению
        prompt: промпт для генерации
        output_dir: директория для сохранения результата
        token: API токен Gemini

    Returns:
        dict с результатом: {"success": bool, "output_path": str или "error": str}
    """
    try:
        from google import genai
        from google.genai import types
    except ImportError:
        return {"success": False, "error": "google-genai not installed. Run: pip install google-genai"}

    if not token:
        return {"success": False, "error": "GEMINI_TOKEN not provided"}

    if not os.path.exists(image_path):
        return {"success": False, "error": "Image file not found: {}".format(image_path)}

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # Инициализация клиента
        client = genai.Client(api_key=token)

        # Читаем изображение
        with open(image_path, "rb") as f:
            image_data = f.read()

        # Определяем MIME тип
        mime_type, _ = mimetypes.guess_type(image_path)
        if mime_type is None:
            mime_type = "image/png"

        # Формируем контент с изображением и промптом
        contents = [
            types.Content(
                role="user",
                parts=[
                    types.Part.from_bytes(
                        data=image_data,
                        mime_type=mime_type
                    ),
                    types.Part.from_text(text=prompt),
                ],
            ),
        ]

        # Конфигурация генерации
        generate_content_config = types.GenerateContentConfig(
            response_modalities=["IMAGE", "TEXT"],
            image_config=types.ImageConfig(
                image_size="1K",
            ),
        )

        # Генерация
        model = "gemini-2.0-flash-exp-image-generation"

        output_path = None
        text_response = ""

        for chunk in client.models.generate_content_stream(
            model=model,
            contents=contents,
            config=generate_content_config,
        ):
            if (
                chunk.candidates is None
                or chunk.candidates[0].content is None
                or chunk.candidates[0].content.parts is None
            ):
                continue

            for part in chunk.candidates[0].content.parts:
                if part.inline_data and part.inline_data.data:
                    # Получили изображение
                    inline_data = part.inline_data
                    data_buffer = inline_data.data
                    file_extension = mimetypes.guess_extension(inline_data.mime_type) or ".png"

                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    file_name = "gemini_render_{}{}".format(timestamp, file_extension)
                    output_path = os.path.join(output_dir, file_name)

                    save_binary_file(output_path, data_buffer)
                elif part.text:
                    text_response += part.text

        if output_path and os.path.exists(output_path):
            return {
                "success": True,
                "output_path": output_path,
                "text_response": text_response
            }
        else:
            error_msg = "No image generated"
            if text_response:
                error_msg += ". Model response: " + text_response[:500]
            return {"success": False, "error": error_msg}

    except Exception as e:
        return {"success": False, "error": str(e)}


def main():
    parser = argparse.ArgumentParser(description='Gemini Image Generation Helper')
    parser.add_argument('--image', required=True, help='Path to input image')
    parser.add_argument('--prompt', required=True, help='Prompt for image generation')
    parser.add_argument('--output', required=True, help='Output directory for generated image')
    parser.add_argument('--token', required=True, help='Gemini API token')

    args = parser.parse_args()

    result = generate_image(args.image, args.prompt, args.output, args.token)

    # Выводим результат как JSON для парсинга в IronPython
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
