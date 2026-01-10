# -*- coding: utf-8 -*-
"""Экспорт вида в изображение высокого разрешения с интеграцией Gemini AI."""

__title__ = "Export\nImage"
__author__ = "CPSK"

# 1. Сначала import clr и стандартные модули
import clr
import os
import sys
import subprocess
import json
import codecs
from datetime import datetime

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

# 2. System и WinForms импорты
import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, ComboBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, AnchorStyles, CheckBox, Padding,
    GroupBox, TextBox, NumericUpDown, TrackBar,
    TickStyle, SaveFileDialog, PictureBox, PictureBoxSizeMode,
    ProgressBar, Application, ScrollBars
)
from System.Drawing import Point, Size, Font, FontStyle, Color, Image, Bitmap

# 3. Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# 4. Проверка авторизации
from cpsk_auth import require_auth
from cpsk_notify import show_error, show_success, show_warning, show_info
from cpsk_config import get_clean_env, get_venv_python, require_environment, get_venv_path
import shutil
if not require_auth():
    sys.exit()

# 5. pyrevit и Revit API
from pyrevit import revit

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    View3D,
    ViewPlan,
    ViewSection,
    ImageExportOptions,
    ExportRange,
    ImageFileType,
    ImageResolution,
    FitDirectionType,
    Transaction,
    BuiltInCategory,
    Color as RevitColor
)

doc = revit.doc
uidoc = revit.uidoc


def get_all_views():
    """Получить все виды, которые можно экспортировать."""
    views = []

    # 3D виды
    views3d = FilteredElementCollector(doc).OfClass(View3D).ToElements()
    for v in views3d:
        if not v.IsTemplate:
            views.append(v)

    # Разрезы
    sections = FilteredElementCollector(doc).OfClass(ViewSection).ToElements()
    for v in sections:
        if not v.IsTemplate:
            views.append(v)

    # Планы
    plans = FilteredElementCollector(doc).OfClass(ViewPlan).ToElements()
    for v in plans:
        if not v.IsTemplate:
            views.append(v)

    return views


def get_default_export_path():
    """Получить путь по умолчанию для экспорта."""
    if doc.PathName:
        project_dir = os.path.dirname(doc.PathName)
        project_name = os.path.splitext(os.path.basename(doc.PathName))[0]
        return os.path.join(project_dir, project_name + "_exports")
    return os.path.join(os.path.expanduser("~"), "Documents", "RevitExports")


def export_view_to_image(view, export_path, pixel_size=4096, file_type=ImageFileType.PNG,
                         hide_annotations=False, white_background=False):
    """Экспортировать вид в изображение."""
    export_dir = os.path.dirname(export_path)
    if not os.path.exists(export_dir):
        os.makedirs(export_dir)

    # Сохраняем оригинальные настройки для восстановления
    original_settings = {}

    try:
        # Применяем временные настройки в транзакции
        t = Transaction(doc, "CPSK: Временные настройки экспорта")
        t.Start()

        try:
            # Скрываем аннотации если нужно
            if hide_annotations:
                # Категории для скрытия
                categories_to_hide = [
                    BuiltInCategory.OST_SectionBox,
                    BuiltInCategory.OST_Levels,
                    BuiltInCategory.OST_Grids,
                    BuiltInCategory.OST_ReferencePoints,
                    BuiltInCategory.OST_CLines,
                    BuiltInCategory.OST_Sections,
                    BuiltInCategory.OST_Elev,
                    BuiltInCategory.OST_Callouts,
                    BuiltInCategory.OST_TextNotes,
                    BuiltInCategory.OST_Dimensions,
                    BuiltInCategory.OST_SpotElevations,
                    BuiltInCategory.OST_SpotCoordinates,
                    BuiltInCategory.OST_SpotSlopes,
                ]

                for cat_id in categories_to_hide:
                    try:
                        cat = doc.Settings.Categories.get_Item(cat_id)
                        if cat and view.CanCategoryBeHidden(cat.Id):
                            # Сохраняем оригинальное состояние
                            original_settings[cat.Id] = view.GetCategoryHidden(cat.Id)
                            view.SetCategoryHidden(cat.Id, True)
                    except:
                        pass

            # Белый фон для 3D видов
            if white_background and isinstance(view, View3D):
                try:
                    # Сохраняем оригинальный фон
                    display_style = view.GetRenderingSettings()
                    if display_style:
                        original_settings['background'] = display_style.BackgroundColor
                        # Устанавливаем белый фон
                        display_style.BackgroundColor = RevitColor(255, 255, 255)
                        view.SetRenderingSettings(display_style)
                except:
                    pass

            t.Commit()
        except Exception as e:
            t.RollBack()
            raise e

        # Настройки экспорта
        options = ImageExportOptions()
        options.ExportRange = ExportRange.SetOfViews
        options.SetViewsAndSheets([view.Id])
        options.FilePath = export_path
        options.FitDirection = FitDirectionType.Horizontal
        options.PixelSize = pixel_size
        options.ImageResolution = ImageResolution.DPI_300
        options.HLRandWFViewsFileType = file_type
        options.ShadowViewsFileType = file_type

        # Экспорт
        doc.ExportImage(options)

        # Восстанавливаем настройки
        t2 = Transaction(doc, "CPSK: Восстановить настройки")
        t2.Start()
        try:
            # Восстанавливаем видимость категорий
            for cat_id, was_hidden in original_settings.items():
                if cat_id != 'background':
                    try:
                        view.SetCategoryHidden(cat_id, was_hidden)
                    except:
                        pass

            # Восстанавливаем фон
            if 'background' in original_settings and isinstance(view, View3D):
                try:
                    display_style = view.GetRenderingSettings()
                    if display_style:
                        display_style.BackgroundColor = original_settings['background']
                        view.SetRenderingSettings(display_style)
                except:
                    pass

            t2.Commit()
        except:
            t2.RollBack()

        # Определяем расширение
        if file_type == ImageFileType.PNG:
            ext = ".png"
        elif file_type in [ImageFileType.JPEGLossless, ImageFileType.JPEGMedium, ImageFileType.JPEGSmallest]:
            ext = ".jpg"
        elif file_type == ImageFileType.TIFF:
            ext = ".tif"
        else:
            ext = ".bmp"

        # Ищем созданный файл
        view_name_safe = view.Name.replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace("<", "_").replace(">", "_").replace("|", "_")
        expected_file = export_path + " - " + view_name_safe + ext

        if os.path.exists(expected_file):
            return expected_file

        expected_file2 = export_path + ext
        if os.path.exists(expected_file2):
            return expected_file2

        base_name = os.path.basename(export_path)
        for f in os.listdir(export_dir):
            if f.startswith(base_name) and f.endswith(ext):
                return os.path.join(export_dir, f)

        return None

    except Exception as e:
        show_error("Ошибка экспорта", "Не удалось экспортировать вид", details=str(e))
        return None


class ImagePreviewForm(Form):
    """Форма предпросмотра экспортированного изображения с отправкой в Gemini."""

    def __init__(self, image_path, export_dir):
        self.image_path = image_path
        self.export_dir = export_dir
        self.gemini_result_path = None
        self.setup_form()

    def setup_form(self):
        self.Text = "Предпросмотр изображения"
        self.Width = 900
        self.Height = 700
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MinimumSize = Size(600, 500)

        # Верхняя панель с информацией
        top_panel = Panel()
        top_panel.Dock = DockStyle.Top
        top_panel.Height = 30
        top_panel.Padding = Padding(10, 5, 10, 5)

        self.lbl_path = Label()
        self.lbl_path.Text = "Файл: {}".format(os.path.basename(self.image_path))
        self.lbl_path.Dock = DockStyle.Fill
        self.lbl_path.Font = Font("Segoe UI", 9)
        top_panel.Controls.Add(self.lbl_path)

        # Центральная панель с изображением
        center_panel = Panel()
        center_panel.Dock = DockStyle.Fill
        center_panel.Padding = Padding(10)

        self.picture_box = PictureBox()
        self.picture_box.Dock = DockStyle.Fill
        self.picture_box.SizeMode = PictureBoxSizeMode.Zoom
        self.picture_box.BackColor = Color.White

        try:
            self.picture_box.Image = Image.FromFile(self.image_path)
        except Exception as e:
            self.lbl_path.Text = "Ошибка загрузки: {}".format(str(e))

        center_panel.Controls.Add(self.picture_box)

        # Нижняя панель с промптом и кнопками
        bottom_panel = Panel()
        bottom_panel.Dock = DockStyle.Bottom
        bottom_panel.Height = 180
        bottom_panel.Padding = Padding(10)

        # Группа промпта
        grp_prompt = GroupBox()
        grp_prompt.Text = "Промпт для Gemini AI (рендер изображения)"
        grp_prompt.Dock = DockStyle.Fill
        grp_prompt.Padding = Padding(5)

        self.txt_prompt = TextBox()
        self.txt_prompt.Multiline = True
        self.txt_prompt.ScrollBars = ScrollBars.Vertical
        self.txt_prompt.Dock = DockStyle.Fill
        self.txt_prompt.Text = self.get_default_prompt()
        self.txt_prompt.Font = Font("Segoe UI", 9)

        grp_prompt.Controls.Add(self.txt_prompt)

        # Панель кнопок
        btn_panel = Panel()
        btn_panel.Dock = DockStyle.Bottom
        btn_panel.Height = 45

        self.btn_close = Button()
        self.btn_close.Text = "Закрыть"
        self.btn_close.Size = Size(100, 35)
        self.btn_close.Location = Point(10, 5)
        self.btn_close.Click += self.on_close

        self.btn_open_folder = Button()
        self.btn_open_folder.Text = "Открыть папку"
        self.btn_open_folder.Size = Size(120, 35)
        self.btn_open_folder.Location = Point(120, 5)
        self.btn_open_folder.Click += self.on_open_folder

        self.btn_gemini = Button()
        self.btn_gemini.Text = "Отправить в Gemini"
        self.btn_gemini.Size = Size(160, 35)
        self.btn_gemini.Location = Point(250, 5)
        self.btn_gemini.Click += self.on_send_to_gemini

        self.lbl_status = Label()
        self.lbl_status.Text = ""
        self.lbl_status.Location = Point(420, 12)
        self.lbl_status.Size = Size(400, 20)
        self.lbl_status.ForeColor = Color.Gray

        btn_panel.Controls.Add(self.btn_close)
        btn_panel.Controls.Add(self.btn_open_folder)
        btn_panel.Controls.Add(self.btn_gemini)
        btn_panel.Controls.Add(self.lbl_status)

        bottom_panel.Controls.Add(grp_prompt)
        bottom_panel.Controls.Add(btn_panel)

        # Добавляем контролы (Fill первым!)
        self.Controls.Add(center_panel)
        self.Controls.Add(bottom_panel)
        self.Controls.Add(top_panel)

    def get_default_prompt(self):
        """Получить дефолтный промпт для Gemini."""
        return """Transform this architectural wireframe/3D model into a photorealistic render:

1. STYLE: Photorealistic architectural visualization, professional quality
2. MATERIALS:
   - Steel elements: galvanized steel with natural metallic finish
   - Clean, industrial aesthetic
3. LIGHTING:
   - Bright daylight, clear blue sky
   - Soft shadows, natural ambient occlusion
4. ENVIRONMENT:
   - Clean concrete or asphalt ground plane
   - Simple sky background with subtle clouds
5. CAMERA: Keep the same angle and composition as the original
6. QUALITY: High resolution, sharp details, professional architectural render

Make it look like a real photograph of a built structure."""

    def on_close(self, sender, args):
        """Закрыть форму."""
        self.Close()

    def on_open_folder(self, sender, args):
        """Открыть папку с изображением."""
        try:
            folder = os.path.dirname(self.image_path)
            os.startfile(folder)
        except Exception as e:
            show_error("Ошибка", "Не удалось открыть папку", details=str(e))

    def on_send_to_gemini(self, sender, args):
        """Отправить изображение в Gemini."""
        prompt = self.txt_prompt.Text.strip()
        if not prompt:
            show_warning("Промпт", "Введите промпт для генерации")
            return

        # Проверяем окружение
        if not require_environment():
            return

        venv_python = get_venv_python()
        if not os.path.exists(venv_python):
            show_error("Окружение", "Python venv не найден",
                      details="Путь: {}".format(venv_python))
            return

        # Путь к helper скрипту
        helper_script = os.path.join(LIB_DIR, "gemini_helper.py")
        if not os.path.exists(helper_script):
            show_error("Скрипт", "gemini_helper.py не найден",
                      details="Путь: {}".format(helper_script))
            return

        # Блокируем UI
        self.btn_gemini.Enabled = False
        self.btn_close.Enabled = False
        self.lbl_status.Text = "Отправка в Gemini... (30-60 сек)"
        self.lbl_status.ForeColor = Color.Blue
        self.Refresh()
        Application.DoEvents()

        # Создаём temp папку вне OneDrive (C:\cpsk_envs\temp_gemini)
        temp_dir = r"C:\cpsk_envs\temp_gemini"

        try:
            # Создаём temp папку
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)

            # Копируем изображение в temp (обход проблем OneDrive)
            temp_image_name = "input_{}{}".format(
                datetime.now().strftime("%Y%m%d_%H%M%S"),
                os.path.splitext(self.image_path)[1]
            )
            temp_image_path = os.path.join(temp_dir, temp_image_name)

            try:
                shutil.copy2(self.image_path, temp_image_path)
            except Exception as e:
                show_error("Ошибка копирования", "Не удалось скопировать изображение",
                          details="Из: {}\nВ: {}\nОшибка: {}".format(
                              self.image_path, temp_image_path, str(e)))
                return

            # Загружаем токен из .env файла (в корне проекта pyrevit_rocket)
            project_dir = os.path.dirname(os.path.dirname(LIB_DIR))
            env_file = os.path.join(project_dir, ".env")

            gemini_token = ""
            try:
                with codecs.open(env_file, 'r', 'utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith('GEMINI_TOKEN='):
                            gemini_token = line.split('=', 1)[1].strip().strip("'\"")
                            break
            except Exception as e:
                show_error("Ошибка", "Не удалось прочитать .env файл",
                          details="Путь: {}\nОшибка: {}".format(env_file, str(e)))
                return

            if not gemini_token:
                show_error("Токен не найден", "GEMINI_TOKEN не найден в .env файле",
                          details="Путь: {}".format(env_file))
                return

            # Запускаем helper скрипт (используем temp папку для входа и выхода)
            CREATE_NO_WINDOW = 0x08000000
            cmd = [
                venv_python,
                helper_script,
                "--image", temp_image_path,
                "--prompt", prompt,
                "--output", temp_dir,
                "--token", gemini_token
            ]

            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                env=get_clean_env(),
                creationflags=CREATE_NO_WINDOW
            )

            # Ждём завершения процесса (IronPython 2.7 не поддерживает timeout)
            # Используем poll() с ожиданием
            import time
            timeout_seconds = 120
            poll_interval = 0.5
            elapsed = 0

            while process.poll() is None and elapsed < timeout_seconds:
                time.sleep(poll_interval)
                elapsed += poll_interval
                # Обновляем UI каждые 5 секунд
                if int(elapsed) % 5 == 0:
                    self.lbl_status.Text = "Генерация... ({} сек)".format(int(elapsed))
                    Application.DoEvents()

            if process.poll() is None:
                # Процесс ещё работает - таймаут
                process.kill()
                self.lbl_status.Text = "Таймаут"
                self.lbl_status.ForeColor = Color.Red
                show_error("Gemini", "Превышено время ожидания (120 сек)")
                return

            stdout, stderr = process.communicate()

            # Парсим результат
            output = stdout.decode('utf-8', errors='ignore').strip()

            if output:
                result = json.loads(output)
            else:
                error_msg = stderr.decode('utf-8', errors='ignore').strip()
                result = {"success": False, "error": error_msg or "Empty response"}

            if result.get("success"):
                temp_output_path = result.get("output_path")
                if temp_output_path and os.path.exists(temp_output_path):
                    # Копируем результат из temp в папку экспорта
                    final_filename = os.path.basename(temp_output_path)
                    final_output_path = os.path.join(self.export_dir, final_filename)

                    try:
                        shutil.copy2(temp_output_path, final_output_path)
                    except Exception as e:
                        # Если не удалось скопировать, используем temp путь
                        final_output_path = temp_output_path

                    self.gemini_result_path = final_output_path
                    self.lbl_status.Text = "Готово!"
                    self.lbl_status.ForeColor = Color.Green

                    # Загружаем новое изображение
                    try:
                        if self.picture_box.Image:
                            self.picture_box.Image.Dispose()
                        self.picture_box.Image = Image.FromFile(final_output_path)
                        self.lbl_path.Text = "Gemini: {}".format(os.path.basename(final_output_path))
                    except Exception as e:
                        pass

                    show_success("Gemini", "Рендер сгенерирован",
                               details="Файл: {}".format(final_output_path))

                    # Открываем файл
                    try:
                        os.startfile(final_output_path)
                    except:
                        pass
                else:
                    self.lbl_status.Text = "Файл не создан"
                    self.lbl_status.ForeColor = Color.Red
                    show_error("Gemini", "Файл не создан",
                             details="output_path: {}".format(temp_output_path))
            else:
                error = result.get("error", "Unknown error")
                self.lbl_status.Text = "Ошибка"
                self.lbl_status.ForeColor = Color.Red
                show_error("Gemini", "Ошибка генерации", details=error)

        except Exception as e:
            self.lbl_status.Text = "Ошибка"
            self.lbl_status.ForeColor = Color.Red
            show_error("Gemini", "Ошибка при вызове Gemini", details=str(e))

        finally:
            self.btn_gemini.Enabled = True
            self.btn_close.Enabled = True

            # Очищаем temp папку
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
            except:
                pass


class ExportViewForm(Form):
    """Форма экспорта вида в изображение."""

    def __init__(self):
        self.views = get_all_views()
        self.setup_form()

    def setup_form(self):
        self.Text = "Экспорт вида в изображение"
        self.Width = 500
        self.Height = 385
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y_pos = 20

        # Выбор вида
        lbl_view = Label()
        lbl_view.Text = "Вид для экспорта:"
        lbl_view.Location = Point(15, y_pos)
        lbl_view.Width = 120

        self.cmb_views = ComboBox()
        self.cmb_views.Location = Point(140, y_pos - 3)
        self.cmb_views.Width = 330
        self.cmb_views.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList

        # Заполняем список видов
        current_view = uidoc.ActiveView
        selected_index = 0
        for i, view in enumerate(self.views):
            view_type = ""
            if isinstance(view, View3D):
                view_type = "[3D] "
            elif isinstance(view, ViewSection):
                view_type = "[Разрез] "
            elif isinstance(view, ViewPlan):
                view_type = "[План] "

            self.cmb_views.Items.Add("{}{}".format(view_type, view.Name))

            if view.Id == current_view.Id:
                selected_index = i

        if self.cmb_views.Items.Count > 0:
            self.cmb_views.SelectedIndex = selected_index

        y_pos += 35

        # Разрешение
        lbl_resolution = Label()
        lbl_resolution.Text = "Разрешение:"
        lbl_resolution.Location = Point(15, y_pos)
        lbl_resolution.Width = 120

        self.cmb_resolution = ComboBox()
        self.cmb_resolution.Location = Point(140, y_pos - 3)
        self.cmb_resolution.Width = 180
        self.cmb_resolution.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.cmb_resolution.Items.Add("800 (Превью)")
        self.cmb_resolution.Items.Add("1280 (HD)")
        self.cmb_resolution.Items.Add("1920 (Full HD)")
        self.cmb_resolution.Items.Add("2560 (2K)")
        self.cmb_resolution.Items.Add("3840 (4K)")
        self.cmb_resolution.Items.Add("4096 (4K Cinema)")
        self.cmb_resolution.Items.Add("5120 (5K)")
        self.cmb_resolution.Items.Add("7680 (8K)")
        self.cmb_resolution.SelectedIndex = 5  # 4096 по умолчанию

        y_pos += 35

        # Формат
        lbl_format = Label()
        lbl_format.Text = "Формат:"
        lbl_format.Location = Point(15, y_pos)
        lbl_format.Width = 120

        self.cmb_format = ComboBox()
        self.cmb_format.Location = Point(140, y_pos - 3)
        self.cmb_format.Width = 180
        self.cmb_format.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.cmb_format.Items.Add("PNG (без потерь)")
        self.cmb_format.Items.Add("JPEG (высокое)")
        self.cmb_format.Items.Add("JPEG (среднее)")
        self.cmb_format.Items.Add("JPEG (низкое)")
        self.cmb_format.Items.Add("TIFF")
        self.cmb_format.Items.Add("BMP")
        self.cmb_format.SelectedIndex = 0

        y_pos += 35

        # Путь сохранения
        lbl_path = Label()
        lbl_path.Text = "Папка:"
        lbl_path.Location = Point(15, y_pos)
        lbl_path.Width = 120

        self.txt_path = TextBox()
        self.txt_path.Location = Point(140, y_pos - 3)
        self.txt_path.Width = 280
        self.txt_path.Text = get_default_export_path()

        btn_browse = Button()
        btn_browse.Text = "..."
        btn_browse.Location = Point(425, y_pos - 4)
        btn_browse.Size = Size(45, 23)
        btn_browse.Click += self.on_browse

        y_pos += 35

        # Имя файла
        lbl_filename = Label()
        lbl_filename.Text = "Имя файла:"
        lbl_filename.Location = Point(15, y_pos)
        lbl_filename.Width = 120

        self.txt_filename = TextBox()
        self.txt_filename.Location = Point(140, y_pos - 3)
        self.txt_filename.Width = 330
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.txt_filename.Text = "export_{}".format(timestamp)

        y_pos += 40

        # === Группа настроек отображения ===
        grp_display = GroupBox()
        grp_display.Text = "Настройки отображения"
        grp_display.Location = Point(15, y_pos)
        grp_display.Size = Size(455, 90)

        # Скрыть аннотации
        self.chk_hide_annotations = CheckBox()
        self.chk_hide_annotations.Text = "Скрыть аннотации (размеры, текст, оси, уровни, разрезы)"
        self.chk_hide_annotations.Location = Point(15, 22)
        self.chk_hide_annotations.Width = 420
        self.chk_hide_annotations.Checked = True

        # Белый фон
        self.chk_white_bg = CheckBox()
        self.chk_white_bg.Text = "Белый фон (для 3D видов)"
        self.chk_white_bg.Location = Point(15, 48)
        self.chk_white_bg.Width = 300
        self.chk_white_bg.Checked = True

        grp_display.Controls.Add(self.chk_hide_annotations)
        grp_display.Controls.Add(self.chk_white_bg)

        y_pos += 100

        # Статус
        self.lbl_status = Label()
        self.lbl_status.Text = ""
        self.lbl_status.Location = Point(15, y_pos)
        self.lbl_status.Size = Size(455, 25)
        self.lbl_status.ForeColor = Color.Gray

        y_pos += 30

        # Кнопки
        self.btn_export = Button()
        self.btn_export.Text = "Экспортировать"
        self.btn_export.Location = Point(260, y_pos)
        self.btn_export.Size = Size(120, 35)
        self.btn_export.Click += self.on_export

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(390, y_pos)
        btn_cancel.Size = Size(80, 35)
        btn_cancel.Click += self.on_cancel

        # Добавляем контролы
        self.Controls.Add(lbl_view)
        self.Controls.Add(self.cmb_views)
        self.Controls.Add(lbl_resolution)
        self.Controls.Add(self.cmb_resolution)
        self.Controls.Add(lbl_format)
        self.Controls.Add(self.cmb_format)
        self.Controls.Add(lbl_path)
        self.Controls.Add(self.txt_path)
        self.Controls.Add(btn_browse)
        self.Controls.Add(lbl_filename)
        self.Controls.Add(self.txt_filename)
        self.Controls.Add(grp_display)
        self.Controls.Add(self.lbl_status)
        self.Controls.Add(self.btn_export)
        self.Controls.Add(btn_cancel)

    def on_browse(self, sender, args):
        """Выбрать папку для сохранения."""
        from System.Windows.Forms import FolderBrowserDialog
        dialog = FolderBrowserDialog()
        dialog.SelectedPath = self.txt_path.Text
        if dialog.ShowDialog() == DialogResult.OK:
            self.txt_path.Text = dialog.SelectedPath

    def on_export(self, sender, args):
        """Экспортировать вид."""
        if self.cmb_views.SelectedIndex < 0:
            show_warning("Выбор", "Выберите вид для экспорта")
            return

        view = self.views[self.cmb_views.SelectedIndex]

        # Получаем разрешение
        resolution_map = {
            0: 800,
            1: 1280,
            2: 1920,
            3: 2560,
            4: 3840,
            5: 4096,
            6: 5120,
            7: 7680
        }
        pixel_size = resolution_map.get(self.cmb_resolution.SelectedIndex, 4096)

        # Получаем формат
        format_map = {
            0: ImageFileType.PNG,
            1: ImageFileType.JPEGLossless,
            2: ImageFileType.JPEGMedium,
            3: ImageFileType.JPEGSmallest,
            4: ImageFileType.TIFF,
            5: ImageFileType.BMP
        }
        file_type = format_map.get(self.cmb_format.SelectedIndex, ImageFileType.PNG)

        # Путь
        export_dir = self.txt_path.Text
        filename = self.txt_filename.Text

        if not filename:
            show_warning("Имя файла", "Введите имя файла")
            return

        export_path = os.path.join(export_dir, filename)

        hide_annotations = self.chk_hide_annotations.Checked
        white_background = self.chk_white_bg.Checked

        self.lbl_status.Text = "Экспорт..."
        self.lbl_status.ForeColor = Color.Blue
        self.Refresh()

        # Экспортируем
        result_path = export_view_to_image(
            view, export_path, pixel_size, file_type,
            hide_annotations, white_background
        )

        if result_path:
            self.lbl_status.Text = "Сохранено: {}".format(os.path.basename(result_path))
            self.lbl_status.ForeColor = Color.Green

            # Закрываем форму экспорта и открываем предпросмотр
            self.Close()

            # Показываем форму предпросмотра с интеграцией Gemini
            preview_form = ImagePreviewForm(result_path, export_dir)
            preview_form.ShowDialog()
        else:
            self.lbl_status.Text = "Ошибка экспорта"
            self.lbl_status.ForeColor = Color.Red

    def on_cancel(self, sender, args):
        """Отмена."""
        self.Close()


def main():
    """Основная функция."""
    if not doc.PathName:
        show_warning("Проект не сохранён",
                    "Сохраните проект перед экспортом изображений")

    form = ExportViewForm()
    form.ShowDialog()


if __name__ == "__main__":
    main()
