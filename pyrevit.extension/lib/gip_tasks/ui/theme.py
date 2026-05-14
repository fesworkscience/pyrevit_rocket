# -*- coding: utf-8 -*-
try:
    from System import Enum
    from System.Windows import CornerRadius, Setter, Style, Thickness, Trigger
    from System.Windows.Controls import Border, Button, CheckBox, ComboBox, DataGrid, DataGridColumnHeader, DatePicker, GridViewColumnHeader, ListBox, ListView, TextBlock, TextBox
    from System.Windows.Media import ColorConverter, SolidColorBrush, VisualTreeHelper
except Exception:
    Enum = None
    CornerRadius = Setter = Style = Thickness = Trigger = None
    Border = Button = CheckBox = ComboBox = DataGrid = DataGridColumnHeader = DatePicker = GridViewColumnHeader = ListBox = ListView = TextBlock = TextBox = None
    ColorConverter = SolidColorBrush = VisualTreeHelper = None


LIGHT = {
    "window": "#F7F8FC",
    "surface": "#FFFFFF",
    "panel": "#F1F4FF",
    "input": "#FFFFFF",
    "text": "#1F2430",
    "muted": "#6B7280",
    "border": "#E2E5F0",
    "primary": "#3136FF",
    "primary_hover": "#252AE6",
    "secondary": "#5331FF",
    "accent": "#660FD1",
    "secondary_hover": "#EEF2FF",
    "button_text": "#FFFFFF",
    "success": "#16A34A",
    "offline": "#9CA3AF",
    "danger": "#DC2626",
}

DARK = {
    "window": "#111827",
    "surface": "#172033",
    "panel": "#1E1B4B",
    "input": "#0F172A",
    "text": "#E5E7EB",
    "muted": "#A5B4FC",
    "border": "#374151",
    "primary": "#3136FF",
    "primary_hover": "#5331FF",
    "secondary": "#5331FF",
    "accent": "#A78BFA",
    "secondary_hover": "#27324A",
    "button_text": "#FFFFFF",
    "success": "#22C55E",
    "offline": "#9CA3AF",
    "danger": "#F87171",
}


def _brush(hex_value):
    return SolidColorBrush(ColorConverter.ConvertFromString(hex_value))


def _safe_set(obj, name, value):
    try:
        setattr(obj, name, value)
    except Exception:
        pass


def _tag_value(control):
    try:
        return str(control.Tag)
    except Exception:
        return ""


def _should_keep_style(control):
    tag = _tag_value(control).lower()
    return tag in ("keepstyle", "no_theme", "notheme")


def is_dark_theme():
    try:
        from Autodesk.Revit.UI import UIThemeManager
        theme = str(UIThemeManager.CurrentTheme).lower()
        if "dark" in theme:
            return True
        if "light" in theme:
            return False
    except Exception:
        pass
    try:
        from Microsoft.Win32 import Registry
        key = Registry.CurrentUser.OpenSubKey(r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        value = key.GetValue("AppsUseLightTheme")
        return int(value) == 0
    except Exception:
        return False


def _children(element):
    if VisualTreeHelper is None:
        return []
    try:
        count = VisualTreeHelper.GetChildrenCount(element)
        return [VisualTreeHelper.GetChild(element, i) for i in range(count)]
    except Exception:
        return []


def _walk(element):
    yield element
    for child in _children(element):
        for item in _walk(child):
            yield item


def _button_style(brushes, primary=False):
    if Style is None or Setter is None or Trigger is None or Button is None:
        return None
    style = Style(Button)
    if primary:
        background = brushes["primary"]
        foreground = brushes["button_text"]
        border = brushes["primary_hover"]
        hover_background = brushes["primary_hover"]
        hover_border = brushes["accent"]
        pressed_background = brushes["accent"]
    else:
        background = brushes["surface"]
        foreground = brushes["text"]
        border = brushes["border"]
        hover_background = brushes["secondary_hover"]
        hover_border = brushes["primary"]
        pressed_background = brushes["panel"]
    style.Setters.Add(Setter(Button.BackgroundProperty, background))
    style.Setters.Add(Setter(Button.ForegroundProperty, foreground))
    style.Setters.Add(Setter(Button.BorderBrushProperty, border))
    style.Setters.Add(Setter(Button.PaddingProperty, Thickness(10, 3, 10, 3)))
    style.Setters.Add(Setter(Button.MinHeightProperty, 28))
    hover = Trigger()
    hover.Property = Button.IsMouseOverProperty
    hover.Value = True
    hover.Setters.Add(Setter(Button.BackgroundProperty, hover_background))
    hover.Setters.Add(Setter(Button.BorderBrushProperty, hover_border))
    style.Triggers.Add(hover)
    pressed = Trigger()
    pressed.Property = Button.IsPressedProperty
    pressed.Value = True
    pressed.Setters.Add(Setter(Button.BackgroundProperty, pressed_background))
    style.Triggers.Add(pressed)
    return style


def _style_tree(window, brushes, primary_button_style=None, secondary_button_style=None):
    for control in _walk(window):
        if _should_keep_style(control):
            continue
        if Border is not None and isinstance(control, Border):
            _safe_set(control, "Background", brushes["surface"])
            _safe_set(control, "BorderBrush", brushes["border"])
            _safe_set(control, "CornerRadius", CornerRadius(8))
        elif Button is not None and isinstance(control, Button):
            name = getattr(control, "Name", "")
            style = primary_button_style if name in ("SendButton", "NewTaskButton", "PrimaryButton") else secondary_button_style
            if style is not None:
                _safe_set(control, "Style", style)
            if name == "FindMarkerButton":
                _safe_set(control, "Foreground", brushes["accent"])
                _safe_set(control, "BorderBrush", brushes["accent"])
                _safe_set(control, "Background", brushes["surface"])
        elif TextBox is not None and isinstance(control, TextBox):
            _safe_set(control, "Background", brushes["input"])
            _safe_set(control, "Foreground", brushes["text"])
            _safe_set(control, "BorderBrush", brushes["border"])
            _safe_set(control, "Padding", Thickness(6, 3, 6, 3))
        elif ComboBox is not None and isinstance(control, ComboBox):
            _safe_set(control, "Background", brushes["input"])
            _safe_set(control, "Foreground", brushes["text"])
            _safe_set(control, "BorderBrush", brushes["border"])
            _safe_set(control, "Padding", Thickness(4, 2, 4, 2))
            _safe_set(control, "MinHeight", 26)
        elif DatePicker is not None and isinstance(control, DatePicker):
            _safe_set(control, "Background", brushes["input"])
            _safe_set(control, "Foreground", brushes["text"])
            _safe_set(control, "BorderBrush", brushes["border"])
            _safe_set(control, "MinHeight", 28)
        elif ListBox is not None and isinstance(control, ListBox):
            _safe_set(control, "Background", brushes["surface"])
            _safe_set(control, "Foreground", brushes["text"])
            _safe_set(control, "BorderBrush", brushes["border"])
        elif ListView is not None and isinstance(control, ListView):
            _safe_set(control, "Background", brushes["surface"])
            _safe_set(control, "Foreground", brushes["text"])
            _safe_set(control, "BorderBrush", brushes["border"])
        elif DataGrid is not None and isinstance(control, DataGrid):
            _safe_set(control, "Background", brushes["surface"])
            _safe_set(control, "Foreground", brushes["text"])
            _safe_set(control, "BorderBrush", brushes["border"])
            _safe_set(control, "RowBackground", brushes["surface"])
            _safe_set(control, "AlternatingRowBackground", brushes["window"])
        elif GridViewColumnHeader is not None and isinstance(control, GridViewColumnHeader):
            _safe_set(control, "Background", brushes["panel"])
            _safe_set(control, "Foreground", brushes["text"])
            _safe_set(control, "BorderBrush", brushes["border"])
            _safe_set(control, "Padding", Thickness(6, 4, 6, 4))
        elif DataGridColumnHeader is not None and isinstance(control, DataGridColumnHeader):
            _safe_set(control, "Background", brushes["panel"])
            _safe_set(control, "Foreground", brushes["text"])
            _safe_set(control, "BorderBrush", brushes["border"])
            _safe_set(control, "Padding", Thickness(6, 4, 6, 4))
            _safe_set(control, "MinHeight", 28)
        elif CheckBox is not None and isinstance(control, CheckBox):
            _safe_set(control, "Foreground", brushes["text"])
        elif TextBlock is not None and isinstance(control, TextBlock):
            pass

    for name in ("ProjectText", "StatusText", "MarkerIdText", "ModeBadge", "ConnectionStatusText"):
        try:
            getattr(window, name).Foreground = brushes["accent" if name in ("ProjectText", "ModeBadge") else "muted"]
        except Exception:
            pass


def apply_theme(window):
    if SolidColorBrush is None:
        return
    palette = DARK if is_dark_theme() else LIGHT
    brushes = {}
    for key, value in palette.items():
        brushes[key] = _brush(value)

    _safe_set(window, "Background", brushes["window"])
    _safe_set(window, "Foreground", brushes["text"])

    try:
        window.Resources["GIP_WindowBrush"] = brushes["window"]
        window.Resources["GIP_SurfaceBrush"] = brushes["surface"]
        window.Resources["GIP_PanelBrush"] = brushes["panel"]
        window.Resources["GIP_InputBrush"] = brushes["input"]
        window.Resources["GIP_TextBrush"] = brushes["text"]
        window.Resources["GIP_MutedBrush"] = brushes["muted"]
        window.Resources["GIP_BorderBrush"] = brushes["border"]
        window.Resources["GIP_PrimaryBrush"] = brushes["primary"]
        window.Resources["GIP_SecondaryBrush"] = brushes["secondary"]
        window.Resources["GIP_AccentBrush"] = brushes["accent"]
    except Exception:
        pass

    primary_button_style = _button_style(brushes, primary=True)
    secondary_button_style = _button_style(brushes, primary=False)
    _style_tree(window, brushes, primary_button_style, secondary_button_style)
    try:
        def _loaded(sender, args):
            _style_tree(window, brushes, primary_button_style, secondary_button_style)
        window.Loaded += _loaded
        window._gip_theme_loaded_handler = _loaded
    except Exception:
        pass


def set_connection_state(dot, text):
    if SolidColorBrush is None or dot is None:
        return
    lowered = (text or "").lower()
    color = LIGHT["success"]
    if u"недоступен" in lowered:
        color = LIGHT["danger"]
    elif u"локальный" in lowered or u"кэш" in lowered:
        color = LIGHT["offline"]
    try:
        dot.Background = _brush(color)
    except Exception:
        pass


def set_quick_tab_state(button, active):
    if SolidColorBrush is None or button is None:
        return
    try:
        button.Foreground = _brush(LIGHT["primary"] if active else LIGHT["muted"])
        button.BorderBrush = _brush(LIGHT["primary"] if active else "#00FFFFFF")
        button.Background = _brush("#F4F5FF" if active else "#00FFFFFF")
        button.BorderThickness = Thickness(0, 0, 0, 2) if active else Thickness(0)
    except Exception:
        pass


def set_chip_state(button, active):
    if SolidColorBrush is None or button is None:
        return
    try:
        if not button.IsEnabled:
            button.Foreground = _brush("#A0A7B5")
            button.BorderBrush = _brush("#E5E7EB")
            button.Background = _brush("#F5F6FA")
            return
        button.Foreground = _brush(LIGHT["primary"] if active else LIGHT["text"])
        button.BorderBrush = _brush(LIGHT["primary"] if active else LIGHT["border"])
        button.Background = _brush("#F4F5FF" if active else LIGHT["surface"])
    except Exception:
        pass


def set_badge(border, text_block, background, foreground):
    if SolidColorBrush is None:
        return
    try:
        if border is not None:
            border.Background = _brush(background)
            border.BorderBrush = _brush(background)
        if text_block is not None:
            text_block.Foreground = _brush(foreground)
    except Exception:
        pass
