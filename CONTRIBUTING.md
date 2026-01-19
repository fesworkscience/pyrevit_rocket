# Руководство по участию в проекте

Спасибо за интерес к CPSK Tools! Мы рады любому вкладу в проект.

## Как внести вклад

### Сообщить о баге

1. Проверьте, что баг ещё не зарегистрирован в [Issues](https://github.com/fesworkscience/pyrevit_rocket/issues)
2. Создайте новый Issue с тегом `bug`
3. Опишите проблему максимально подробно:
   - Версия Revit
   - Версия pyRevit
   - Шаги для воспроизведения
   - Ожидаемое и фактическое поведение
   - Скриншоты (если применимо)

### Предложить улучшение

1. Создайте Issue с тегом `enhancement`
2. Опишите предлагаемую функциональность
3. Объясните, какую проблему это решает

### Внести код

```bash
# 1. Форк репозитория на GitHub

# 2. Клонируйте свой форк
git clone https://github.com/YOUR_USERNAME/pyrevit_rocket.git
cd pyrevit_rocket

# 3. Создайте ветку для изменений
git checkout -b feature/my-feature

# 4. Внесите изменения...

# 5. ОБЯЗАТЕЛЬНО: проверьте код чекером
python pyrevit.extension/lib/pyrevit_checker.py pyrevit.extension/CPSK.tab/ПАНЕЛЬ/КНОПКА/script.py

# 6. Закоммитьте изменения
git add .
git commit -m "feat: описание изменений"

# 7. Запушьте ветку
git push origin feature/my-feature

# 8. Создайте Pull Request на GitHub
```

## Требования к коду

### IronPython 2.7 совместимость

pyRevit использует IronPython 2.7, поэтому **запрещено**:

| Запрещено | Использовать |
|-----------|--------------|
| f-строки `f"text {x}"` | `"text {}".format(x)` |
| Walrus operator `:=` | Обычное присваивание |
| `open(encoding=)` | `codecs.open(path, 'r', 'utf-8')` |
| Type hints `def foo(x: int)` | `def foo(x)` |
| `async/await` | Синхронный код |

### Обязательные проверки

1. **Авторизация** — все кнопки (кроме Login) должны проверять `require_auth()`
2. **Уведомления** — использовать `cpsk_notify` вместо `MessageBox.Show`
3. **Обработка ошибок** — все `except` блоки должны уведомлять пользователя
4. **Чекер** — код должен проходить `pyrevit_checker.py` без ошибок

### Стиль коммитов

Используйте [Conventional Commits](https://www.conventionalcommits.org/):

- `feat:` — новая функциональность
- `fix:` — исправление бага
- `docs:` — изменения документации
- `refactor:` — рефакторинг кода
- `test:` — добавление тестов
- `chore:` — прочие изменения

## Pull Request

1. Заполните описание PR
2. Свяжите с Issue (если есть)
3. Убедитесь, что код проходит чекер
4. Дождитесь ревью

## Вопросы?

- [Telegram @tehzak2](https://t.me/tehzak2)
- [GitHub Issues](https://github.com/fesworkscience/pyrevit_rocket/issues)
