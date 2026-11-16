# Шрифты тем

- `Inter-VariableFont_opsz,wght.ttf`, `Inter-Italic-VariableFont_opsz,wght.ttf` —
  Inter (SIL Open Font License 1.1, https://rsms.me/inter/). Основной текстовый
  шрифт всех тем; в matplotlib регистрируется автоматически из этой папки.
  Для корректного просмотра pptx на чужих машинах Inter должен быть установлен
  в систему (или сохранён внутри презентации: Файл → Параметры → Сохранение →
  «Внедрить шрифты»).
  Variable-шрифт: PowerPoint выбирает вес сам; matplotlib видит только Regular —
  выделять размером/цветом, не weight=bold.
- Тема `academic` использует системный Georgia для заголовков (есть в Windows и
  macOS; в Linux-контейнере при отсутствии — фолбэк на serif, не критично).
  Отдельно не поставляется (проприетарный).

Перед установкой на новые машины: Inter — свободно (OFL), Georgia — системный.
