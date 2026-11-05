# -*- coding: utf-8 -*-
"""Генерує Lab2_presentation.pptx і СЦЕНАРІЙ_ВИСТУПУ.md з одних і тих самих нотаток."""
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))

INK = RGBColor(0x0F, 0x2A, 0x33)
TEAL = RGBColor(0x1F, 0x7A, 0x8C)
AMBER = RGBColor(0xF2, 0xA5, 0x41)
LIGHT = RGBColor(0xEA, 0xF2, 0xF3)
TEXT = RGBColor(0x1B, 0x2B, 0x30)
MUTED = RGBColor(0x5B, 0x6B, 0x70)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xBF, 0xDD, 0xE3)

HEAD = "Cambria"
BODY = "Calibri"
CODE = "Consolas"

REPO = "github.com/DanyloKuch/gemini-study-helper"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
NOTES = []


def new_slide(bg=WHITE, notes=""):
    slide = prs.slides.add_slide(BLANK)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = bg
    slide.notes_slide.notes_text_frame.text = notes
    NOTES.append(notes)
    return slide


def text(slide, x, y, w, h, content, size=16, bold=False, color=TEXT, font=BODY,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, name=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = content if isinstance(content, list) else [content]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = font
        run.font.color.rgb = color
    return box


def shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=1.5, radius=None, name=None):
    s = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    return s


def card(slide, x, y, w, h, fill=LIGHT, name=None):
    return shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill, radius=0.06, name=name)


def badge(slide, x, y, d, label, fill=AMBER, color=INK, size=16):
    s = shape(slide, MSO_SHAPE.OVAL, x, y, d, d, fill=fill)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.name = BODY
    r.font.color.rgb = color
    return s


def title(slide, content, color=INK):
    text(slide, 0.7, 0.5, 11.9, 0.9, content, size=38, bold=True, color=color, font=HEAD,
         anchor=MSO_ANCHOR.MIDDLE, name="Title")


# ---------------------------------------------------------------- 1. Титул
s = new_slide(INK, (
    "Добрий день. Мене звати Данило Кучерук. Моя робота називається Study Helper. "
    "Це помічник для навчання: ви даєте йому конспект, а він за допомогою штучного інтелекту "
    "Gemini робить короткий переказ і запитання для самоперевірки."))
for i, d in enumerate((6.4, 4.6, 2.8)):
    shape(s, MSO_SHAPE.OVAL, 10.1 - d / 2, 3.75 - d / 2, d, d,
          line=TEAL, line_w=1.25, name="Orbit %d" % (i + 1))
text(s, 0.9, 2.2, 5.8, 1.3, "Study Helper", size=60, bold=True, color=WHITE, font=HEAD, name="Title")
text(s, 0.9, 3.65, 5.8, 1.2, ["AI-помічник для конспекту", "на базі Gemini API"], size=26, color=SOFT,
     name="Subtitle")
text(s, 0.9, 6.2, 9.0, 0.6, "Кучерук Данило, ТТП-41  ·  Хмарні обчислення, лабораторна 2", size=16,
     color=SOFT, name="Author")

# ---------------------------------------------------------------- 2. Суть
s = new_slide(WHITE, (
    "Ідея проста. На вході текст конспекту. Всередині працює Gemini. На виході короткий переказ "
    "і запитання з відповідями. Це корисно для підготовки до заліків, для швидкого підсумку довгих "
    "документів, а також як основа для навчальних платформ і чат-ботів."))
title(s, "Суть проєкту")
steps = [("1", "Вхід", "Текст конспекту у звичайному файлі"),
         ("2", "Обробка", "Gemini аналізує текст і формує відповідь за заданою схемою"),
         ("3", "Вихід", "Стислий переказ і запитання для самоперевірки з відповідями")]
for i, (num, head, desc) in enumerate(steps):
    x = 0.7 + i * (3.8 + 0.25)
    card(s, x, 1.7, 3.8, 2.5, name="Card %s" % num)
    badge(s, x + 0.3, 1.95, 0.6, num)
    text(s, x + 1.1, 1.95, 2.5, 0.6, head, size=22, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.3, 2.85, 3.2, 1.2, desc, size=16, color=TEXT)
text(s, 0.7, 4.65, 11.9, 0.5, "Де це можна використати", size=22, bold=True, color=INK, font=HEAD)
uses = ["Підготовка до заліків і екзаменів", "Стислий підсумок документів і статей",
        "Основа для навчальних платформ і чат-ботів"]
for i, u in enumerate(uses):
    x = 0.7 + i * (3.8 + 0.25)
    shape(s, MSO_SHAPE.OVAL, x, 5.5, 0.22, 0.22, fill=TEAL)
    text(s, x + 0.4, 5.38, 3.4, 0.9, u, size=16, color=TEXT)

# ---------------------------------------------------------------- 3. API
s = new_slide(WHITE, (
    "Застосунок використовує API Gemini від Google. Головний виклик називається generateContent: "
    "ми надсилаємо текст конспекту і просимо відповісти. Ще є виклик списку моделей, ним я "
    "перевіряв, які моделі доступні. Важлива річ: structured output. Ми задаємо схему, "
    "і модель повертає чіткий JSON, який легко розібрати. Доступ — через ключ у заголовку запиту."))
title(s, "Які API використовуються")
rows = [("POST", AMBER, "/v1beta/models/{model}:generateContent",
         "Основний виклик: аналіз конспекту, генерація переказу і запитань"),
        ("GET", TEAL, "/v1beta/models", "Перевірка доступних моделей під час розробки"),
        ("JSON", INK, "responseSchema", "Structured output: модель повертає гарантовано розбірний JSON")]
for i, (verb, col, endpoint, desc) in enumerate(rows):
    y = 1.65 + i * 1.3
    card(s, 0.7, y, 11.9, 1.1, name="API row %d" % (i + 1))
    b = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.95, y + 0.3, 1.0, 0.5, fill=col, radius=0.25)
    tf = b.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = verb
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.name = BODY
    r.font.color.rgb = INK if col == AMBER else WHITE
    text(s, 2.25, y + 0.14, 5.0, 0.4, endpoint, size=15, bold=True, color=INK, font=CODE)
    text(s, 2.25, y + 0.55, 10.0, 0.4, desc, size=15, color=MUTED)
text(s, 0.7, 5.7, 11.9, 0.9,
     "Авторизація: заголовок x-goog-api-key. Ключ береться зі змінної середовища "
     "GEMINI_API_KEY і не потрапляє в код чи репозиторій.", size=16, color=TEXT)

# ---------------------------------------------------------------- 4. Взаємодія
s = new_slide(WHITE, (
    "Як це все зв'язано. Користувач дає файл. Програма на C# читає його, додає інструкцію і схему "
    "відповіді та надсилає запит у Gemini по HTTPS. Gemini повертає JSON, програма його розбирає "
    "і красиво друкує переказ та запитання."))
title(s, "Як взаємодіють сервіси")
boxes = [("Користувач", "файл конспекту"), ("StudyHelper", "C#, консольний застосунок"),
         ("Gemini API", "Google Cloud"), ("Результат", "переказ і запитання")]
bw, gap, x0, by = 2.5, 0.65, 0.69, 2.4
for i, (head, sub) in enumerate(boxes):
    x = x0 + i * (bw + gap)
    dark = i in (1, 2)
    card(s, x, by, bw, 1.6, fill=INK if dark else LIGHT, name="Node %d" % (i + 1))
    text(s, x + 0.15, by + 0.35, bw - 0.3, 0.5, head, size=20, bold=True,
         color=WHITE if dark else INK, align=PP_ALIGN.CENTER, font=HEAD)
    text(s, x + 0.15, by + 0.95, bw - 0.3, 0.5, sub, size=14,
         color=SOFT if dark else MUTED, align=PP_ALIGN.CENTER)
    if i < 3:
        shape(s, MSO_SHAPE.RIGHT_ARROW, x + bw + 0.08, by + 0.62, gap - 0.16, 0.36, fill=TEAL,
              name="Arrow %d" % (i + 1))
        badge(s, x + bw + gap / 2 - 0.2, by - 0.52, 0.4, str(i + 1), size=14)
legend = ["Програма читає файл, додає інструкцію і схему відповіді",
          "Надсилає HTTPS POST generateContent, ключ у заголовку",
          "Gemini повертає JSON, програма розбирає його і друкує"]
for i, item in enumerate(legend):
    y = 4.7 + i * 0.65
    badge(s, 0.7, y, 0.4, str(i + 1), size=14)
    text(s, 1.3, y, 11.2, 0.4, item, size=17, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- 5. Код
s = new_slide(WHITE, (
    "Ось головний шматок коду. У запиті ми просимо відповідь у форматі JSON за нашою схемою. "
    "Ключ передаємо в заголовку, а не в адресі і не в коді. Також програма обробляє помилки: "
    "невірний ключ, перевищення ліміту, порожню відповідь, і пише зрозуміле повідомлення."))
title(s, "Ключова частина коду")
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.7, 1.65, 7.1, 4.6, fill=INK, radius=0.03, name="Code panel")
code = ['["generationConfig"] = new JsonObject',
        '{',
        '    ["responseMimeType"] = "application/json",',
        '    ["responseSchema"] = BuildSchema()',
        '}',
        '',
        'request.Headers.Add(',
        '    "x-goog-api-key", apiKey);',
        '',
        'using var response =',
        '    await http.SendAsync(request);']
text(s, 1.0, 1.95, 6.6, 4.0, code, size=15, color=WHITE, font=CODE, name="Code")
points = [("Схема відповіді", "responseMimeType і responseSchema змушують модель повертати JSON"),
          ("Безпечний ключ", "Ключ у заголовку із змінної середовища, не в коді"),
          ("Обробка помилок", "429, невірний ключ, порожня відповідь: зрозумілі повідомлення")]
for i, (h, d) in enumerate(points):
    y = 1.65 + i * 1.55
    badge(s, 8.2, y, 0.5, str(i + 1), size=16)
    text(s, 8.95, y - 0.02, 3.7, 0.5, h, size=18, bold=True, color=INK, font=HEAD)
    text(s, 8.95, y + 0.5, 3.7, 0.95, d, size=14, color=MUTED)

# ---------------------------------------------------------------- 6. Демо
s = new_slide(WHITE, (
    "Тепер демонстрація. Спочатку показую конспект про хмарні обчислення, потім запускаю програму. "
    "[Запустити вживу; якщо інтернет або ліміт підведе, клацнути по відео.] Бачимо переказ і "
    "чотири запитання з відповідями, які спираються лише на текст конспекту."))
title(s, "Демонстрація")
# Відео 800x720, вписуємо по висоті рамки 6.0 x 4.6 і центруємо
video_w = 4.6 * 800 / 720
s.shapes.add_movie(os.path.join(HERE, "..", "demo", "demo.mp4"),
                   Inches(0.7 + (6.0 - video_w) / 2), Inches(1.65), Inches(video_w), Inches(4.6),
                   poster_frame_image=os.path.join(HERE, "..", "demo", "demo_poster.png"),
                   mime_type="video/mp4")
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 7.0, 1.65, 5.65, 4.6, fill=INK, radius=0.03, name="Terminal")
term = ["> dotnet run --project StudyHelper --",
        "  samples\\conspect2_cloud.txt 4",
        "",
        "=== ЗАПИТАННЯ ДЛЯ САМОПЕРЕВІРКИ ===",
        "2. Які три основні моделі",
        "   обслуговування існують?",
        "   Відповідь: IaaS, PaaS, SaaS",
        "4. Які два ризики пов'язані",
        "   з хмарними обчисленнями?",
        "   Відповідь: залежність від",
        "   провайдера і конфіденційність"]
text(s, 7.3, 1.95, 5.1, 4.1, term, size=13, color=WHITE, font=CODE, name="Terminal text")

# ---------------------------------------------------------------- 7. Ризики
s = new_slide(WHITE, (
    "Є обмеження. Безкоштовний ліміт запитів, тому на випадок збою є записане відео. "
    "Модель іноді помиляється, тому відповіді треба перевіряти. Текст іде на сервери Google, "
    "тому конфіденційні матеріали не варто. І ключ треба берегти, після демо його перевипускають."))
title(s, "Обмеження та ризики")
risks = [("Ліміти free tier", "Обмеження на частоту запитів і довжину тексту. На випадок збою є записане відео."),
         ("Помилки моделі", "У промпті вимагаємо спиратися лише на конспект, але перевірка залишається за людиною."),
         ("Приватність", "Текст конспекту надсилається на сервери Google. Конфіденційні матеріали не використовувати."),
         ("Безпека ключа", "Ключ лише у змінній середовища, у репозиторій не потрапляє. Після демо його перевипускають.")]
for i, (h, d) in enumerate(risks):
    x = 0.7 + (i % 2) * (5.85 + 0.2)
    y = 1.65 + (i // 2) * (2.25 + 0.25)
    card(s, x, y, 5.85, 2.25, name="Risk %d" % (i + 1))
    shape(s, MSO_SHAPE.OVAL, x + 0.3, y + 0.35, 0.28, 0.28, fill=AMBER)
    text(s, x + 0.75, y + 0.28, 4.8, 0.5, h, size=20, bold=True, color=INK, font=HEAD)
    text(s, x + 0.3, y + 1.0, 5.25, 1.1, d, size=15, color=TEXT)

# ---------------------------------------------------------------- 8. Підсумок
s = new_slide(INK, (
    "Підсумок: застосунок працює, код на C# без зайвих бібліотек, лежить на GitHub. "
    "Далі можна додати веб-інтерфейс і режим запитань по конспекту. Дякую за увагу, "
    "готовий відповісти на запитання."))
title(s, "Підсумок", color=WHITE)
wrap = [("Працює", "Переказ і запитання від Gemini API у вигляді структурованого JSON"),
        ("Код", "C#, REST-запити без сторонніх SDK, ключ у змінній середовища"),
        ("Далі", "Веб-інтерфейс і режим запитань по конспекту")]
for i, (h, d) in enumerate(wrap):
    y = 1.7 + i * 1.1
    text(s, 0.7, y, 2.0, 0.6, h, size=22, bold=True, color=AMBER, font=HEAD)
    text(s, 2.9, y + 0.04, 9.7, 0.9, d, size=18, color=WHITE)
text(s, 0.7, 5.35, 11.9, 0.4, "Код на GitHub", size=16, color=SOFT)
text(s, 0.7, 5.8, 11.9, 0.7, REPO, size=30, bold=True, color=AMBER, font=CODE, name="Repo link")

out = os.path.join(HERE, "Lab2_presentation.pptx")
prs.save(out)
print("saved", out)

# ---------------------------------------------------------------- сценарій виступу
titles = ["Титул", "Суть проєкту", "Які API використовуються", "Як взаємодіють сервіси",
          "Ключова частина коду", "Демонстрація", "Обмеження та ризики", "Підсумок"]
lines = ["# Сценарій виступу (лаба 2): Study Helper", "",
         "Кожен пункт відповідає слайду. Текст можна читати вголос простими словами, поки слайд на екрані.",
         "Приблизний час: 1,5–2 хвилини на слайд, разом до 12 хвилин із демо.", ""]
for i, (t, n) in enumerate(zip(titles, NOTES), start=1):
    lines += ["## Слайд %d. %s" % (i, t), "", n, ""]
script_path = os.path.join(os.path.dirname(HERE), "СЦЕНАРІЙ_ВИСТУПУ.md")
with open(script_path, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print("saved", script_path)
