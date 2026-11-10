# krk_builder.py

Исходный файл: `originals/krk_builder.py`. Модуль Python проекта: функция `build_krk()` (сборка КРК по шаблону), `it16_limits()` (границы по квалитету 16, ГОСТ 25346-2013), `extract_ol_images()` (рисунки из PDF опросного листа). Ниже — полный исходный код без изменений.

```python
"""krk_builder.py — сборка КРК (карты результатов контроля) по шаблону АО «МОСГАЗ»
с автоматическим заключением «Соответствует / Не соответствует» и заливкой.

Использование (в чате проекта Claude запускает сам):
    from krk_builder import build_krk
    build_krk(
        template="КРК_шаблон_с_автозаключением.xlsx",
        out="КРК_Отвод_П90-57x5.xlsx",
        name="Отвод П90-57х5-09Г2С ГОСТ 17375-2001",
        nd="ГОСТ 17375-2001, ГОСТ 17380-2001",
        params=[
            {"text": "Визуальный контроль ...", "ref": "ГОСТ 17380, п. 5.1.1"},    # качественный
            {"text": "D (торец 1) = 57 ±0,6 мм", "min": 56.4, "max": 57.6, "ref": "ГОСТ 17380, п. 5.1.3.1"},
            {"text": "Q ≤ 0,5 мм", "max": 0.5},                                  # односторонний
        ],
        measured={2: 57.1},   # необязательно: № п/п -> измеренное значение
    )
Правило: есть min и/или max -> числовой параметр; нет обоих -> качественный.
Норма не установлена, но параметр измеряемый -> {"text": ..., "numeric": True}: ячейка принимает только
число, а до ввода границ любое число даёт «Не соответствует» (защита от применения КРК-проекта).
Блоки (кран / привод): {"section": "Электропривод ..."} — строка-заголовок без номера.
Рисунок из ОЛ: image="путь.png" — вставляется в окно «Дополнительные идентификационные признаки».
ЗРА: границы по ОЛ (H16 / h16 / IT16/2) — it16_limits(номинал, "H16" | "h16" | "js16").
Рисунки из ОЛ: extract_ol_images("ОЛ.pdf", "img") -> список PNG в порядке позиций.
Все линейные размеры и допуски — в мм (unit по умолчанию "мм"); углы — "°".
В «Измеренном значении» число хранится числом, единица выводится форматом ячейки («57,5 мм»).
(«Соответствует» только при явной записи «Соответствует», иначе «Не соответствует»).
"""
import math
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import PatternFill as _PF

# Допуски IT16 по ГОСТ 25346-2013 (ISO 286-1), мм: (свыше, до включ., IT16)
IT16 = [(0, 3, 0.60), (3, 6, 0.75), (6, 10, 0.90), (10, 18, 1.10), (18, 30, 1.30), (30, 50, 1.60),
        (50, 80, 1.90), (80, 120, 2.20), (120, 180, 2.50), (180, 250, 2.90), (250, 315, 3.20),
        (315, 400, 3.60), (400, 500, 4.00), (500, 630, 4.40), (630, 800, 5.00), (800, 1000, 5.60),
        (1000, 1250, 6.60), (1250, 1600, 7.80), (1600, 2000, 9.20), (2000, 2500, 11.00), (2500, 3150, 13.50)]


def it16(size):
    for lo, hi, t in IT16:
        if lo < size <= hi:
            return t
    raise ValueError(f"размер {size} вне таблицы IT16")


def it16_limits(nom, kind):
    """kind: 'H16' (отверстия, +IT/0), 'h16' (валы/наружные, 0/−IT), 'js16' (прочие, ±IT/2).
    Возвращает (min, max, текст допуска)."""
    t = it16(nom)
    f = lambda x: f"{x:.2f}".rstrip("0").rstrip(".").replace(".", ",")
    if kind == "H16":
        return nom, round(nom + t, 3), f"+{f(t)}/0 мм"
    if kind == "h16":
        return round(nom - t, 3), nom, f"+0/−{f(t)} мм"
    h = math.floor(round(t / 2 * 100, 6)) / 100          # IT16/2, округление в сторону ужесточения до 0,01
    return round(nom - h, 3), round(nom + h, 3), f"±{f(h)} мм"


def extract_ol_images(pdf, outdir):
    """Рисунки типоразмеров из PDF опросного листа -> PNG на белом фоне, в порядке позиций ОЛ
    (сверху вниз, страница за страницей). Широкие полосы (штамп, логотип) отбрасываются."""
    import glob, os, subprocess
    from PIL import Image
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(["pdfimages", "-png", pdf, os.path.join(outdir, "raw")], check=True)
    raw = sorted(glob.glob(os.path.join(outdir, "raw-*.png")))
    out, i = [], 0
    while i < len(raw):
        im = Image.open(raw[i])
        nxt = Image.open(raw[i + 1]) if i + 1 < len(raw) else None
        mask = nxt is not None and nxt.mode in ("L", "1") and nxt.size == im.size and im.mode != "L"
        if im.width <= 2.5 * im.height:
            bg = Image.new("RGB", im.size, "white")
            bg.paste(im.convert("RGB"), mask=nxt.convert("L") if mask else None)
            path = os.path.join(outdir, f"img_{len(out) + 1:02d}.png")
            bg.save(path)
            out.append(path)
        i += 2 if mask else 1
    return out

OK, NOK = "Соответствует", "Не соответствует"
GREEN = PatternFill("solid", start_color="92D050", end_color="92D050")
PINK = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
FIRST, TPL_LAST = 17, 31          # строки параметров в шаблоне
MERGES = (("C", "H"), ("I", "N"), ("O", "Q"), ("R", "W"))
COLS = "BCDEFGHIJKLMNOPQRSTUVW"


def _formula(r):
    i, y, z = f"I{r}", f"Y{r}", f"Z{r}"
    text = (f'IF(OR(ISNUMBER(SEARCH("не соотв",{i})),ISNUMBER(SEARCH("несоотв",{i}))),"{NOK}",'
            f'IF(ISNUMBER(SEARCH("соответ",{i})),"{OK}","{NOK}"))')
    num = (f'IF(NOT(ISNUMBER({i})),"{NOK}",IF(AND(OR({y}="",{i}>={y}-0.000001),'
           f'OR({z}="",{i}<={z}+0.000001)),"{OK}","{NOK}"))')
    return f'=IF(TRIM({i}&"")="","",IF(AND({y}="",{z}=""),{text},{num}))'


def _row_height(text):
    lines = sum(max(1, math.ceil(len(part) / 40)) for part in str(text).split("\n"))
    return max(30, 12.5 * lines + 8)


def build_krk(template, out, name, nd, params, serial="", supplier="", extra="-", measured=None,
              image=None, supplier_label=None):
    measured = measured or {}
    wb = load_workbook(template)
    ws = wb.active
    ws.conditional_formatting = ConditionalFormattingList()
    ws.data_validations.dataValidation = []

    # шапка
    ws["G5"], ws["G7"], ws["G8"], ws["G9"] = name, serial, supplier, nd
    ws["M14"] = None if image else extra
    if supplier_label:
        ws["B8"] = supplier_label

    n = len(params)
    last = FIRST + n - 1
    # снять объединения в строках параметров шаблона
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= FIRST:
            ws.unmerge_cells(str(m))
    # лишние строки шаблона удалить
    if last < TPL_LAST:
        ws.delete_rows(last + 1, TPL_LAST - last)
    # эталонный стиль строки = строка 17 шаблона
    style = {c: ws[f"{c}{FIRST}"]._style for c in COLS}
    for k, r in enumerate(range(FIRST, last + 1)):
        for c in COLS:
            ws[f"{c}{r}"]._style = copy(style[c])
        if "section" in params[k]:
            ws.merge_cells(f"B{r}:W{r}")
        else:
            for a, b in MERGES:
                ws.merge_cells(f"{a}{r}:{b}{r}")

    # служебные столбцы вне области печати
    heads = (("X", "Справочно", 40), ("Y", "Мин (для формулы)", 11),
             ("Z", "Макс (для формулы)", 11), ("AA", "Ед. ввода", 9))
    for c, t, w in heads:
        ws[f"{c}15"] = t
        ws[f"{c}15"].font = Font(name="Calibri", size=9, bold=True, color="7F7F7F")
        ws[f"{c}15"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c].width = w
    ws["Y15"].comment = Comment("Границы допуска из графы «Проверяемый параметр». Пусто в обеих "
                                "ячейках = качественный параметр.", "ОТК")

    dv_num = DataValidation(type="decimal", operator="between", formula1="-1000000", formula2="1000000",
                            showErrorMessage=True, errorTitle="Нужно число",
                            error="Введите измеренное значение числом (например 159,4).")
    dv_txt = DataValidation(type="list", formula1=f'"{OK},{NOK}"', showErrorMessage=False,
                            showInputMessage=True, promptTitle="Качественный параметр",
                            prompt="«Соответствует» или «Не соответствует» (можно дописать пояснение).")
    ws.add_data_validation(dv_num)
    ws.add_data_validation(dv_txt)
    gray = Font(name="Calibri", size=9, color="7F7F7F")

    num = 0
    for k, p in enumerate(params):
        r = FIRST + k
        if "section" in p:  # заголовок блока (кран / привод): без номера и без формулы
            c = ws[f"B{r}"]
            c.value = p["section"]
            c.font = Font(name="Calibri", size=9, bold=True)
            c.fill = _PF("solid", start_color="F2F2F2", end_color="F2F2F2")
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.row_dimensions[r].height = 20.1
            continue
        num += 1
        lo, hi = p.get("min"), p.get("max")
        numeric = lo is not None or hi is not None or p.get("numeric", False)
        ws[f"B{r}"] = num
        ws[f"C{r}"] = p["text"]
        ws[f"I{r}"] = measured.get(num)
        ws[f"O{r}"] = _formula(r)
        ws[f"R{r}"] = None                      # «Примечание» оставляем пустым
        ws[f"X{r}"] = p.get("ref")
        ws[f"Y{r}"], ws[f"Z{r}"] = lo, hi
        ws[f"AA{r}"] = p.get("unit", "мм" if numeric else "текст")
        for c in ("X", "Y", "Z", "AA"):
            ws[f"{c}{r}"].font = gray
            ws[f"{c}{r}"].alignment = Alignment(horizontal="left" if c == "X" else "center",
                                                vertical="top", wrap_text=True)
        for c in ("I", "O"):
            ws[f"{c}{r}"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        (dv_num if numeric else dv_txt).add(f"I{r}")
        if numeric:  # число остаётся числом, единица только отображается: «57,5 мм»
            u = p.get("unit", "мм")
            ws[f"I{r}"].number_format = f'General"{u}"' if u == "°" else f'General" {u}"'
        ws.row_dimensions[r].height = _row_height(p["text"])

    rng = f"O{FIRST}:Q{last}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{OK}"'], fill=GREEN))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{NOK}"'], fill=PINK,
                                                  font=Font(color="9C0006")))
    if image:  # рисунок из опросного листа — в окно «Дополнительные идентификационные признаки»
        pic = XLImage(image)
        ws.row_dimensions[14].height = 165.75           # окно под рисунок, как в образце КРК крана
        box_h = 165.75 * 96 / 72 - 10
        scale = min(box_h / pic.height, 380 / pic.width)
        pic.width, pic.height = int(pic.width * scale), int(pic.height * scale)
        ws.add_image(pic, "O14")
    ws.print_area = f"B1:W{last}"
    wb.save(out)
    return out
```
