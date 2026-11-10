"""
krk_tools.py — оформление КРК после внесения результатов.

Использование (после того как значения и заключения записаны):
    from krk_tools import finalize_krk
    finalize_krk("КРК_....xlsx")                # на месте
    finalize_krk("вход.xlsx", "выход.xlsx")     # в новый файл

Что делает:
1. Удаляет пустые строки таблицы (есть только № п/п, нет параметра,
   значения и заключения), со сдвигом объединённых ячеек.
2. Перенумеровывает № п/п по порядку (строки-заголовки блоков не трогает).
3. Подбирает высоту строк таблицы под текст (Excel не делает этого сам
   для объединённых ячеек).
4. Выставляет область печати B1:W<последняя строка>.
"""
import math
import sys
import openpyxl
from openpyxl.utils import get_column_letter

FIRST_ROW = 17          # первая строка таблицы параметров
TEXT_COLS = range(3, 24)  # C..W — всё, кроме № п/п
LINE_FACTOR = 1.35      # высота строки текста = кегль * коэффициент
PADDING = 6             # запас по высоте, пт
MIN_HEIGHT = 15         # минимальная высота строки, пт


def _col_width(ws, idx):
    letter = get_column_letter(idx)
    for dim in ws.column_dimensions.values():
        if dim.min and dim.max and dim.min <= idx <= dim.max and dim.width:
            return dim.width
    dim = ws.column_dimensions.get(letter)
    if dim is not None and dim.width:
        return dim.width
    return ws.sheet_format.defaultColWidth or 8.43


def _area(ws, row, col):
    """Границы объединённой области, в которую входит ячейка."""
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return m.min_col, m.max_col, m.max_row - m.min_row + 1
    return col, col, 1


def _lines(text, width_units, font_size):
    px = width_units * 7 + 5 - 6                  # ширина области в пикселях минус поля
    char_px = font_size * 96 / 72 * 0.62          # средняя ширина символа
    per_line = max(1, int(px / char_px))
    total = 0
    for part in str(text).split("\n"):
        total += max(1, math.ceil(len(part) / per_line))
    return total


def _is_empty(ws, r):
    return all(ws.cell(r, c).value in (None, "") for c in TEXT_COLS)


def _delete_row(ws, r):
    heights = {i: ws.row_dimensions[i].height for i in range(r, ws.max_row + 2)}
    for m in list(ws.merged_cells.ranges):
        if m.min_row == r and m.max_row == r:
            ws.merged_cells.remove(m)
    ws.delete_rows(r)
    for m in ws.merged_cells.ranges:
        if m.min_row > r:
            m.shift(0, -1)
    for i in range(r, ws.max_row + 1):
        ws.row_dimensions[i].height = heights.get(i + 1)


def finalize_krk(path_in, path_out=None, first_row=FIRST_ROW):
    wb = openpyxl.load_workbook(path_in)
    ws = wb.worksheets[0]

    # 1. удалить пустые строки таблицы (снизу вверх)
    for r in range(ws.max_row, first_row - 1, -1):
        if _is_empty(ws, r):
            _delete_row(ws, r)
    last = first_row - 1
    for r in range(first_row, ws.max_row + 1):
        if not _is_empty(ws, r) or ws.cell(r, 2).value not in (None, ""):
            last = r

    # 2. перенумеровать № п/п
    n = 0
    for r in range(first_row, last + 1):
        b = ws.cell(r, 2).value
        if b is None or str(b).strip().isdigit():
            if not _is_empty(ws, r):
                n += 1
                ws.cell(r, 2).value = str(n)

    # 3. высота строк
    for r in range(first_row, last + 1):
        need = 1
        size = 9
        for c in TEXT_COLS:
            v = ws.cell(r, c).value
            if v in (None, ""):
                continue
            c1, c2, span = _area(ws, r, c)
            if span > 1:
                continue
            width = sum(_col_width(ws, i) for i in range(c1, c2 + 1))
            size = ws.cell(r, c).font.sz or 9
            need = max(need, _lines(v, width, size))
        ws.row_dimensions[r].height = max(MIN_HEIGHT, round(need * size * LINE_FACTOR + PADDING, 1))

    # 4. область печати
    ws.print_area = f"B1:W{last}"

    wb.save(path_out or path_in)
    return last


if __name__ == "__main__":
    finalize_krk(*sys.argv[1:])
