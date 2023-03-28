import pandas as pd
import openpyxl
import json


# у openpyxl индексация ячеек с 1, у pandas с 0

wb = openpyxl.load_workbook('Raspisanie_2_semestr.xlsx')

ws = wb.worksheets[5]

merged_ranges = ws.merged_cells.ranges


# получение индекса строки левой верхней ячейки соединенной группы
# (если текущая ячейка является левой верхней в этой группе, то ворнет её индекс)
def findLeftTopIndex(row_idx, col_idx):
    # поиск объединенной группы, которой принадлежит заданная ячейка
    for merged_range in merged_ranges:
        if row_idx in range(merged_range.min_row, merged_range.max_row + 1) and col_idx in range(merged_range.min_col, merged_range.max_col + 1):
            left_idx = merged_range.min_row
            return left_idx - 1


df = pd.read_excel("Raspisanie_2_semestr.xlsx", sheet_name=5, header=None)

# заполнение колонки дней
df[0] = df[0].fillna(method="ffill")

to_json = {}

# range(2, len(df), 3) - для перебора всех нужных колонок
for j in [2, 5, 8, 11]:
    numerator = {}
    denominator = {}

    for i in [i for i in range(len(df[1])) if str(df[1][i]) != 'nan']:
        ind = findLeftTopIndex(i + 2, j + 1)  # индекс для текущей ячейки с парой
        place_ind = findLeftTopIndex(i + 2, j + 3)  # индекс для текущей ячейки с аудиторией

        day = df[0][i]
        time = df[1][i]

        if str(df[j][i]) != 'nan':
            if day not in numerator:
                numerator.update({day: [{"time": time, "lesson": ' '.join(df[j][i].split()), "place": ' '.join(str(df[j + 2][i]).split())}]})
            else:
                numerator.update({day: [*numerator.get(day), {"time": time, "lesson": ' '.join(df[j][i].split()), "place": ' '.join(str(df[j + 2][i]).split())}]})
        if str(df[j][ind]) != 'nan':
            if day not in denominator:
                denominator.update({day: [{"time": time, "lesson": ' '.join(df[j][ind].split()), "place": ' '.join(str(df[j + 2][place_ind if place_ind is not None else i + 1]).split())}]})
            else:
                denominator.update({day: [*denominator.get(day), {"time": time, "lesson": ' '.join(df[j][ind].split()), "place": ' '.join(str(df[j + 2][place_ind if place_ind is not None else i + 1]).split())}]})

    to_json.update({df[j][1]: {"numerator": [numerator], "denominator": [denominator]}})

with open("Groups.json", "w", encoding="utf-8") as f:
    json.dump(to_json, f, ensure_ascii=False, indent=4)
