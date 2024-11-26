import pandas as pd

# Открытие файла
file_path = "input.xlsx"  # Укажите путь к вашему файлу Excel
data = pd.read_excel(file_path)

# Названия колонок
columns = ["Дата", "Аналитика", "Аналитика (Дебит)", "Аналитика (Кредит)", "Счет дебита", "Сумма дебита", "Счет кредита", "Сумма кредита"]

# Фильтрация данных для страницы "Приход"
prihod = data[
    (data["Счет дебита"].astype(str).str.startswith("10")) &
    (~data["Счет кредита"].astype(str).str.startswith("10"))
]
prihod = prihod.drop(columns=["Сумма кредита"])  # Удаляем "Сумма кредита" из "Приход"

# Фильтрация данных для страницы "Расход"
rashod = data[
    (~data["Счет дебита"].astype(str).str.startswith("10")) &
    (data["Счет кредита"].astype(str).str.startswith("10"))
]
rashod = rashod.drop(columns=["Сумма дебита"])  # Удаляем "Сумма дебита" из "Расход"

# Функция для разделения текста по колонкам
def split_text_to_columns(row, column_name, num_columns):
    values = str(row[column_name]).split('\n')  # Разделение текста по абзацам
    return values[:num_columns] + [''] * (num_columns - len(values))  # Заполняем пустыми строками, если текстов меньше

# Максимальное количество строк, которые нужно перенести в отдельные колонки
max_split = 3

# Обработка аналитики для таблиц
def process_analytics(dataframe):
    new_columns = []
    for col in ["Аналитика", "Аналитика (Дебит)", "Аналитика (Кредит)"]:
        split_data = dataframe.apply(split_text_to_columns, axis=1, column_name=col, num_columns=max_split)
        split_df = pd.DataFrame(split_data.tolist(), columns=[f"{col} {i+1}" for i in range(max_split)])
        new_columns.extend(split_df.columns)
        dataframe = pd.concat([dataframe.reset_index(drop=True), split_df], axis=1)

    dataframe = dataframe.drop(columns=["Аналитика", "Аналитика (Дебит)", "Аналитика (Кредит)"])
    return dataframe, new_columns

# Обработка аналитики для "Приход" и "Расход"
prihod, prihod_columns = process_analytics(prihod)
rashod, rashod_columns = process_analytics(rashod)

# Расположение колонок для "Приход" и "Расход"
prihod = prihod[
    ["Дата"] +
    [col for col in prihod_columns if "Аналитика" in col and "Дебит" not in col and "Кредит" not in col] +
    [col for col in prihod_columns if "Дебит" in col] +
    [col for col in prihod_columns if "Кредит" in col] +
    ["Счет дебита", "Сумма дебита", "Счет кредита"]
]

rashod = rashod[
    ["Дата"] +
    [col for col in rashod_columns if "Аналитика" in col and "Дебит" not in col and "Кредит" not in col] +
    [col for col in rashod_columns if "Дебит" in col] +
    [col for col in rashod_columns if "Кредит" in col] +
    ["Счет дебита", "Счет кредита", "Сумма кредита"]
]

# Преобразование даты и добавление колонки с названием месяца
prihod["Месяц"] = pd.to_datetime(prihod["Дата"], dayfirst=True, errors="coerce").dt.month_name()

# Проверка на строки, которые не удалось преобразовать
if prihod["Месяц"].isna().any():
    print("Не удалось преобразовать следующие даты:")
    print(prihod[prihod["Месяц"].isna()])

# Создание классификации доходов
classification = (
    prihod.groupby("Аналитика (Дебит) 3", as_index=False)
    .agg({"Сумма дебита": "sum"})
    .rename(columns={"Аналитика (Дебит) 3": "Статья доходов", "Сумма дебита": "Итоговая сумма"})
)

# Добавление месячных колонок
month_columns = prihod.pivot_table(
    index="Аналитика (Дебит) 3",
    columns="Месяц",
    values="Сумма дебита",
    aggfunc="sum",
    fill_value=0
).reset_index()

classification = classification.merge(month_columns, how="left", left_on="Статья доходов", right_on="Аналитика (Дебит) 3").drop(columns=["Аналитика (Дебит) 3"])

# Обработка контрагентов для "Реализация работ и услуг"
realization = prihod[prihod["Аналитика (Дебит) 3"] == "Реализация работ и услуг"]
realization_details = (
    realization.groupby(["Аналитика (Кредит) 2", "Месяц"], as_index=False)
    .agg({"Сумма дебита": "sum"})
    .pivot_table(index="Аналитика (Кредит) 2", columns="Месяц", values="Сумма дебита", fill_value=0)
    .reset_index()
)

# Добавляем итоговую сумму по контрагентам
realization_details["Итоговая сумма"] = realization_details.sum(axis=1, numeric_only=True)

# Итоговая строка для "Реализация работ и услуг"
realization_total = pd.DataFrame({
    "Статья доходов": ["Реализация работ и услуг (итого)"],
    "Итоговая сумма": [realization["Сумма дебита"].sum()],
    "November": [realization[realization["Месяц"] == "November"]["Сумма дебита"].sum()]
})

# Объединение данных для "Реализация работ и услуг"
realization_rows = pd.concat([
    realization_total,
    realization_details.rename(columns={"Аналитика (Кредит) 2": "Статья доходов"})
])

classification = pd.concat([
    realization_rows,
    classification[classification["Статья доходов"] != "Реализация работ и услуг"]
], ignore_index=True)

# Сохранение в Excel
output_file = "результат.xlsx"
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    prihod.to_excel(writer, index=False, sheet_name="Приход")
    rashod.to_excel(writer, index=False, sheet_name="Расход")

    workbook = writer.book
    classification.to_excel(writer, index=False, sheet_name="Классификация доходов")
    worksheet = writer.sheets["Классификация доходов"]

    # Форматирование шрифта
    bold_format = workbook.add_format({"bold": True})
    italic_format = workbook.add_format({"italic": True})

    # Применение форматирования
    for row_num, row_data in classification.iterrows():
        if row_data["Статья доходов"].startswith("Реализация работ и услуг"):
            worksheet.write(row_num + 1, 0, row_data["Статья доходов"], bold_format)
        elif pd.notna(row_data["Итоговая сумма"]):  # Если это контрагент внутри "Реализация работ и услуг"
            worksheet.write(row_num + 1, 0, row_data["Статья доходов"], italic_format)

print(f"Файл сохранен как {output_file}")
