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

# Добавление колонок "Год" и "Месяц"
prihod["Год"] = pd.to_datetime(prihod["Дата"], dayfirst=True, errors="coerce").dt.year
prihod["Месяц"] = pd.to_datetime(prihod["Дата"], dayfirst=True, errors="coerce").dt.month_name()

rashod["Год"] = pd.to_datetime(rashod["Дата"], dayfirst=True, errors="coerce").dt.year
rashod["Месяц"] = pd.to_datetime(rashod["Дата"], dayfirst=True, errors="coerce").dt.month_name()

# Расположение колонок для "Приход"
prihod = prihod[
    ["Год", "Месяц", "Дата"] +
    [col for col in prihod_columns if "Аналитика" in col and "Дебит" not in col and "Кредит" not in col] +
    [col for col in prihod_columns if "Дебит" in col] +
    [col for col in prihod_columns if "Кредит" in col] +
    ["Счет дебита", "Сумма дебита", "Счет кредита"]
]

# Расположение колонок для "Расход"
rashod = rashod[
    ["Год", "Месяц", "Дата"] +
    [col for col in rashod_columns if "Аналитика" in col and "Дебит" not in col and "Кредит" not in col] +
    [col for col in rashod_columns if "Дебит" in col] +
    [col for col in rashod_columns if "Кредит" in col] +
    ["Счет дебита", "Счет кредита", "Сумма кредита"]
]

# Создание классификации доходов
classification = (
    prihod.groupby("Аналитика (Дебит) 3", as_index=False)
    .agg({"Сумма дебита": "sum"})
    .rename(columns={"Аналитика (Дебит) 3": "Статья доходов", "Сумма дебита": "Итого доходов"})
)

# Добавление месячных колонок для доходов
month_columns = prihod.pivot_table(
    index="Аналитика (Дебит) 3",
    columns="Месяц",
    values="Сумма дебита",
    aggfunc="sum",
    fill_value=0
).reset_index()

classification = classification.merge(month_columns, how="left", left_on="Статья доходов", right_on="Аналитика (Дебит) 3").drop(columns=["Аналитика (Дебит) 3"])

# Сортировка доходов от большего к меньшему
classification = classification.sort_values(by="Итого доходов", ascending=False)

# Упорядочивание месяцев
month_order = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]
classification_columns = ["Статья доходов", "Итого доходов"] + [month for month in month_order if month in classification.columns]
classification = classification[classification_columns]

# Создание классификации расходов
classification_rashod = (
    rashod.groupby("Аналитика (Кредит) 3", as_index=False)
    .agg({"Сумма кредита": "sum"})
    .rename(columns={"Аналитика (Кредит) 3": "Статья расходов", "Сумма кредита": "Итого расходов"})
)

# Добавление месячных колонок для расходов
rashod_month_columns = rashod.pivot_table(
    index="Аналитика (Кредит) 3",
    columns="Месяц",
    values="Сумма кредита",
    aggfunc="sum",
    fill_value=0
).reset_index()

classification_rashod = classification_rashod.merge(
    rashod_month_columns, how="left", left_on="Статья расходов", right_on="Аналитика (Кредит) 3"
).drop(columns=["Аналитика (Кредит) 3"])

# Сортировка расходов от большего к меньшему
classification_rashod = classification_rashod.sort_values(by="Итого расходов", ascending=False)

# Упорядочивание месяцев для расходов
classification_rashod_columns = ["Статья расходов", "Итого расходов"] + [
    month for month in month_order if month in classification_rashod.columns
]
classification_rashod = classification_rashod[classification_rashod_columns]

# Проверки по доходам
check_total_prihod = prihod["Сумма дебита"].sum()
check_total_classification = classification["Итого доходов"].sum()
difference_total = check_total_prihod - check_total_classification
check_result = "True" if abs(difference_total) < 1e-5 else "False"

# Проверки по расходам
check_total_rashod = rashod["Сумма кредита"].sum()
check_total_classification_rashod = classification_rashod["Итого расходов"].sum()
difference_rashod_total = check_total_rashod - check_total_classification_rashod
check_rashod_result = "True" if abs(difference_rashod_total) < 1e-5 else "False"

# Сохранение в Excel с проверками, форматированием чисел и названиями месяцев
output_file = "результат.xlsx"
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    prihod.to_excel(writer, index=False, sheet_name="Приход")
    rashod.to_excel(writer, index=False, sheet_name="Расход")
    classification.to_excel(writer, index=False, sheet_name="Классификация доходов", startrow=1)
    classification_rashod.to_excel(writer, index=False, sheet_name="Классификация расходов", startrow=1)

    workbook = writer.book
    number_format = workbook.add_format({"num_format": "#,##0", "align": "right"})
    merge_format = workbook.add_format({"align": "center", "bold": True, "border": 1})

    # Добавление проверок по месяцам для "Классификация доходов"
    worksheet = writer.sheets["Классификация доходов"]
    worksheet.merge_range(0, 2, 0, len(classification_columns) - 1, "2024", merge_format)

    # Проверки по горизонтали
    for i, month in enumerate(month_order):
        worksheet.write(len(classification) + 3, i + 2, f"Check ({month})")


#ЗДЕСЬ НОВЫЙ КОД







print(f"Файл сохранен как {output_file}")
