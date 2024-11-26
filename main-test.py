# Добавление итогов и проверок для "Классификация доходов"
classification["Итого доходов"] = classification.iloc[:, 2:].sum(axis=1)  # Суммирование всех месячных значений

# Итоговая строка для "Итого доходов" (сумма всех ячеек над ней)
total_income = classification["Итого доходов"].sum()

# Суммы по месяцам из страницы "Приход"
monthly_totals_income = (
    prihod.groupby("Месяц")["Сумма дебита"].sum()
).reindex(month_order, fill_value=0)

# Добавление итоговой строки
classification.loc[len(classification)] = ["Итого доходов", total_income] + monthly_totals_income.tolist()

# Суммы из страницы "Приход" (для проверки)
total_prihod = prihod["Сумма дебита"].sum()
classification.loc[len(classification)] = ["Итого приход", total_prihod] + monthly_totals_income.tolist()

# Разница между "Итого доходов" и "Итого приход"
difference_income = total_income - total_prihod
classification.loc[len(classification)] = ["Разница", difference_income] + [""] * (len(month_order))

# Проверка: совпадает ли сумма доходов с итогами "Приход"
check_result_income = "TRUE" if abs(difference_income) < 1e-5 else "FALSE"
classification.loc[len(classification)] = ["Check!", check_result_income] + [""] * (len(month_order))

# Аналогичная логика для "Классификация расходов"
classification_rashod["Итого расходов"] = classification_rashod.iloc[:, 2:].sum(axis=1)

# Итоговая строка для "Итого расходов" (сумма всех ячеек над ней)
total_expenses = classification_rashod["Итого расходов"].sum()

# Суммы по месяцам из страницы "Расход"
monthly_totals_expenses = (
    rashod.groupby("Месяц")["Сумма кредита"].sum()
).reindex(month_order, fill_value=0)

# Добавление итоговой строки
classification_rashod.loc[len(classification_rashod)] = ["Итого расходов", total_expenses] + monthly_totals_expenses.tolist()

# Суммы из страницы "Расход" (для проверки)
total_rashod = rashod["Сумма кредита"].sum()
classification_rashod.loc[len(classification_rashod)] = ["Итого расход", total_rashod] + monthly_totals_expenses.tolist()

# Разница между "Итого расходов" и "Итого расход"
difference_expenses = total_expenses - total_rashod
classification_rashod.loc[len(classification_rashod)] = ["Разница", difference_expenses] + [""] * (len(month_order))

# Проверка: совпадает ли сумма расходов с итогами "Расход"
check_result_expenses = "TRUE" if abs(difference_expenses) < 1e-5 else "FALSE"
classification_rashod.loc[len(classification_rashod)] = ["Check!", check_result_expenses] + [""] * (len(month_order))

# Сохранение итогового файла
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
    worksheet_income = writer.sheets["Классификация доходов"]
    worksheet_income.merge_range(0, 2, 0, len(classification_columns) - 1, "2024", merge_format)

    # Добавление проверок по месяцам для "Классификация расходов"
    worksheet_expenses = writer.sheets["Классификация расходов"]
    worksheet_expenses.merge_range(0, 2, 0, len(classification_rashod_columns) - 1, "2024", merge_format)

print(f"Файл сохранен как {output_file}")
