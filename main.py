import pandas as pd

# Чтение Excel файла
input_file = 'input.xlsx'  # Имя входного файла
output_file = 'output.xlsx'  # Имя выходного файла

# Чтение листа 0
df = pd.read_excel(input_file, sheet_name=0, header=[0, 1])

# Преобразование многоуровневых заголовков в одномерные
df.columns = df.columns.map(lambda x: ' '.join([str(i) for i in x if i]).strip())

# Разделение объединенных строк на отдельные столбцы
df['Счет1'] = df['Дебет Счет']
df['Сумма1'] = df['Дебет Сумма']
df['Счет2'] = df['Кредит Счет']
df['Сумма2'] = df['Кредит Сумма']

# Конвертация сумм в числовой формат
df['Сумма Дебет'] = pd.to_numeric(df['Сумма Дебет'], errors='coerce')
df['Сумма Кредит'] = pd.to_numeric(df['Сумма Кредит'], errors='coerce')

# Фильтрация приходов
prihod_df = df[(df['Счет Дебет'].str.startswith('10')) & (~df['Счет Кредит'].str.startswith('10'))]

# Разделение столбцов Документ, Аналитика Дт и Аналитика Кт
prihod_df[['Документ 1.1', 'Документ 1.2', 'Документ 1.3']] = prihod_df['Документ'].str.split('\n', expand=True)
prihod_df[['Аналитика Дт 2.1', 'Аналитика Дт 2.2', 'Аналитика Дт 2.3']] = prihod_df['Аналитика Дт'].str.split('\n',
                                                                                                              expand=True)
prihod_df[['Аналитика Кт 3.1', 'Аналитика Кт 3.2', 'Аналитика Кт 3.3']] = prihod_df['Аналитика Кт'].str.split('\n',
                                                                                                              expand=True)

# Оставляем необходимые столбцы
prihod_df = prihod_df[['Период', 'Документ 1.1', 'Документ 1.2', 'Документ 1.3', 'Аналитика Дт 2.1', 'Аналитика Дт 2.2',
                       'Аналитика Дт 2.3', 'Аналитика Кт 3.1', 'Аналитика Кт 3.2', 'Аналитика Кт 3.3',
                       'Счет Дебет', 'Сумма Дебет', 'Счет Кредит', 'Сумма Кредит', 'Общий оборот', 'Текущее сальдо']]

# Получаем уникальные значения из Аналитика Кт 3.3 и суммируем по ним
klassifikaciya_df = prihod_df.groupby('Аналитика Кт 3.3')['Сумма Дебет'].sum().reset_index()

# Фильтруем данные для "Реализация работ и услуг"
realizaciya_df = prihod_df[prihod_df['Аналитика Кт 3.3'] == 'Реализация работ и услуг']
kontragenty_df = realizaciya_df.groupby('Аналитика Кт 3.2')['Сумма Дебет'].sum().reset_index()

# Запись на второй лист
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    prihod_df.to_excel(writer, sheet_name='Приходы', index=False)
    klassifikaciya_df.to_excel(writer, sheet_name='Классификация доходов', index=False)
    kontragenty_df.to_excel(writer, sheet_name='Реализация услуг', index=False)
