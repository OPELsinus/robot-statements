import pandas as pd

from config import save_xlsx_path

df = pd.read_excel(f'{save_xlsx_path}\\Шаблоны для робота (не удалять)\\Копия Сверка ОБРАЗЕЦ.xlsx', sheet_name=2)

df.columns = ['', 'Статья в ДДС', 'Код', 'Название статьи'] + [''] * (len(df.columns) - 4)

df = df.dropna(subset=['Статья в ДДС'])

df1 = pd.read_excel('koks.xlsx')

print(df1['Код БДДС'])

# print(df[df['Код'] == df1['Код БДДС']]['Статья в ДДС'])
df1 = df1.merge(df, left_on='Код БДДС', right_on='Код', how='left')
print(df1)
df1.to_excel('pokus.xlsx')
