import pandas as pd
import json
from datetime import datetime

def parse_excel():
    # Укажите путь к вашему Excel-файлу (в репозитории он должен лежать в корне)
    df = pd.read_excel('Книга1.xlsx', sheet_name='Товары 9170+1955', header=3)  # header=3, т.к. шапка на 4-й строке
    rows = []
    for _, row in df.iterrows():
        # Проверяем, что номер заказа не пустой
        if pd.isna(row.get('Номер заказа')):
            continue
        rows.append([
            str(row.get('Номер заказа', '')),
            int(row.get('N', 0)) if pd.notna(row.get('N')) else 0,
            str(row.get('Номенклатура', '')),
            str(row.get('Проект', 'Без проекта')) if pd.notna(row.get('Проект')) else 'Без проекта',
            str(row.get('Поставщик', 'Не указан')) if pd.notna(row.get('Поставщик')) else 'Не указан',
            str(row.get('Подразделение-получатель', 'Не указано')) if pd.notna(row.get('Подразделение-получатель')) else 'Не указано',
            float(row.get('Сумма', 0)) if pd.notna(row.get('Сумма')) else 0,
            float(row.get('Сумма с НДС', 0)) if pd.notna(row.get('Сумма с НДС')) else 0,
            float(row.get('Сумма без НДС в рублях', 0)) if pd.notna(row.get('Сумма без НДС в рублях')) else 0,
            float(row.get('Кол-во', 0)) if pd.notna(row.get('Кол-во')) else 0,
            row.get('Отменено', 'Нет') == 'Да',
            row.get('Факт на складе', 'Нет') != 'Нет на складе',
            str(row.get('Статья расходов', 'Прочее')) if pd.notna(row.get('Статья расходов')) else 'Прочее',
            row.get('Дата').strftime('%Y-%m-%d') if pd.notna(row.get('Дата')) else datetime.now().strftime('%Y-%m-%d'),
            str(row.get('Валюта', 'руб.')) if pd.notna(row.get('Валюта')) else 'руб.',
            str(row.get('Отменено по причине', '')) if pd.notna(row.get('Отменено по причине')) else ''
        ])
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    print(f'Сохранено {len(rows)} записей в data.json')

if __name__ == '__main__':
    parse_excel()