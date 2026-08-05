import pandas as pd
import json
from datetime import datetime

def parse_excel():
    # Читаем файл, пропуская первые 3 строки (заголовок, пустая, курсы)
    df = pd.read_excel('Книга1.xlsx', sheet_name='Товары 9170+1955', header=3)
    
    # Удаляем строки, где все значения NaN (пустые)
    df = df.dropna(how='all')
    
    rows = []
    for _, row in df.iterrows():
        # Проверяем, что есть номер заказа (чтобы не брать пустые строки)
        if pd.isna(row.get('Номер заказа')):
            continue
        
        # Преобразуем дату в строку, если она есть
        date_val = row.get('Дата')
        if pd.isna(date_val):
            date_str = datetime.now().strftime('%Y-%m-%d')
        else:
            date_str = date_val.strftime('%Y-%m-%d')
        
        # Обрабатываем числовые поля
        def safe_float(val):
            try:
                return float(val)
            except:
                return 0.0
        
        rows.append([
            str(row.get('Номер заказа', '')),
            int(row.get('N', 0)) if pd.notna(row.get('N')) else 0,
            str(row.get('Номенклатура', '')),
            str(row.get('Проект', 'Без проекта')) if pd.notna(row.get('Проект')) else 'Без проекта',
            str(row.get('Поставщик', 'Не указан')) if pd.notna(row.get('Поставщик')) else 'Не указан',
            str(row.get('Подразделение-получатель', 'Не указано')) if pd.notna(row.get('Подразделение-получатель')) else 'Не указано',
            safe_float(row.get('Сумма', 0)),
            safe_float(row.get('Сумма с НДС', 0)),
            safe_float(row.get('Сумма без НДС в рублях', 0)),
            safe_float(row.get('Кол-во', 0)),
            row.get('Отменено', 'Нет') == 'Да',
            row.get('Факт на складе', 'Нет') != 'Нет на складе',
            str(row.get('Статья расходов', 'Прочее')) if pd.notna(row.get('Статья расходов')) else 'Прочее',
            date_str,
            str(row.get('Валюта', 'руб.')) if pd.notna(row.get('Валюта')) else 'руб.',
            str(row.get('Отменено по причине', '')) if pd.notna(row.get('Отменено по причине')) else ''
        ])
    
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    
    print(f'✅ Сохранено {len(rows)} записей в data.json')

if __name__ == '__main__':
    parse_excel()
