import pandas as pd
from openpyxl import load_workbook


def update_ids(file_paths):
    last_id = None
    
    for i, file_path in enumerate(file_paths):
        # Загрузка файла
        df = pd.read_excel(file_path)
        
        # Определение индекса столбца с ID (3 для первого файла, 4 для остальных)
        id_column_index = 3 if i == 0 else 4
        
        if i == 0:
            # Для первого файла берем значение ID из второй строки
            last_id = df.iloc[1, id_column_index]
        
        # Если last_id все еще None (например, если первый файл пустой), начинаем с 1
        if last_id is None:
            last_id = 0
        
        # Создание новых ID
        new_ids = range(last_id + 1, last_id + len(df) + 1)
        
        # Замена ID в DataFrame
        df.iloc[:, id_column_index] = new_ids
        
        # Обновление last_id для следующего файла
        last_id = last_id + len(df)
        
        # Сохранение изменений в файл
        book = load_workbook(file_path)
        sheet_name = book.sheetnames[0]
        
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            # Копирование существующих листов, кроме того, который мы обновляем
            for sheet in book.sheetnames:
                if sheet != sheet_name:
                    book[sheet].copy(writer.book, sheet)
            
            # Запись обновленных данных
            df.to_excel(writer, index=False, sheet_name=sheet_name)

# Пути к файлам
file_paths = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx']

# Вызов функции
update_ids(file_paths)

print("ID обновлены во всех файлах.")
