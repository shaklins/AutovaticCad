import pandas as pd
import win32com.client

# Чтение данных из Excel-файла
data = pd.read_excel("cad/data.xlsx")

# Подключение к AutoCAD и активному документу
acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

# Название блока
block_name = "_AS_Base_blck"

# Начальная точка для вставки первого блока
start_x = 0
start_y = 0
delta_y = -50  # Расстояние между блоками по оси Y

# Привязка атрибутов AutoCAD к столбцам Excel
attribute_mapping = {
    "1НОМЕР": "Номер",
    "2ПОД_НОМЕР": "Под_номер",
    "3НАИМЕНОВАНИЕ": "Наименование",
    "4ПРОИЗВОДИТЕЛЬ": "Производитель",
    "5МОДЕЛЬ": "Модель",
    "6КАБЕЛЬ": "Кабель",
    "7ЗОНА": "Зона",
    "8ЭВ": "ЭВ",
    "9МОЩЬ": "220В",
    "10ЛВС": "ЛВС",
    "11ВЫСОТА": "Высота",
    "12ВЫВОД": "Вывод"
}

# Основной цикл для вставки блоков с атрибутами
for index, row in data.iterrows():
    # Позиция вставки блока
    insertion_point = win32com.client.VARIANT(
        win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8,
        [start_x, start_y + index * delta_y, 0]
    )

    # Вставка блока
    try:
        block_ref = doc.ModelSpace.InsertBlock(insertion_point, block_name, 1, 1, 1, 0)

        # Заполнение атрибутов блока
        for attrib in block_ref.GetAttributes():
            attribute_tag = attrib.TagString
            if attribute_tag in attribute_mapping:
                # Берем значение из соответствующего столбца в data и задаем атрибуту
                attrib.TextString = str(row[attribute_mapping[attribute_tag]])

        print(f"Блок '{block_name}' с атрибутами добавлен в точку ({start_x}, {start_y + index * delta_y}, 0)")

    except Exception as e:
        print(f"Ошибка при вставке блока '{block_name}': {e}")

print("Завершено!")
