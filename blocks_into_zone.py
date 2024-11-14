import pandas as pd
import win32com.client


def get_coord_zones(doc, layer="_AS-Зоны"):
    """
    Получает координаты текстовых объектов на указанном слое и возвращает их в виде словаря.
    
    :param doc: Активный документ AutoCAD.
    :param layer: Имя слоя для поиска текстовых объектов (по умолчанию "_AS-Зоны").
    :return: Словарь, где ключ — текст, а значение — кортеж с координатами (X, Y).
    """
    # Удаляем старую выборку, если она существует
    if "TextSelection" in [s.Name for s in doc.SelectionSets]:
        doc.SelectionSets.Item("TextSelection").Delete()

    # Создаем новую выборку
    selection_set = doc.SelectionSets.Add("TextSelection")
    
    # Задаем фильтр для выбора текстовых объектов на указанном слое
    filter_type = [0, 8]  # 0 = тип объекта, 8 = имя слоя
    filter_data = ["TEXT,MULTILINE TEXT", layer]
    selection_set.Select(5, filter_type, filter_data)

    # Словарь для хранения текстов и их координат
    text_coords = {}

    # Перебираем объекты в выборке
    for entity in selection_set:
        # Проверяем тип объекта и получаем текстовое содержимое
        if entity.EntityName == "AcDbText":
            text_content = entity.TextString
            insertion_point = entity.InsertionPoint
            text_coords[text_content] = (insertion_point[0], insertion_point[1])
        # elif entity.EntityName == "AcDbMText":
        #     text_content = entity.Contents
        #     insertion_point = entity.InsertionPoint
        #     text_coords[text_content] = (insertion_point[0], insertion_point[1])

    # Удаляем выборку, чтобы освободить ресурсы
    selection_set.Delete()

    return text_coords

def insert_block_to_zone(doc, data, mapping, zone_coords, block_name="_AS_Base_blck", y_offset=500):
    """
    Расставляет блоки в AutoCAD на основе данных из DataFrame и координат зон.
    
    :param doc: Объект документа AutoCAD.
    :param data: DataFrame с данными из Excel.
    :param mapping: Словарь для сопоставления атрибутов блока с колонками DataFrame.
    :param zone_coords: Словарь с координатами зон, где ключ - имя зоны, а значение - (x, y).
    :param block_name: Название блока для вставки (по умолчанию "_AS_Base_blck").
    :param y_offset: Шаг смещения блока по оси X, по умолчанию 500.
    """
    
    # Проходим по каждой строке в data
    for i, row in data.iterrows():
        # Определяем зону для текущего блока
        zone_name = row.get("Зона")
        
        # Проверяем, есть ли координаты для этой зоны
        if zone_name not in zone_coords:
            print(f"Координаты для зоны '{zone_name}' не найдены. Пропуск.")
            continue
        
        # Получаем координаты зоны
        base_x, base_y = zone_coords[zone_name]
        base_y = base_y + y_offset

        # Смещаем координаты по x для текущего блока
        insertion_point = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8,
            [base_x, base_y, 0]
        )

        zone_coords[zone_name] = base_x, base_y

        # Вставляем блок в указанные координаты
        block_ref = doc.ModelSpace.InsertBlock(insertion_point, block_name, 1, 1, 1, 0)
        
        # Устанавливаем атрибуты блока
        for attrib in block_ref.GetAttributes():
            # Получаем имя атрибута из блока
            attr_tag = attrib.TagString
            
            # Сопоставляем имя атрибута с данными из DataFrame
            if attr_tag in mapping:
                data_field = mapping[attr_tag]
                attrib.TextString = str(row.get(data_field, ""))  # Устанавливаем текст атрибута
        
        print(f"Блок '{block_name}' успешно вставлен в координаты {insertion_point} с атрибутами из строки {i}.")

mapping = {
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


# Чтение данных из Excel-файла
data = pd.read_excel("cad/data.xlsx")

# Подключение к AutoCAD и активному документу
acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

# Название блока
block_name = "_AS_Base_blck"

zone_coords = get_coord_zones(doc)
insert_block_to_zone(doc, data, mapping, zone_coords, block_name, y_offset=500)
