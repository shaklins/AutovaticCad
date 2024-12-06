import win32com.client
from win32com.client import VARIANT
from pythoncom import VT_R8, VT_ARRAY
import json


def set_leader_ro(doc, data_transform, layer='_AC1-ОБОР-ТЕКСТ'):
    """
    Вставляет выноски с текстом в AutoCAD по данным из data_transform.
    :param doc: Документ AutoCAD (объект COM).
    :param data_transform: Список словарей с координатами и данными для выносок.
    :param layer: Имя слоя, на котором будут размещены выноски.
    """
    for data in data_transform:
        try:
            # Форматируем данные для текста выноски
            clr_data = [int(float(x)) for x in data['1НОМЕР'].split(', ')]
            data_ldr = ', '.join(map(str, clr_data))
            insert_point_x = round(data['coordinates'][0],2)
            insert_point_y = round(data['coordinates'][1],2)

            # Получаем коллекцию слоёв и добавляем слой, если он не существует
            layers = doc.Layers
            if layer not in [l.Name for l in layers]:
                new_layer = layers.Add(layer)
                new_layer.Color = 1  # Пример: цвет слоя (1 - красный)

            # Устанавливаем активный слой
            doc.ActiveLayer = layers.Item(layer)

            # Создаём массив точек для выноски
            points = [
                insert_point_x, insert_point_y, 0,
                insert_point_x + 500, insert_point_y + 300, 0
            ]

            # Создаём объект выноски (Leader)
            leader = doc.ModelSpace.AddMLeader(
                VARIANT(VT_ARRAY | VT_R8, points), 0
            )
            # Добавляем текст к выноске
            leader[0].TextString = data_ldr

        except Exception as e:
            print(f"Ошибка при вставке выноски: {e}")

# Подключение к Автокаду
acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

# Загрузка данных из json
with open('data/data_transform.json', 'r', encoding='utf-8') as f:
    data_transform = json.load(f)

set_leader_ro(doc, data_transform, layer='_AC1-ОБОР-ТЕКСТ')