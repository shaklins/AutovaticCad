import json
import time

import win32com.client
import sys

def transform_data_coord(data_blocks, threshold=50):
    """
    Преобразует список данных блоков, объединяя записи с близкими координатами.
    :param data_blocks: Список словарей с данными блоков (из функции get_data_blocks).
    :param threshold: Максимальное расстояние между координатами для объединения (по умолчанию 50).
    :return: Список словарей, где ключами являются координаты (x, y),
             а значениями - объединённые данные из исходного списка.
    """
    result = {}
    len_data = len(data_blocks)
    for idx, block in enumerate(data_blocks):
        x, y = block["x"], block["y"]

        # Найти существующий ключ с близкими координатами
        found_key = None
        for coord_key in result:
            if abs(coord_key[0] - x) < threshold and abs(coord_key[1] - y) < threshold:
                found_key = coord_key
                break

        # Если найден близкий ключ, объединить данные
        if found_key:
            for key, value in block.items():
                if key not in ["x", "y", "id_block"]:
                    if key in result[found_key]:
                        result[found_key][key] += f", {value}"
                    else:
                        result[found_key][key] = value
        else:
            # Если ключа нет, добавить новую запись
            coord_key = (x, y)
            result[coord_key] = {k: v for k, v in block.items() if k not in ["x", "y", "id_block"]}

        sys.stdout.write(f'\rProcessing: {idx + 1}/{len_data}')
        sys.stdout.flush()
        time.sleep(0.05)
    print('\rProcces complete!')

    # Преобразовать результат в список
    return [{"coordinates": coord, **data} for coord, data in result.items()]

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

with open('data/data_acad.json','r',encoding="utf-8") as f:
    data_block = json.load(f)

transform_data = transform_data_coord(data_block)

with open('data/data_transform.json','w',encoding="utf-8") as f:
    json.dump(transform_data, f, ensure_ascii=False)