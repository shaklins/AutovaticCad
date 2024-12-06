import win32com.client
import json
import sys


def get_data_blocks(doc, block_name="_AS_Base_blck"):
    """
    Извлекает данные из блоков с указанным именем в AutoCAD с использованием SelectionSets.
    :param doc: Объект документа AutoCAD.
    :param block_name: Название блока для извлечения данных (по умолчанию "_AS_Base_blck").
    :return: Список словарей с данными блоков.
    """
    data_blocks = []

    # Удаляем старый SelectionSet с таким же именем, если он существует
    sel_set_name = "GetBlocksSelectionSet"
    for sel in doc.SelectionSets:
        if sel.Name == sel_set_name:
            sel.Delete()

    # Создаём новый SelectionSet
    sel_set = doc.SelectionSets.Add(sel_set_name)

    # Устанавливаем фильтр для выбора только блоков с именем block_name
    filter_type = [0, 2]  # 0 - тип объекта, 2 - имя блока
    filter_data = ["INSERT", block_name]
    sel_set.Select(5, [0, 2], ["INSERT", block_name])
    total_blcks = len(sel_set)
    # Обрабатываем каждый выбранный блок
    for idx, entity in enumerate(sel_set):
        block_data = {}
        try:
            # Извлекаем координаты блока
            block_data["x"] = entity.InsertionPoint[0]
            block_data["y"] = entity.InsertionPoint[1]
            block_data["id_block"] = entity.ObjectID

            # Извлекаем атрибуты блока
            for attrib in entity.GetAttributes():
                attr_tag = attrib.TagString
                block_data[attr_tag] = attrib.TextString

            # Добавляем данные блока в список
            data_blocks.append(block_data)
        except:
            continue

        sys.stdout.write(f'\rProcessing: {idx + 1}/{total_blcks}')
        sys.stdout.flush()

    # Удаляем SelectionSet, чтобы не засорять память
    sel_set.Delete()
    print('\rProcces complete!')
    return data_blocks

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

# Название блока
block_name = "_AS_Base_blck"

data_block = get_data_blocks(doc,block_name)

with open('data/data_acad.json','w',encoding="utf-8") as f:
    json.dump(data_block, f, ensure_ascii=False)