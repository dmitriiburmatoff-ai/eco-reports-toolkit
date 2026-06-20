import copy
from docx import Document

def copy_table_preserving_all(src_table, dst_doc):
    """Копирует таблицу из исходного документа в новый, сохраняя все форматирование."""
    tbl_xml = src_table._element
    new_tbl_xml = copy.deepcopy(tbl_xml)
    dst_doc._element.body.append(new_tbl_xml)
    return dst_doc.tables[-1]

def delete_rows_after(table, keep_rows):
    """Удаляет все строки, начиная с индекса keep_rows (0‑based)."""
    # Удаляем с конца, чтобы не сбивать индексы
    for i in range(len(table.rows) - 1, keep_rows - 1, -1):
        table._element.remove(table.rows[i]._element)
