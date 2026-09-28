#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Простой конвертер ODT / DOC → DOCX
Только текст + жирный. Без офисных редакторов.
После успешной конвертации оригинал удаляется.
"""

import os
import sys
import zipfile
import re
import time
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from docx import Document
    from docx.shared import Pt
except ImportError:
    print("Ошибка: нужна библиотека python-docx")
    print("Установите: pip install python-docx odfpy olefile")
    sys.exit(1)


# ==================== ODT ====================

def get_bold_styles(odt_zip):
    bold_styles = set()
    for xml_name in ('styles.xml', 'content.xml'):
        try:
            with odt_zip.open(xml_name) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                for style in root.iter('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style'):
                    name = style.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name')
                    if not name:
                        continue
                    for props in style.iter('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}text-properties'):
                        weight = props.get('{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}font-weight')
                        if weight and str(weight).lower() in ('bold', '700', '800', '900'):
                            bold_styles.add(name)
        except Exception:
            pass
    return bold_styles


def extract_from_odt(odt_path):
    paragraphs = []
    with zipfile.ZipFile(odt_path, 'r') as z:
        bold_styles = get_bold_styles(z)

        with z.open('content.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()

            for p in root.iter('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p'):
                runs = []

                def process_element(elem, current_bold=False):
                    style_name = elem.get('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name')
                    is_bold = current_bold
                    if style_name and style_name in bold_styles:
                        is_bold = True

                    if elem.tag == '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}span':
                        style_name = elem.get('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name')
                        if style_name and style_name in bold_styles:
                            is_bold = True

                    if elem.text:
                        runs.append((elem.text, is_bold))

                    for child in list(elem):
                        process_element(child, is_bold)
                        if child.tail:
                            runs.append((child.tail, current_bold))

                process_element(p)

                if runs:
                    paragraphs.append(runs)
                else:
                    paragraphs.append([('', False)])

    return paragraphs if paragraphs else [[('(пустой документ)', False)]]


# ==================== DOC (максимально безопасное извлечение) ====================

def extract_from_doc(doc_path):
    try:
        import olefile
    except ImportError:
        raise RuntimeError("Нужна библиотека olefile (pip install olefile)")

    if not olefile.isOleFile(doc_path):
        raise RuntimeError("Файл не является OLE-документом (.doc)")

    ole = olefile.OleFileIO(doc_path)
    try:
        if not ole.exists('WordDocument'):
            raise RuntimeError("В файле нет потока WordDocument")

        data = ole.openstream('WordDocument').read()

        # Ограничения
        MAX_SCAN = 6 * 1024 * 1024
        if len(data) > MAX_SCAN:
            data = data[:MAX_SCAN]

        text_parts = []
        i = 0
        data_len = len(data)
        start_time = time.time()
        MAX_TIME = 6.0

        while i < data_len - 1:
            if time.time() - start_time > MAX_TIME:
                break

            try:
                if data[i + 1] == 0 and (32 <= data[i] <= 126 or data[i] >= 0xA0):
                    chars = []
                    j = i
                    while j < data_len - 1:
                        lo = data[j]
                        hi = data[j + 1]
                        if hi == 0 and (32 <= lo <= 126 or lo >= 0xA0):
                            chars.append(chr(lo))
                            j += 2
                        elif hi == 0 and lo in (9, 10, 13):
                            chars.append(chr(lo))
                            j += 2
                        else:
                            break
                    if len(chars) >= 4:
                        text_parts.append(''.join(chars))
                    i = max(j, i + 1)
                else:
                    i += 1
            except IndexError:
                break

        full_text = '\n'.join(text_parts)
        full_text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', full_text)

        paras = []
        for line in full_text.splitlines():
            line = line.strip()
            if line:
                paras.append([(line, False)])

        # Запасной вариант — чистый ASCII
        if not paras:
            try:
                ascii_chunks = re.findall(rb'[\x20-\x7E]{6,}', data)
                for chunk in ascii_chunks[:500]:  # ограничение
                    t = chunk.decode('ascii', errors='ignore').strip()
                    if t and len(t) > 3:
                        paras.append([(t, False)])
            except Exception:
                pass

        if not paras:
            paras = [[('(не удалось извлечь текст из .doc)', False)]]

        return paras

    finally:
        try:
            ole.close()
        except Exception:
            pass


# ==================== Создание DOCX ====================

def create_docx(paragraphs, output_path):
    if not paragraphs:
        paragraphs = [[('(пустой документ)', False)]]

    doc = Document()

    # Безопасно очищаем первый параграф
    if len(doc.paragraphs) > 0:
        try:
            doc.paragraphs[0].clear()
        except Exception:
            pass

    first = True
    for runs in paragraphs:
        if not runs:
            continue

        if first and len(doc.paragraphs) > 0:
            para = doc.paragraphs[0]
            first = False
        else:
            para = doc.add_paragraph()
            first = False

        for item in runs:
            try:
                if isinstance(item, (list, tuple)) and len(item) >= 1:
                    text = str(item[0]) if item[0] is not None else ''
                    is_bold = bool(item[1]) if len(item) > 1 else False
                else:
                    text = str(item)
                    is_bold = False

                if text:
                    run = para.add_run(text)
                    run.bold = is_bold
                    run.font.size = Pt(11)
            except Exception:
                continue

    # Если в итоге документ совсем пустой — добавляем заглушку
    has_text = any(p.text.strip() for p in doc.paragraphs)
    if not has_text:
        if len(doc.paragraphs) > 0:
            doc.paragraphs[0].add_run('(не удалось извлечь текст)')
        else:
            doc.add_paragraph('(не удалось извлечь текст)')

    doc.save(output_path)


# ==================== Основная логика ====================

def convert_file(input_path, output_path):
    ext = Path(input_path).suffix.lower()

    if ext == '.odt':
        paragraphs = extract_from_odt(input_path)
    elif ext == '.doc':
        paragraphs = extract_from_doc(input_path)
    else:
        raise RuntimeError(f"Неподдерживаемый формат: {ext}")

    create_docx(paragraphs, output_path)


def main():
    print("=== Конвертер ODT / DOC → DOCX ===")
    print("Только текст + жирный. Без офисных программ.")
    print("После успешной конвертации оригинал УДАЛЯЕТСЯ.\n")

    if getattr(sys, 'frozen', False):
        root_dir = Path(sys.executable).parent
    else:
        root_dir = Path(__file__).parent.resolve()

    print(f"Рабочая папка: {root_dir}\n")

    files = []
    for pattern in ('*.odt', '*.doc'):
        files.extend(root_dir.rglob(pattern))

    files = sorted(set(files), key=lambda p: str(p).lower())

    if not files:
        print("Файлы .odt и .doc не найдены.")
        input("Нажмите Enter для выхода...")
        return

    print(f"Найдено файлов: {len(files)}\n")

    success = skipped = failed = 0

    for fpath in files:
        try:
            relative = fpath.relative_to(root_dir)
        except ValueError:
            relative = fpath.name

        out_path = fpath.with_suffix('.docx')

        if out_path.exists():
            print(f"[ПРОПУСК] {relative}  →  уже есть .docx")
            skipped += 1
            continue

        print(f"[КОНВЕРТАЦИЯ] {relative} ... ", end='', flush=True)

        try:
            convert_file(str(fpath), str(out_path))

            if out_path.exists() and out_path.stat().st_size > 100:
                try:
                    fpath.unlink()
                    print("OK (оригинал удалён)")
                except Exception as del_err:
                    print(f"OK (не удалось удалить оригинал: {del_err})")
                success += 1
            else:
                print("ОШИБКА (файл не создан или пустой)")
                if out_path.exists():
                    try:
                        out_path.unlink()
                    except Exception:
                        pass
                failed += 1

        except Exception as e:
            print("ОШИБКА")
            print(f"   → {type(e).__name__}: {e}")
            failed += 1
            # Удаляем битый docx если появился
            if out_path.exists():
                try:
                    out_path.unlink()
                except Exception:
                    pass

    print("\n==============================")
    print(f"Успешно:   {success}")
    print(f"Пропущено: {skipped}")
    print(f"Ошибок:    {failed}")
    print("==============================")
    input("\nГотово. Нажмите Enter для выхода...")


if __name__ == '__main__':
    main()
