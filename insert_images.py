import os
import glob
import sys
import win32com.client as win32

# === НАСТРОЙКИ ===
IMG_FOLDER = r"F:\\"
IMAGE_EXTENSIONS = ["*.png", "*.jpg", "*.jpeg"]
# =================

# Собираем изображения
image_files = []
for ext in IMAGE_EXTENSIONS:
    image_files.extend(glob.glob(os.path.join(IMG_FOLDER, ext)))
image_files = sorted(image_files)

if not image_files:
    print(f"❌ Нет изображений в папке: {IMG_FOLDER}")
    input("Нажмите Enter для выхода...")
    sys.exit(1)

try:
    word = win32.gencache.EnsureDispatch('Word.Application')
except:
    print("❌ Не удалось запустить Word!")
    input("Нажмите Enter для выхода...")
    sys.exit(1)

try:
    doc = word.ActiveDocument
    print(f"📄 Документ: {doc.Name}")
except:
    print("❌ Нет активного документа!")
    input("Нажмите Enter для выхода...")
    sys.exit(1)

# Получаем константы Word
wdCollapseEnd = 0
wdAlignParagraphCenter = 1
wdFieldSequence = 54

print(f"Найдено {len(image_files)} изображений для обработки\n")

# ВСТАВКА ИЗОБРАЖЕНИЙ С АВТОНУМЕРАЦИЕЙ
for i, img_path in enumerate(image_files, start=1):
    label = f"#{i:04d}"
    print(f"🔍 Ищем метку: {label}")
    
    # Ищем ВСЕ вхождения этой метки в документе
    find = doc.Content.Find
    find.Text = label
    find.Forward = True
    find.Wrap = 0  # wdFindStop = 0
    
    found_count = 0
    
    while find.Execute():
        found_count += 1
        print(f"  Найдено вхождение #{found_count} для метки {label}")
        
        # --- Сохраняем позицию метки ---
        label_range = find.Parent.Duplicate
        pos = label_range.Start

        # --- Удаляем метку ---
        label_range.Delete()

        # --- Извлекаем текст из имени файла ---
        filename = os.path.basename(img_path)
        name_without_ext = os.path.splitext(filename)[0]
        caption_text = name_without_ext.split("_", 1)[1] if "_" in name_without_ext else name_without_ext

        # --- Создаём диапазон на месте удалённой метки ---
        rng = doc.Range(pos, pos)
        
        # 1. Вставляем пустую строку перед картинкой
        rng.InsertBefore("\r")
        rng.Collapse(wdCollapseEnd)
        
        # 2. Сохраняем позицию для вставки картинки
        img_pos = rng.Start
        
        # 3. Вставляем картинку
        try:
            img_range = doc.Range(img_pos, img_pos)
            img = img_range.InlineShapes.AddPicture(
                FileName=img_path,
                LinkToFile=False,
                SaveWithDocument=True
            )
            print(f"    ✅ Картинка вставлена: {filename}")
        except Exception as e:
            print(f"    ❌ Ошибка вставки картинки: {e}")
            continue
        
        # 4. Центрируем картинку
        img.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        img.Range.ParagraphFormat.FirstLineIndent = 0
        
        # 5. Определяем позицию для подписи
        after_img_pos = img.Range.End
        
        # 6. ВСТАВЛЯЕМ АВТОНУМЕРОВАННУЮ ПОДПИСЬ
        caption_range = doc.Range(after_img_pos, after_img_pos)
        caption_range.Collapse(wdCollapseEnd)
        
        # Вставляем абзац после картинки
        caption_range.InsertParagraphAfter()
        caption_range.Collapse(wdCollapseEnd)
        
        # Теперь в этом диапазоне создаем поле автонумерации
        # Сначала вставляем текст "Рисунок "
        caption_range.Text = "Рисунок "
        
        # Затем добавляем поле автонумерации
        field_range = doc.Range(caption_range.End, caption_range.End)
        field = doc.Fields.Add(
            Range=field_range,
            Type=wdFieldSequence,
            Text=r'"Рисунок" \* ARABIC',
            PreserveFormatting=False
        )
        
        # Добавляем разделитель и текст
        after_field_range = doc.Range(field.Result.End, field.Result.End)
        after_field_range.Text = f" — {caption_text}"
        
        # Объединяем весь диапазон подписи для форматирования
        full_caption_range = doc.Range(caption_range.Start, after_field_range.End)
        
        # 7. Форматируем подпись
        full_caption_range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        full_caption_range.ParagraphFormat.FirstLineIndent = 0
        full_caption_range.ParagraphFormat.LeftIndent = 0
        full_caption_range.ParagraphFormat.LineSpacingRule = 0  # 0 = wdLineSpaceSingle
        
        # Применяем стиль
        try:
            full_caption_range.Style = "Заголовок рисунка"
        except:
            try:
                full_caption_range.Style = "Рисунок"
            except:
                try:
                    full_caption_range.Style = "Основной"
                except:
                    try:
                        full_caption_range.Style = "Normal"
                    except:
                        pass
        
        print(f"    ✅ Добавлена автонумерованная подпись: Рисунок {i} — {caption_text}")
        
        # Обновляем поле, чтобы показался номер
        field.Update()
    
    if found_count == 0:
        print(f"⚠️  Метка {label} не найдена.")
    else:
        print(f"✓ Обработано {found_count} вхождений для {label}\n")

# Обновляем все поля в документе (чтобы номера отобразились)
doc.Fields.Update()

print("\n" + "="*50)
print("✨ ГОТОВО!")
print("="*50)
print(f"Обработано {len(image_files)} изображений.")
print("Теперь вы можете вставлять ссылки через:")
print("Ссылки → Перекрёстная ссылка → Рисунок")
print("\nЕсли нумерация не отображается, нажмите Ctrl+A затем F9")
print("="*50)
input("\nНажмите Enter для завершения...")