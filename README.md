

Example use to your code

```
   use Local\Helpers\WordToHtmlHelper;

   $filePath = $_SERVER["DOCUMENT_ROOT"] . "/upload/test.docx";
   $html = WordToHtmlHelper::convertToHtml($filePath);

   if ($html) {
       echo $html;
   } else {
       echo "Ошибка чтения документа";
   }
```


# 📝 WordToHtmlHelper for Bitrix D7

**WordToHtmlHelper** — легковесный PHP-хелпер для Bitrix D7, позволяющий конвертировать документы Word (`.doc` и `.docx`) в HTML без сторонних библиотек, с сохранением базового форматирования.

> 📦 Без `PhpOffice`, без Composer-зависимостей. Поддержка `.docx` реализована через `ZipArchive` + `SimpleXML`. Для `.doc` используется `LibreOffice` в headless-режиме.

---

## ✨ Возможности

- ✅ Конвертация `.docx` в HTML без сторонних библиотек
- ✅ Конвертация `.doc` в HTML через LibreOffice CLI
- ✅ Сохранение базового форматирования: `<p>`, `<strong>`, `<em>`, `<u>`
- ✅ Поддержка структуры Bitrix D7 (`local/php_interface/lib`)
- 🧩 Расширяемая архитектура: можно добавить таблицы, списки, стили и изображения

---


