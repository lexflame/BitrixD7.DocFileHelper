Ecample use to your code

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
🧩 Поддержка форматирования
Для .docx поддерживается:

Параграфы (<p>)
* Жирный, курсив, подчеркивание (<span style="...">)
* Примитивная работа с текстом


📎 Зависимости
PHP: стандартные модули ZipArchive и SimpleXML.
LibreOffice CLI: только для .doc.
