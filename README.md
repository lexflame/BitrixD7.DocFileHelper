Ecample use to your code


   use Local\Helpers\WordToHtmlHelper;

$filePath = $_SERVER["DOCUMENT_ROOT"] . "/upload/test.docx";
$html = WordToHtmlHelper::convertToHtml($filePath);

if ($html) {
    echo $html;
} else {
    echo "Ошибка чтения документа";
}

