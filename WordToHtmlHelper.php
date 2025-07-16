<?php
namespace Local\Utils;

use Bitrix\Main\Data\Cache;
use Bitrix\Main\Data\TaggedCache;
use DOMDocument;
use SimpleXMLElement;
use ZipArchive;

class WordToHtmlHelper
{
    // Время жизни кэша (в секундах)
    protected static int $cacheTtlSeconds = 86400; // 24 часа

    /**
     * Основная функция: конвертирует DOCX или DOC в HTML,
     * сохраняет результат в кэш, привязывает к инфоблоку через TaggedCache.
     *
     * @param string $filePath Полный путь к файлу .doc или .docx
     * @param int $iblockId ID инфоблока для привязки кэша
     * @return string|null HTML-представление файла или null при ошибке
     */
    public static function convertToHtmlWithIblockTag(string $filePath, int $iblockId): ?string
    {
        if (!is_file($filePath)) {
            return null;
        }

        $cache = Cache::createInstance();
        $cacheId = 'docx_html_' . md5($filePath . filemtime($filePath));
        $cacheDir = '/local/word2html/';
        $taggedCache = new TaggedCache();

        if ($cache->initCache(self::$cacheTtlSeconds, $cacheId, $cacheDir)) {
            $vars = $cache->getVars();
            return $vars['HTML'] ?? null;
        } elseif ($cache->startDataCache()) {
            $taggedCache->startTagCache($cacheDir);
            $taggedCache->registerTag("iblock_id_{$iblockId}");

            $ext = strtolower(pathinfo($filePath, PATHINFO_EXTENSION));
            $html = null;

            if ($ext === 'docx') {
                $html = self::convertDocxToHtml($filePath);
            } elseif ($ext === 'doc') {
                $html = self::convertDocToHtmlViaLibreOffice($filePath);
            }

            if ($html) {
                $taggedCache->endTagCache();
                $cache->endDataCache(['HTML' => $html]);
                return $html;
            } else {
                $taggedCache->abortTagCache();
                $cache->abortDataCache();
            }
        }

        return null;
    }

    /**
     * Преобразует .doc файл в HTML через LibreOffice
     * Требует установленного libreoffice с CLI-интерфейсом
     *
     * @param string $filePath Путь к .doc файлу
     * @return string|null HTML содержимое или null
     */
    protected static function convertDocToHtmlViaLibreOffice(string $filePath): ?string
    {
        $tempDir = sys_get_temp_dir() . '/' . uniqid('doc_', true);
        mkdir($tempDir, 0777, true);

        $cmd = sprintf(
            'libreoffice --headless --convert-to html --outdir %s %s 2>&1',
            escapeshellarg($tempDir),
            escapeshellarg($filePath)
        );
        exec($cmd);

        $htmlFile = $tempDir . '/' . pathinfo($filePath, PATHINFO_FILENAME) . '.html';
        if (is_file($htmlFile)) {
            $html = file_get_contents($htmlFile);
            array_map('unlink', glob("$tempDir/*"));
            rmdir($tempDir);
            return $html;
        }
        return null;
    }

    /**
     * Извлекает текст из .docx файла, разбирая его XML содержимое
     *
     * @param string $filePath Путь к .docx файлу
     * @return string|null HTML или null при ошибке
     */
    protected static function convertDocxToHtml(string $filePath): ?string
    {
        $zip = new ZipArchive();
        if ($zip->open($filePath) !== true) return null;

        $body = new SimpleXMLElement($zip->getFromName("word/document.xml"));
        $rels = new SimpleXMLElement($zip->getFromName("word/_rels/document.xml.rels"));
        $mediaFiles = [];

        // Извлечение встроенных изображений
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $stat = $zip->statIndex($i);
            if (str_starts_with($stat['name'], 'word/media/')) {
                $mediaFiles[$stat['name']] = base64_encode($zip->getFromIndex($i));
            }
        }
        $zip->close();

        return self::parseBody($body->body, $rels, $mediaFiles);
    }

    /**
     * Разбор XML тела документа и конвертация в HTML
     *
     * @param SimpleXMLElement $body Основной XML-узел
     * @param SimpleXMLElement $rels Отношения (включая изображения)
     * @param array $mediaFiles Ассоциативный массив base64-картинок
     * @return string HTML
     */
    protected static function parseBody($body, $rels, $mediaFiles): string
    {
        $html = '';
        $listBuffer = [];
        $listType = 'ul';

        foreach ($body->children() as $node) {
            $name = $node->getName();

            // Обработка списков
            if ($name === 'p' && isset($node->pPr->numPr)) {
                $abstractId = (int) ($node->pPr->numPr->numId['val'] ?? 1);
                $currentListType = ($abstractId % 2 === 0) ? 'ol' : 'ul';

                if ($listType !== $currentListType && !empty($listBuffer)) {
                    $html .= "<$listType>" . implode('', $listBuffer) . "</$listType>";
                    $listBuffer = [];
                }

                $listType = $currentListType;
                $listBuffer[] = self::parseParagraph($node);
                continue;
            }

            // Закрыть текущий список
            if (!empty($listBuffer)) {
                $html .= "<$listType>" . implode('', $listBuffer) . "</$listType>";
                $listBuffer = [];
            }

            switch ($name) {
                case 'p':
                    $html .= self::parseParagraph($node);
                    break;
                case 'tbl':
                    $html .= self::parseTable($node);
                    break;
                case 'sdt':
                    $html .= self::parseStructuredDocumentTag($node, $rels, $mediaFiles);
                    break;
            }
        }

        if (!empty($listBuffer)) {
            $html .= "<$listType>" . implode('', $listBuffer) . "</$listType>";
        }

        return $html;
    }

    /**
     * Конвертация параграфа в HTML, с поддержкой выравнивания и отступов
     * @param SimpleXMLElement $node
     * @return string HTML строки
     */
    protected static function parseParagraph($node): string
    {
        $style = '';

        if (isset($node->pPr->jc)) {
            $align = (string)$node->pPr->jc['val'];
            $style .= "text-align:$align;";
        }
        if (isset($node->pPr->ind)) {
            $left = (int)$node->pPr->ind['left'];
            $style .= "margin-left:" . round($left / 20) . "px;";
        }

        $text = '';
        foreach ($node->r as $run) {
            foreach ($run->t as $t) {
                $text .= htmlspecialchars((string)$t);
            }
        }

        return '<li><p style="' . $style . '">' . $text . '</p></li>';
    }

    /**
     * Конвертация таблицы Word в HTML <table>
     * @param SimpleXMLElement $node
     * @return string HTML таблица
     */
    protected static function parseTable($node): string
    {
        $html = '<table border="1">';
        foreach ($node->tr as $tr) {
            $html .= '<tr>';
            foreach ($tr->tc as $tc) {
                $html .= '<td>';
                foreach ($tc->p as $p) {
                    $html .= self::parseParagraph($p);
                }
                $html .= '</td>';
            }
            $html .= '</tr>';
        }
        return $html . '</table>';
    }

    /**
     * Обработка встроенных изображений через <img src="data:">
     * @param SimpleXMLElement $node
     * @param SimpleXMLElement $rels
     * @param array $mediaFiles
     * @return string
     */
    protected static function parseStructuredDocumentTag($node, $rels, $mediaFiles): string
    {
        foreach ($node->xpath('.//a:blip') as $blip) {
            $rId = (string)$blip['r:embed'];
            foreach ($rels->Relationship as $rel) {
                if ((string)$rel['Id'] === $rId) {
                    $target = 'word/' . (string)$rel['Target'];
                    if (isset($mediaFiles[$target])) {
                        $base64 = $mediaFiles[$target];
                        return '<img src="data:image/png;base64,' . $base64 . '" />';
                    }
                }
            }
        }
        return '';
    }
}
