<?php
namespace Local\Utils;

use Bitrix\Main\Data\Cache;
use Bitrix\Main\Data\TaggedCache;
use DOMDocument;
use SimpleXMLElement;
use ZipArchive;

class WordToHtmlHelper
{
    protected static int $cacheTtlSeconds = 86400; // 24 часа

    /**
     * Конвертация DOCX или DOC в HTML с кэшированием и тегами инфоблока
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

    protected static function convertDocxToHtml(string $filePath): ?string
    {
        $zip = new ZipArchive();
        if ($zip->open($filePath) !== true) return null;

        $body = new SimpleXMLElement($zip->getFromName("word/document.xml"));
        $rels = new SimpleXMLElement($zip->getFromName("word/_rels/document.xml.rels"));
        $mediaFiles = [];

        for ($i = 0; $i < $zip->numFiles; $i++) {
            $stat = $zip->statIndex($i);
            if (str_starts_with($stat['name'], 'word/media/')) {
                $mediaFiles[$stat['name']] = base64_encode($zip->getFromIndex($i));
            }
        }
        $zip->close();

        return self::parseBody($body->body, $rels, $mediaFiles);
    }

    protected static function parseBody($body, $rels, $mediaFiles): string
    {
        $html = '';
        $listBuffer = [];
        $listType = 'ul';

        foreach ($body->children() as $node) {
            $name = $node->getName();

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
