<?php

namespace Local\Helpers;

class WordToHtmlHelper
{
    public static function convertToHtml(string $filePath): ?string
    {
        if (!is_file($filePath)) {
            return null;
        }

        $ext = strtolower(pathinfo($filePath, PATHINFO_EXTENSION));

        if ($ext === 'docx') {
            return self::convertDocxToHtml($filePath);
        } elseif ($ext === 'doc') {
            return self::convertDocToHtmlViaLibreOffice($filePath);
        }

        return null;
    }

    protected static function convertDocxToHtml(string $filePath): ?string
    {
        $zip = new \ZipArchive();
        if ($zip->open($filePath) !== true) {
            return null;
        }

        $documentXml = $zip->getFromName('word/document.xml');
        $relsXml = $zip->getFromName('word/_rels/document.xml.rels');
        $mediaFiles = [];

        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = $zip->getNameIndex($i);
            if (strpos($name, 'word/media/') === 0) {
                $mediaFiles[$name] = base64_encode($zip->getFromName($name));
            }
        }

        $zip->close();

        if (!$documentXml) {
            return null;
        }

        $rels = [];
        if ($relsXml) {
            $relsXml = simplexml_load_string($relsXml);
            foreach ($relsXml->Relationship as $rel) {
                $rels[(string) $rel['Id']] = (string) $rel['Target'];
            }
        }

        $documentXml = str_replace('w:', '', $documentXml);
        $document = simplexml_load_string($documentXml);
        if (!$document) {
            return null;
        }

        $html = self::parseBody($document->body, $rels, $mediaFiles);

        return "<div class=\"docx-content\">$html</div>";
    }

    protected static function parseBody($body, $rels, $mediaFiles): string
    {
        $html = '';
        foreach ($body->children() as $node) {
            switch ($node->getName()) {
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
        return $html;
    }

    protected static function parseParagraph($p): string
    {
        $align = (string) $p->pPr->jc['val'] ?? 'left';
        $indentLeft = (int) $p->pPr->ind['left'] ?? 0;

        $style = "text-align:$align;";
        if ($indentLeft > 0) {
            $px = intval($indentLeft / 20); // приблизительно 1px ~ 20pt
            $style .= "padding-left:{$px}px;";
        }

        $html = '';
        foreach ($p->r as $run) {
            $text = (string) $run->t;
            if (trim($text) === '') continue;

            $bold = isset($run->rPr->b);
            $italic = isset($run->rPr->i);
            $underline = isset($run->rPr->u);

            $styles = [];
            if ($bold) $styles[] = 'font-weight:bold';
            if ($italic) $styles[] = 'font-style:italic';
            if ($underline) $styles[] = 'text-decoration:underline';

            $styleAttr = $styles ? ' style="' . implode(';', $styles) . '"' : '';
            $html .= "<span$styleAttr>" . htmlspecialchars($text) . "</span>";
        }

        if (isset($p->pPr->numPr)) {
            // List paragraph
            return "<li style=\"$style\">$html</li>";
        }

        return "<p style=\"$style\">$html</p>";
    }

    protected static function parseTable($tbl): string
    {
        $html = '<table border="1" cellspacing="0" cellpadding="4">';
        foreach ($tbl->tr as $tr) {
            $html .= '<tr>';
            foreach ($tr->tc as $tc) {
                $content = '';
                foreach ($tc->p as $p) {
                    $content .= self::parseParagraph($p);
                }
                $html .= "<td>$content</td>";
            }
            $html .= '</tr>';
        }
        $html .= '</table>';
        return $html;
    }

    protected static function parseStructuredDocumentTag($sdt, $rels, $mediaFiles): string
    {
        $html = '';
        foreach ($sdt->xpath('.//drawing') as $drawing) {
            $blip = $drawing->xpath('.//a:blip')[0] ?? null;
            if ($blip && $blip->attributes('r', true)['embed']) {
                $rid = (string) $blip->attributes('r', true)['embed'];
                $target = $rels[$rid] ?? null;
                if ($target && isset($mediaFiles["word/$target"])) {
                    $ext = pathinfo($target, PATHINFO_EXTENSION);
                    $base64 = $mediaFiles["word/$target"];
                    $html .= "<img src=\"data:image/$ext;base64,$base64\" alt=\"image\" />";
                }
            }
        }
        return $html;
    }

    protected static function convertDocToHtmlViaLibreOffice(string $filePath): ?string
    {
        $outputDir = sys_get_temp_dir() . '/doc_' . uniqid();
        if (!mkdir($outputDir, 0777, true)) {
            return null;
        }

        $cmd = sprintf(
            'libreoffice --headless --convert-to html --outdir %s %s',
            escapeshellarg($outputDir),
            escapeshellarg($filePath)
        );

        exec($cmd, $output, $code);
        if ($code !== 0) {
            return null;
        }

        $basename = pathinfo($filePath, PATHINFO_FILENAME) . '.html';
        $htmlPath = $outputDir . '/' . $basename;

        if (!file_exists($htmlPath)) {
            return null;
        }

        $html = file_get_contents($htmlPath);
        unlink($htmlPath);
        rmdir($outputDir);

        return $html ?: null;
    }
}
