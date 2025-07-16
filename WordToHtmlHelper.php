<?php

namespace Local\Helpers;

class WordToHtmlHelper
{
    /**
     * Преобразует .doc или .docx в HTML без сторонних PHP-библиотек.
     *
     * @param string $filePath Полный путь к файлу.
     * @return string|null
     */
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

    /**
     * Преобразование DOCX -> HTML через парсинг document.xml.
     */
    protected static function convertDocxToHtml(string $filePath): ?string
    {
        $zip = new \ZipArchive();
        if ($zip->open($filePath) !== true) {
            return null;
        }

        $xml = $zip->getFromName('word/document.xml');
        $zip->close();

        if (!$xml) {
            return null;
        }

        $xml = str_replace('w:', '', $xml); // убрать namespace
        $xml = simplexml_load_string($xml);
        if (!$xml) {
            return null;
        }

        $html = '';
        foreach ($xml->body->p as $paragraph) {
            $html .= self::parseParagraph($paragraph);
        }

        return "<div class=\"docx-content\">$html</div>";
    }

    /**
     * Преобразует параграф XML в HTML.
     */
    protected static function parseParagraph(\SimpleXMLElement $p): string
    {
        $html = '';
        foreach ($p->r as $run) {
            $text = (string) $run->t;
            $bold = isset($run->rPr->b);
            $italic = isset($run->rPr->i);
            $underline = isset($run->rPr->u);

            $tag = 'span';
            $styles = [];

            if ($bold) $styles[] = 'font-weight:bold';
            if ($italic) $styles[] = 'font-style:italic';
            if ($underline) $styles[] = 'text-decoration:underline';

            $styleAttr = $styles ? ' style="' . implode(';', $styles) . '"' : '';
            $html .= "<$tag$styleAttr>" . htmlspecialchars($text) . "</$tag>";
        }

        return "<p>$html</p>";
    }

    /**
     * Преобразование .doc в HTML через LibreOffice CLI.
     */
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

        // Удалить временные файлы
        unlink($htmlPath);
        rmdir($outputDir);

        return $html ?: null;
    }
}
