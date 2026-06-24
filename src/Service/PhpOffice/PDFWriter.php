<?php

/*
 * This file is part of the sbiCloud Addons module.
 *
 * Copyright (c) 2019-2022, BRAC IT SERVICES LIMITED <https://www.bracits.com>
 */

namespace App\Service\PhpOffice;

use Knp\Snappy\Pdf;
use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;

class PDFWriter extends HTMLWriter
{
    private ?Pdf $writer = null;

    /**
     * Orientation (Over-ride).
     *
     * @var string
     */
    protected $orientation;

    /**
     * Paper size (Over-ride).
     *
     * @var int
     */
    protected $paperSize;

    /**
     * Temporary storage for Save Array Return type.
     *
     * @var string
     */
    private $saveArrayReturnType;

    /**
     * Paper Sizes xRef List.
     *
     * @var array
     */
    protected static $paperSizes = [
        PageSetup::PAPERSIZE_LETTER => 'LETTER', //    (8.5 in. by 11 in.)
        PageSetup::PAPERSIZE_LETTER_SMALL => 'LETTER', //    (8.5 in. by 11 in.)
        PageSetup::PAPERSIZE_TABLOID => [792.00, 1224.00], //    (11 in. by 17 in.)
        PageSetup::PAPERSIZE_LEDGER => [1224.00, 792.00], //    (17 in. by 11 in.)
        PageSetup::PAPERSIZE_LEGAL => 'LEGAL', //    (8.5 in. by 14 in.)
        PageSetup::PAPERSIZE_STATEMENT => [396.00, 612.00], //    (5.5 in. by 8.5 in.)
        PageSetup::PAPERSIZE_EXECUTIVE => 'EXECUTIVE', //    (7.25 in. by 10.5 in.)
        PageSetup::PAPERSIZE_A3 => 'A3', //    (297 mm by 420 mm)
        PageSetup::PAPERSIZE_A4 => 'A4', //    (210 mm by 297 mm)
        PageSetup::PAPERSIZE_A4_SMALL => 'A4', //    (210 mm by 297 mm)
        PageSetup::PAPERSIZE_A5 => 'A5', //    (148 mm by 210 mm)
        PageSetup::PAPERSIZE_B4 => 'B4', //    (250 mm by 353 mm)
        PageSetup::PAPERSIZE_B5 => 'B5', //    (176 mm by 250 mm)
        PageSetup::PAPERSIZE_FOLIO => 'FOLIO', //    (8.5 in. by 13 in.)
        PageSetup::PAPERSIZE_QUARTO => [609.45, 779.53], //    (215 mm by 275 mm)
        PageSetup::PAPERSIZE_STANDARD_1 => [720.00, 1008.00], //    (10 in. by 14 in.)
        PageSetup::PAPERSIZE_STANDARD_2 => [792.00, 1224.00], //    (11 in. by 17 in.)
        PageSetup::PAPERSIZE_NOTE => 'LETTER', //    (8.5 in. by 11 in.)
        PageSetup::PAPERSIZE_NO9_ENVELOPE => [279.00, 639.00], //    (3.875 in. by 8.875 in.)
        PageSetup::PAPERSIZE_NO10_ENVELOPE => [297.00, 684.00], //    (4.125 in. by 9.5 in.)
        PageSetup::PAPERSIZE_NO11_ENVELOPE => [324.00, 747.00], //    (4.5 in. by 10.375 in.)
        PageSetup::PAPERSIZE_NO12_ENVELOPE => [342.00, 792.00], //    (4.75 in. by 11 in.)
        PageSetup::PAPERSIZE_NO14_ENVELOPE => [360.00, 828.00], //    (5 in. by 11.5 in.)
        PageSetup::PAPERSIZE_C => [1224.00, 1584.00], //    (17 in. by 22 in.)
        PageSetup::PAPERSIZE_D => [1584.00, 2448.00], //    (22 in. by 34 in.)
        PageSetup::PAPERSIZE_E => [2448.00, 3168.00], //    (34 in. by 44 in.)
        PageSetup::PAPERSIZE_DL_ENVELOPE => [311.81, 623.62], //    (110 mm by 220 mm)
        PageSetup::PAPERSIZE_C5_ENVELOPE => 'C5', //    (162 mm by 229 mm)
        PageSetup::PAPERSIZE_C3_ENVELOPE => 'C3', //    (324 mm by 458 mm)
        PageSetup::PAPERSIZE_C4_ENVELOPE => 'C4', //    (229 mm by 324 mm)
        PageSetup::PAPERSIZE_C6_ENVELOPE => 'C6', //    (114 mm by 162 mm)
        PageSetup::PAPERSIZE_C65_ENVELOPE => [323.15, 649.13], //    (114 mm by 229 mm)
        PageSetup::PAPERSIZE_B4_ENVELOPE => 'B4', //    (250 mm by 353 mm)
        PageSetup::PAPERSIZE_B5_ENVELOPE => 'B5', //    (176 mm by 250 mm)
        PageSetup::PAPERSIZE_B6_ENVELOPE => [498.90, 354.33], //    (176 mm by 125 mm)
        PageSetup::PAPERSIZE_ITALY_ENVELOPE => [311.81, 651.97], //    (110 mm by 230 mm)
        PageSetup::PAPERSIZE_MONARCH_ENVELOPE => [279.00, 540.00], //    (3.875 in. by 7.5 in.)
        PageSetup::PAPERSIZE_6_3_4_ENVELOPE => [261.00, 468.00], //    (3.625 in. by 6.5 in.)
        PageSetup::PAPERSIZE_US_STANDARD_FANFOLD => [1071.00, 792.00], //    (14.875 in. by 11 in.)
        PageSetup::PAPERSIZE_GERMAN_STANDARD_FANFOLD => [612.00, 864.00], //    (8.5 in. by 12 in.)
        PageSetup::PAPERSIZE_GERMAN_LEGAL_FANFOLD => 'FOLIO', //    (8.5 in. by 13 in.)
        PageSetup::PAPERSIZE_ISO_B4 => 'B4', //    (250 mm by 353 mm)
        PageSetup::PAPERSIZE_JAPANESE_DOUBLE_POSTCARD => [566.93, 419.53], //    (200 mm by 148 mm)
        PageSetup::PAPERSIZE_STANDARD_PAPER_1 => [648.00, 792.00], //    (9 in. by 11 in.)
        PageSetup::PAPERSIZE_STANDARD_PAPER_2 => [720.00, 792.00], //    (10 in. by 11 in.)
        PageSetup::PAPERSIZE_STANDARD_PAPER_3 => [1080.00, 792.00], //    (15 in. by 11 in.)
        PageSetup::PAPERSIZE_INVITE_ENVELOPE => [623.62, 623.62], //    (220 mm by 220 mm)
        PageSetup::PAPERSIZE_LETTER_EXTRA_PAPER => [667.80, 864.00], //    (9.275 in. by 12 in.)
        PageSetup::PAPERSIZE_LEGAL_EXTRA_PAPER => [667.80, 1080.00], //    (9.275 in. by 15 in.)
        PageSetup::PAPERSIZE_TABLOID_EXTRA_PAPER => [841.68, 1296.00], //    (11.69 in. by 18 in.)
        PageSetup::PAPERSIZE_A4_EXTRA_PAPER => [668.98, 912.76], //    (236 mm by 322 mm)
        PageSetup::PAPERSIZE_LETTER_TRANSVERSE_PAPER => [595.80, 792.00], //    (8.275 in. by 11 in.)
        PageSetup::PAPERSIZE_A4_TRANSVERSE_PAPER => 'A4', //    (210 mm by 297 mm)
        PageSetup::PAPERSIZE_LETTER_EXTRA_TRANSVERSE_PAPER => [667.80, 864.00], //    (9.275 in. by 12 in.)
        PageSetup::PAPERSIZE_SUPERA_SUPERA_A4_PAPER => [643.46, 1009.13], //    (227 mm by 356 mm)
        PageSetup::PAPERSIZE_SUPERB_SUPERB_A3_PAPER => [864.57, 1380.47], //    (305 mm by 487 mm)
        PageSetup::PAPERSIZE_LETTER_PLUS_PAPER => [612.00, 913.68], //    (8.5 in. by 12.69 in.)
        PageSetup::PAPERSIZE_A4_PLUS_PAPER => [595.28, 935.43], //    (210 mm by 330 mm)
        PageSetup::PAPERSIZE_A5_TRANSVERSE_PAPER => 'A5', //    (148 mm by 210 mm)
        PageSetup::PAPERSIZE_JIS_B5_TRANSVERSE_PAPER => [515.91, 728.50], //    (182 mm by 257 mm)
        PageSetup::PAPERSIZE_A3_EXTRA_PAPER => [912.76, 1261.42], //    (322 mm by 445 mm)
        PageSetup::PAPERSIZE_A5_EXTRA_PAPER => [493.23, 666.14], //    (174 mm by 235 mm)
        PageSetup::PAPERSIZE_ISO_B5_EXTRA_PAPER => [569.76, 782.36], //    (201 mm by 276 mm)
        PageSetup::PAPERSIZE_A2_PAPER => 'A2', //    (420 mm by 594 mm)
        PageSetup::PAPERSIZE_A3_TRANSVERSE_PAPER => 'A3', //    (297 mm by 420 mm)
        PageSetup::PAPERSIZE_A3_EXTRA_TRANSVERSE_PAPER => [912.76, 1261.42], //    (322 mm by 445 mm)
    ];

    /**
     * Get Paper Size.
     *
     * @return int
     */
    public function getPaperSize()
    {
        return $this->paperSize;
    }

    /**
     * Set Paper Size.
     *
     * @param string $pValue Paper size see PageSetup::PAPERSIZE_*
     *
     * @return self
     */
    public function setPaperSize($pValue)
    {
        $this->paperSize = $pValue;

        return $this;
    }

    /**
     * Get Orientation.
     */
    public function getOrientation(): string
    {
        return $this->orientation;
    }

    /**
     * Set Orientation.
     *
     * @param string $pValue Page orientation see PageSetup::ORIENTATION_*
     *
     * @return self
     */
    public function setOrientation($pValue)
    {
        $this->orientation = $pValue;

        return $this;
    }

    /**
     * Save Spreadsheet to PDF file, pre-save.
     *
     * @param string $pFilename Name of the file to save as
     *
     * @return resource
     *
     * @throws WriterException
     */
    protected function prepareForSave($pFilename)
    {
        //  garbage collect
        $this->spreadsheet->garbageCollect();

        $this->saveArrayReturnType = Calculation::getArrayReturnType();
        Calculation::setArrayReturnType(Calculation::RETURN_ARRAY_AS_VALUE);

        //  Open file
        $fileHandle = fopen($pFilename, 'w');
        if (false === $fileHandle) {
            throw new WriterException("Could not open file $pFilename for writing.");
        }

        //  Set PDF
        $this->isPdf = true;
        //  Build CSS
        $this->buildCSS(true);

        return $fileHandle;
    }

    /**
     * Save PhpSpreadsheet to PDF file, post-save.
     *
     * @param resource $fileHandle
     */
    protected function restoreStateAfterSave($fileHandle)
    {
        //  Close file
        fclose($fileHandle);

        Calculation::setArrayReturnType($this->saveArrayReturnType);
    }

    /**
     * Save Spreadsheet to file.
     *
     * @param string $pFilename Name of the file to save as
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws PhpSpreadsheetException|\Exception
     */
    public function save($pFilename, int $flags = 0): void
    {
        if (!$this->writer instanceof \Knp\Snappy\Pdf) {
            throw new \Exception('You should set the writer first');
        }

        $fileHandle = $this->prepareForSave($pFilename);

        //  Default PDF paper size
        $paperSize = 'LETTER'; //    Letter    (8.5 in. by 11 in.)

        //  Check for paper size and page orientation
        if (null === $this->getSheetIndex()) {
            $orientation = $this->spreadsheet->getSheet(0)->getPageSetup()->getOrientation();
            $printPaperSize = $this->spreadsheet->getSheet(0)->getPageSetup()->getPaperSize();
        } else {
            $orientation = $this->spreadsheet->getSheet($this->getSheetIndex())->getPageSetup()->getOrientation();
            $printPaperSize = $this->spreadsheet->getSheet($this->getSheetIndex())->getPageSetup()->getPaperSize();
        }
        $this->setOrientation($orientation);

        //  Override Page Orientation
        if (null !== $this->getOrientation()) {
            $orientation = (PageSetup::ORIENTATION_DEFAULT === $this->getOrientation())
                ? PageSetup::ORIENTATION_PORTRAIT
                : $this->getOrientation();
        }
        $orientation = strtoupper($orientation);

        //  Override Paper Size
        if (null !== $this->getPaperSize()) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$paperSizes[$printPaperSize])) {
            $paperSize = self::$paperSizes[$printPaperSize];
        }

        $this->setDefaultOptions($paperSize, $orientation);

        $html = $this->generateHTMLHeader(!$this->getUseInlineCss());
        $html .= $this->generateSheetData();
        $html .= $this->generateHTMLFooter();

        //  Write to file
        fwrite($fileHandle, $this->writer->getOutputFromHtml($html));
        $this->restoreStateAfterSave($fileHandle);
    }

    /**
     * Convert inches to mm.
     *
     * @param float $inches
     *
     * @return float
     */
    private function inchesToMm($inches)
    {
        return $inches * 25.4;
    }

    public function setInternalWriter(Pdf $writer): self
    {
        $writer->setTimeout(900);

        $this->writer = $writer;

        return $this;
    }

    public function getWriter(): ?Pdf
    {
        return $this->writer;
    }

    public function setOptions(array $options)
    {
        foreach ($options as $key => $value) {
            $this->writer->setOption($key, $value);
        }
    }

    private function setDefaultOptions(
        string $paperSize = 'LETTER',
        string $orientation = PageSetup::ORIENTATION_DEFAULT
    ): void {
        $this->writer->setOptions(
            [
                'page-size' => strtoupper($paperSize),
                'orientation' => $orientation,
                'margin-left' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getLeft()),
                'margin-right' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getRight()),
                'margin-top' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getTop()),
                'margin-bottom' => $this->inchesToMm(
                    $this->spreadsheet->getActiveSheet()->getPageMargins()->getBottom()
                ),
                'title' => $this->spreadsheet->getProperties()->getTitle(),
            ]
        );
    }
}
