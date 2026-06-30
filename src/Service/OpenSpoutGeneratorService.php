<?php

namespace App\Service;

use App\Component\Report\Column;
use App\Component\Report\ReportInterface;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Border;
use OpenSpout\Common\Entity\Style\BorderName;
use OpenSpout\Common\Entity\Style\BorderPart;
use OpenSpout\Common\Entity\Style\BorderStyle;
use OpenSpout\Common\Entity\Style\BorderWidth;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\CellVerticalAlignment;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\XLSX\Entity\SheetView;
use OpenSpout\Writer\XLSX\Options;
use OpenSpout\Writer\XLSX\Writer;

class OpenSpoutGeneratorService
{
    private int $currentRow = 1;
    private int $contentStartRow = 0;
    private int $lastRow = 1;

    public function writeData(ReportInterface $report, string $outputPath, \Generator $dataGenerator): array
    {
        $populateStartTime = microtime(true);
        // 1. Initialize Options with default fallback style
        $fallbackStyle = new Style(
            fontSize: 9,
            fontName: 'Arial'
        );

        $options = new Options($fallbackStyle);
        $writer = new Writer($options);
        $writer->openToFile($outputPath);

        // 2. Setup Sheets
        // DataSheet
        $dataSheet = $writer->getCurrentSheet();
        $dataSheet->setName('DataSheet');

        // ConfigSheet
        $configSheet = $writer->addNewSheetAndMakeItCurrent();
        $configSheet->setName('ConfigSheet');
        $writer->addRow(Row::fromValues(['Tax Rate', 0.15]));

        // Switch back to DataSheet
        $writer->setCurrentSheet($dataSheet);

        // Write 4 empty rows to align with easy-excel (which leaves rows 1-4 empty for the logo)
        for ($i = 1; $i <= 4; $i++) {
            $writer->addRow(new Row([], 20));
        }

        // 3. Pre-build Border Styles
        $metadata = $report->getMetadata();
        $borderStyleConfig = $metadata['style']['border'] ?? 'thin';
        $borderStyleName = $borderStyleConfig === 'dotted' ? BorderStyle::DOTTED : BorderStyle::SOLID;
        $borderWidthName = BorderWidth::THIN;

        $thinBorder = new Border(
            new BorderPart(BorderName::LEFT, '000000', $borderWidthName, $borderStyleName),
            new BorderPart(BorderName::RIGHT, '000000', $borderWidthName, $borderStyleName),
            new BorderPart(BorderName::TOP, '000000', $borderWidthName, $borderStyleName),
            new BorderPart(BorderName::BOTTOM, '000000', $borderWidthName, $borderStyleName)
        );

        // 4. Pre-build Column Styles to minimize memory allocations
        $colStyles = [];
        $highlightColStyles = [];
        $columns = $report->getColumns();
        
        foreach ($columns as $index => $column) {
            $align = CellAlignment::LEFT;
            $format = null;

            if ($column->getFormat()) {
                $format = $column->getFormat();
            }

            switch ($column->getType()) {
                case Column::TYPE_FLOAT:
                    $align = CellAlignment::RIGHT;
                    if (!$format) {
                        $format = '$#,##0.00';
                    }
                    break;
                case Column::TYPE_INTEGER:
                    $align = CellAlignment::CENTER;
                    break;
                case Column::TYPE_DATE:
                    $align = CellAlignment::LEFT;
                    if (!$format) {
                        $format = 'yyyy-mm-dd';
                    }
                    break;
            }

            if ($column->getIsLargeNumber()) {
                $align = CellAlignment::RIGHT;
            }

            // Wrapping text setting
            $wrapText = $column->isWrapText() ? true : null;

            $colStyles[$index] = new Style(
                fontSize: 9,
                fontName: 'Arial',
                cellAlignment: $align,
                cellVerticalAlignment: CellVerticalAlignment::CENTER,
                shouldWrapText: $wrapText,
                border: $thinBorder,
                format: $format
            );

            // Highlighted version (for quantity > 90 simulation)
            $highlightColStyles[$index] = $colStyles[$index]
                ->withFontBold(true)
                ->withBackgroundColor('E2EFDA'); // Soft green background
        }

        // Pre-build bold/subtotal versions of column styles
        $boldColStyles = [];
        foreach ($colStyles as $index => $style) {
            $boldColStyles[$index] = $style->withFontBold(true);
        }

        // Set column widths
        $colIndex = 1;
        foreach ($columns as $column) {
            $width = $column->getWidth();
            if ($width !== null) {
                $dataSheet->setColumnWidth($width, $colIndex);
            }
            $colIndex++;
        }

        $this->currentRow = 5; // Start row below logo space

        // 5. Render Main Headers
        $headerTitleStyle = new Style(
            fontBold: true,
            fontSize: 9,
            fontName: 'Arial',
            cellAlignment: CellAlignment::CENTER,
            cellVerticalAlignment: CellVerticalAlignment::CENTER
        );

        // Row 5: BRAC
        $writer->addRow($this->createHeaderRow('BRAC', $headerTitleStyle, 8, 20));
        $options->mergeCells(0, 5, 7, 5);
        $this->currentRow++;

        // Row 6: Title
        $writer->addRow($this->createHeaderRow($report->getDefinition()->getTitle(), $headerTitleStyle, 8, 20));
        $options->mergeCells(0, 6, 7, 6);
        $this->currentRow++;

        // Row 7: Parameters
        if (!empty($report->getHeaders())) {
            $writer->addRow($this->createHeaderRow($report->getHeaders()[0][0], $headerTitleStyle, 8, 20));
            $options->mergeCells(0, 7, 7, 7);
            $this->currentRow++;
        }

        // Row 8: Table Header
        $tableHeaderBg = $report->getReportConfig()->getMetaDataByArrayKey('tableHeaderBg') ?? 'FF1F4E78';
        $rgbHeaderBg = strlen($tableHeaderBg) === 8 ? substr($tableHeaderBg, 2) : $tableHeaderBg;

        $tableHeaderStyle = new Style(
            fontBold: true,
            fontSize: 9,
            fontName: 'Arial',
            fontColor: 'FFFFFF',
            backgroundColor: $rgbHeaderBg,
            cellAlignment: CellAlignment::CENTER,
            cellVerticalAlignment: CellVerticalAlignment::CENTER,
            border: $thinBorder
        );

        $reportHeader = $report->getDefinition()->getReportHeader();
        $writer->addRow($this->createHeaderRow($reportHeader[0][0]['label'], $tableHeaderStyle, 8, 20));
        $options->mergeCells(0, 8, 7, 8);
        $this->currentRow++;

        $this->contentStartRow = $this->currentRow; // Content starts at Row 9

        // 6. Write Data Rows
        $subtotalInterval = 50000;
        $blockStartRow = $this->currentRow;
        $subTotalRows = [];
        $rowCount = 0;

        foreach ($dataGenerator as $rowData) {
            $quantity = $rowData['quantity'];
            $isHighlighted = $quantity > 90;
            
            // Choose styles array
            $activeStyles = $isHighlighted ? $highlightColStyles : $colStyles;

            // Row Group Merging Feature Demo (apply bold style and merge cells)
            $isMergedRow = ($rowCount > 0 && $rowCount <= 100 && $rowCount % 20 === 0);
            if ($isMergedRow) {
                // If it is a merged row, we merge columns A-H (0 to 7) and make the row bold
                $cells = [
                    Cell::fromValue($rowData['id'], $boldColStyles[0]),
                    Cell::fromValue($rowData['name'], $boldColStyles[1]),
                    Cell::fromValue($rowData['category'], $boldColStyles[2]),
                    Cell::fromValue($rowData['quantity'], $boldColStyles[3]),
                    Cell::fromValue($rowData['price'], $boldColStyles[4]),
                    Cell::fromValue("=D{$this->currentRow}*E{$this->currentRow}", $boldColStyles[5]),
                    Cell::fromValue("=F{$this->currentRow}*ConfigSheet!\$B\$1", $boldColStyles[6]),
                    Cell::fromValue($rowData['date'], $boldColStyles[7]),
                ];
                $writer->addRow(new Row($cells, 20));
                $options->mergeCells(0, $this->currentRow, 7, $this->currentRow);
            } else {
                // Standard data row
                $cells = [
                    Cell::fromValue($rowData['id'], $activeStyles[0]),
                    Cell::fromValue($rowData['name'], $activeStyles[1]),
                    Cell::fromValue($rowData['category'], $activeStyles[2]),
                    Cell::fromValue($rowData['quantity'], $activeStyles[3]),
                    Cell::fromValue($rowData['price'], $activeStyles[4]),
                    Cell::fromValue("=D{$this->currentRow}*E{$this->currentRow}", $activeStyles[5]),
                    Cell::fromValue("=F{$this->currentRow}*ConfigSheet!\$B\$1", $activeStyles[6]),
                    Cell::fromValue($rowData['date'], $activeStyles[7]),
                ];
                $writer->addRow(new Row($cells, 20));
            }

            $this->currentRow++;
            $rowCount++;

            // Subtotal row insertion periodically
            if ($rowCount % $subtotalInterval === 0) {
                $subTotalRows[] = $this->currentRow;
                $this->writeSubtotalRow($writer, $blockStartRow, $this->currentRow - 1, $boldColStyles, $options);
                $this->currentRow++;
                $blockStartRow = $this->currentRow;
            }

            // Periodic Garbage Collection to free memory
            if ($rowCount % 10000 === 0) {
                gc_collect_cycles();
            }
        }

        // Write final subtotal if any leftover data
        if ($this->currentRow > $blockStartRow) {
            $subTotalRows[] = $this->currentRow;
            $this->writeSubtotalRow($writer, $blockStartRow, $this->currentRow - 1, $boldColStyles, $options);
            $this->currentRow++;
        }

        $this->lastRow = $this->currentRow - 1;

        // 7. Render Grand Total Row
        if ($report->getDefinition()->hasTotal()) {
            $cells = [
                Cell::fromValue('Grand Total', $boldColStyles[0]),
                Cell::fromValue('', $boldColStyles[1]),
                Cell::fromValue('', $boldColStyles[2]),
                Cell::fromValue("=SUM(D{$this->contentStartRow}:D{$this->lastRow})/2", $boldColStyles[3]),
                Cell::fromValue("=SUM(E{$this->contentStartRow}:E{$this->lastRow})/2", $boldColStyles[4]),
                Cell::fromValue("=SUM(F{$this->contentStartRow}:F{$this->lastRow})/2", $boldColStyles[5]),
                Cell::fromValue("=SUM(G{$this->contentStartRow}:G{$this->lastRow})/2", $boldColStyles[6]),
                Cell::fromValue('', $boldColStyles[7]),
            ];
            $writer->addRow(new Row($cells, 20));
            $options->mergeCells(0, $this->currentRow, 2, $this->currentRow);
            $this->currentRow++;
        }

        // 8. Signatory block
        $signatory = $report->getReportConfig()->getMetaDataByArrayKey('signatory');
        if (!empty($signatory)) {
            $this->currentRow += 3;
            // Prepared By label
            $writer->addRow(Row::fromValues([$report->getUserFullName()]));
            $this->currentRow++;

            // Signatory Titles
            foreach ($signatory as $sigRow) {
                $writer->addRow(Row::fromValues($sigRow));
                $this->currentRow++;
            }
        }

        // 9. Freeze Panes
        $freezeConfig = $metadata['freeze'] ?? 'A10';
        if ($freezeConfig) {
            // Find row number from freezeConfig (e.g. 'A10' -> 10)
            preg_match('/\d+/', $freezeConfig, $matches);
            $freezeRow = !empty($matches) ? (int)$matches[0] : 10;
            $sheetView = new SheetView(freezeRow: $freezeRow, freezeColumn: 'A');
            $dataSheet->setSheetView($sheetView);
        }
        $populateEndTime = microtime(true);
        $populatePeakMemory = memory_get_peak_usage();

        $saveStartTime = microtime(true);

        // Close the writer
        $writer->close();

        $saveEndTime = microtime(true);
        $savePeakMemory = memory_get_peak_usage();

        return [
            'populate_time' => $populateEndTime - $populateStartTime,
            'populate_memory' => $populatePeakMemory,
            'save_time' => $saveEndTime - $saveStartTime,
            'save_memory' => $savePeakMemory,
        ];
    }

    private function createHeaderRow(string $label, Style $style, int $colCount, float $height): Row
    {
        $cells = [];
        for ($i = 0; $i < $colCount; $i++) {
            $cells[] = Cell::fromValue($i === 0 ? $label : '', $style);
        }
        return new Row($cells, $height);
    }

    private function writeSubtotalRow(Writer $writer, int $start, int $end, array $boldStyles, Options $options): void
    {
        $cells = [
            Cell::fromValue('Sub-total', $boldStyles[0]),
            Cell::fromValue('', $boldStyles[1]),
            Cell::fromValue('', $boldStyles[2]),
            Cell::fromValue("=SUM(D{$start}:D{$end})", $boldStyles[3]),
            Cell::fromValue("=SUM(E{$start}:E{$end})", $boldStyles[4]),
            Cell::fromValue("=SUM(F{$start}:F{$end})", $boldStyles[5]),
            Cell::fromValue("=SUM(G{$start}:G{$end})", $boldStyles[6]),
            Cell::fromValue('', $boldStyles[7]),
        ];
        
        $writer->addRow(new Row($cells, 20));
        $options->mergeCells(0, $this->currentRow, 2, $this->currentRow);
    }
}
