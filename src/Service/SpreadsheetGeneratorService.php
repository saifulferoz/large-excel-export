<?php

namespace App\Service;

use App\Component\Report\Column;
use App\Component\Report\PhpOffice\StringValueBinder;
use App\Component\Report\ReportInterface;
use App\Utility\SpreadsheetUtil;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Conditional;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Settings;
use Symfony\Component\Cache\Adapter\FilesystemAdapter;
use Symfony\Component\Cache\Psr16Cache;

class SpreadsheetGeneratorService
{
    protected int $currentRow = 1;
    private int $contentStartRow = 0;
    private string $lastColumn = 'A';
    private array $borderStyle = ['dotted' => Border::BORDER_DOTTED, 'thin' => Border::BORDER_THIN];
    private int $lastRow = 1;
    private ?string $subTotalLabelEndColumn = null;

    private ?Spreadsheet $spreadsheet = null;

    public function __construct()
    {
        // Set up PSR-16 Cell Cache using Filesystem Adapter to cap memory usage
        // This is critical for 400,000 rows
        // $cachePool = new FilesystemAdapter('spreadsheet_cells', 0, sys_get_temp_dir());
        // $psr16Cache = new Psr16Cache($cachePool);
        // Settings::setCache($psr16Cache);
    }

    public function writeData(ReportInterface $report, Spreadsheet $spreadsheet, \Generator $dataGenerator): void
    {
        $this->spreadsheet = $spreadsheet;
        // 1. Apply Default Style
        $spreadsheet->getDefaultStyle()->applyFromArray([
            'font' => [
                'size' => 9,
                'name' => 'Arial',
            ],
        ]);
        
        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet->setTitle('DataSheet');
        Cell::setValueBinder(new StringValueBinder());
        
        // 2. Set Metadata Properties
        $spreadsheet->getProperties()
            ->setTitle($report->getDefinition()->getTitle())
            ->setSubject($report->getDefinition()->getTitle());

        // 3. Create secondary sheet for references (ConfigSheet)
        $configSheet = $spreadsheet->createSheet();
        $configSheet->setTitle('ConfigSheet');
        $configSheet->setCellValue('A1', 'Tax Rate');
        $configSheet->setCellValue('B1', 0.15); // 15% tax rate
        
        // Switch back to active sheet
        $spreadsheet->setActiveSheetIndex(0);
        $worksheet = $spreadsheet->getActiveSheet();

        // 4. Image Insertion (Logo) in Header
        $this->insertLogo($worksheet);

        $this->currentRow = 5; // Move down below the logo
        $this->setPreviousColumnOfFirstSubTotalColumn($report);

        // 5. Render Main Headers
        if (!empty($report->getHeaders())) {
            $worksheet->setCellValue('A' . $this->currentRow, 'BRAC');
            $worksheet->setCellValue('A' . ($this->currentRow + 1), $report->getDefinition()->getTitle());
            $worksheet->getStyle('A' . $this->currentRow . ':A' . ($this->currentRow + 1))->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setVertical(Alignment::VERTICAL_CENTER);
            $worksheet->fromArray($report->getHeaders(), null, 'A' . ($this->currentRow + 2));
            $this->currentRow += count($report->getHeaders()) + 2;
        } else {
            $worksheet->setCellValue('A' . $this->currentRow, 'BRAC');
            $worksheet->setCellValue('A' . ($this->currentRow + 1), $report->getDefinition()->getTitle());
            $worksheet->getStyle('A' . $this->currentRow . ':A' . ($this->currentRow + 1))->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setVertical(Alignment::VERTICAL_CENTER);
            $this->currentRow += 2;
        }

        // 6. Table Headers
        $this->addTableHeaderRow($worksheet, $report);
        $this->contentStartRow = $this->currentRow;
        
        $lastCol = $worksheet->getHighestColumn($this->currentRow - 1);
        $this->lastColumn = $lastCol;
        
        $worksheet->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(
            $this->currentRow - 1,
            $this->currentRow - 1
        );
        $worksheet->getStyle(
            'A' . ($this->currentRow - 1) . ':' . $lastCol . ($this->currentRow - 1)
        )->getAlignment()
            ->setVertical(Alignment::VERTICAL_CENTER);

        // Merge headers A1/A2 etc.
        $worksheet->mergeCells("A5:{$lastCol}5");
        $worksheet->mergeCells("A6:{$lastCol}6");
        if (!empty($report->getHeaders())) {
            $numberOfHeaderRow = 6 + count($report->getHeaders());
            for ($i = 7; $i <= $numberOfHeaderRow; ++$i) {
                $worksheet->mergeCells("A{$i}:{$lastCol}{$i}");
            }
        }

        // 7. Write Data Rows
        $subtotalInterval = 50000;
        $blockStartRow = $this->currentRow;
        $subTotalRows = [];
        $mergedRows = [];
        $rowCount = 0;

        foreach ($dataGenerator as $rowData) {
            $worksheet->setCellValue('A' . $this->currentRow, $rowData['id']);
            $worksheet->setCellValue('B' . $this->currentRow, $rowData['name']);
            $worksheet->setCellValue('C' . $this->currentRow, $rowData['category']);
            $worksheet->setCellValue('D' . $this->currentRow, $rowData['quantity']);
            $worksheet->setCellValue('E' . $this->currentRow, $rowData['price']);
            
            // Formula 1: Subtotal = Quantity * Price
            $worksheet->setCellValue('F' . $this->currentRow, "=D{$this->currentRow}*E{$this->currentRow}");
            
            // Formula 2: Tax = Subtotal * ConfigSheet!$B$1 (Reference to another sheet)
            $worksheet->setCellValue('G' . $this->currentRow, "=F{$this->currentRow}*ConfigSheet!\$B\$1");
            
            $worksheet->setCellValue('H' . $this->currentRow, $rowData['date']);

            // Row Group Merging Feature Demo (limit to 5 rows for demo to avoid OOM/slowness)
            if ($rowCount > 0 && $rowCount <= 100 && $rowCount % 20 === 0) {
                $mergedRows[] = $this->currentRow - $this->contentStartRow;
            }

            $this->currentRow++;
            $rowCount++;

            // Subtotal row insertion periodically
            if ($rowCount % $subtotalInterval === 0) {
                $subTotalRows[] = $this->currentRow - $this->contentStartRow;
                $this->writeSubtotalRow($worksheet, $blockStartRow, $this->currentRow - 1, $report);
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
            $subTotalRows[] = $this->currentRow - $this->contentStartRow;
            $this->writeSubtotalRow($worksheet, $blockStartRow, $this->currentRow - 1, $report);
            $this->currentRow++;
        }

        $this->lastRow = $this->currentRow - 1;
        $highestRow = $this->lastRow;

        // 8. Render Grand Total Row
        if ($report->getDefinition()->hasTotal()) {
            $worksheet->setCellValue('A' . $this->currentRow, 'Grand Total');
            
            // Formulas summing all subtotal rows (column sum divided by 2 since we have subtotal rows in the column)
            $worksheet->setCellValue('D' . $this->currentRow, "=SUM(D{$this->contentStartRow}:D{$this->lastRow})/2");
            $worksheet->setCellValue('E' . $this->currentRow, "=SUM(E{$this->contentStartRow}:E{$this->lastRow})/2");
            $worksheet->setCellValue('F' . $this->currentRow, "=SUM(F{$this->contentStartRow}:F{$this->lastRow})/2");
            $worksheet->setCellValue('G' . $this->currentRow, "=SUM(G{$this->contentStartRow}:G{$this->lastRow})/2");

            $worksheet->getStyle("A{$this->currentRow}:{$lastCol}{$this->currentRow}")->getFont()->setBold(true);
            $worksheet->mergeCells(
                "A{$this->currentRow}:{$this->getPreviousColumnOfFirstTotalColumn($report)}{$this->currentRow}"
            );
            $worksheet->getStyle("A{$this->currentRow}")->getAlignment()->setHorizontal(
                Alignment::HORIZONTAL_CENTER
            );
            $this->currentRow++;
            $highestRow = $this->currentRow - 1;
        }

        // Apply Border styling to columns instead of ranges to save memory
        $metadata = $report->getMetadata();
        $borderStyle = $metadata['style']['border'] ?? 'thin';
        $worksheet->getStyle("A:{$lastCol}")->getBorders()
            ->getAllBorders()->setBorderStyle($this->borderStyle[$borderStyle]);

        $worksheet->getStyle(
            "A1:{$this->lastColumn}{$highestRow}"
        )->getAlignment()->setVertical(
            Alignment::VERTICAL_CENTER
        );

        // 9. Apply Column Formats & Alignments
        $this->applyCellFormat($report, $highestRow);

        // 10. Wrap Text (using column level style to save memory)
        foreach (array_keys($report->getColumnLabels()) as $index => $item) {
            $column = $report->getDefinition()->getColumn($item);
            $excelColumn = SpreadsheetUtil::getColumnByIndex($index + 1);
            if ($column->isWrapText()) {
                $worksheet->getStyle($excelColumn)->getAlignment()->setWrapText(true);
            }
        }

        // Set column widths
        $this->setSheetColumnWidth($worksheet, $report->getColumnLabels(), $report);

        // 11. Row Merging Groups (simulating getMergedRow() action)
        foreach ($mergedRows as $index) {
            $rowNo = $this->contentStartRow + $index;
            $startColumn = 'A';
            if (isset($metadata['rowGroup']) && isset($metadata['rowGroup']['serial']) && $metadata['rowGroup']['serial']) {
                $startColumn = 'B';
            }
            $worksheet->mergeCells($startColumn . $rowNo . ':' . $this->lastColumn . $rowNo);
            $worksheet->getStyle('A' . ($rowNo))->getFont()->setBold(true);
        }

        // 12. Style Sub-totals periodically
        $skipMerge = $metadata['skipSubTotalColumnMerge'] ?? false;
        foreach ($subTotalRows as $row) {
            $actualRow = $this->contentStartRow + $row;
            if (!$skipMerge) {
                $worksheet->mergeCells(
                    'A' . $actualRow . ':' . $this->subTotalLabelEndColumn . $actualRow
                );
            }
            $worksheet->getStyle('A' . $actualRow)->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $worksheet->getStyle(
                'A' . $actualRow . ':' . $this->lastColumn . $actualRow
            )->getFont()->setBold(true);
        }

        // 13. Conditional Styles
        $this->applyConditionalStyle($worksheet, $report, $highestRow);

        // 14. Signatory block
        $signatory = $report->getReportConfig()->getMetaDataByArrayKey('signatory');
        if (!empty($signatory)) {
            $this->currentRow += 3;
            $worksheet->setCellValue('A' . ($this->currentRow - 1), $report->getUserFullName());
            $worksheet->fromArray($signatory, null, 'A' . $this->currentRow);
        }

        // 15. Row Heights
        if (isset($metadata['rowHeight'])) {
            for ($row = 1; $row <= $highestRow; ++$row) {
                $worksheet->getRowDimension($row)->setRowHeight($metadata['rowHeight']);
            }
        }

        // 16. Freeze panes
        if (isset($metadata['freeze'])) {
            $worksheet->freezePane($metadata['freeze']);
        }

        // 17. Background color
        if (isset($metadata['bgColor'])) {
            $worksheet->getStyle('A1:' . $this->lastColumn . $highestRow)
                ->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setARGB($metadata['bgColor']);
        }
    }

    private function writeSubtotalRow(Worksheet $worksheet, int $start, int $end, ReportInterface $report): void
    {
        $worksheet->setCellValue('A' . $this->currentRow, 'Sub-total');
        
        // Sum formulas for each numeric column
        $worksheet->setCellValue('D' . $this->currentRow, "=SUM(D{$start}:D{$end})");
        $worksheet->setCellValue('E' . $this->currentRow, "=SUM(E{$start}:E{$end})");
        $worksheet->setCellValue('F' . $this->currentRow, "=SUM(F{$start}:F{$end})");
        $worksheet->setCellValue('G' . $this->currentRow, "=SUM(G{$start}:G{$end})");
    }

    private function applyCellFormat(ReportInterface $report, int $highestRow): void
    {
        $worksheet = $this->spreadsheet->getActiveSheet();
        $col = 1;
        foreach (array_keys($report->getColumnLabels()) as $key) {
            $column = $report->getDefinition()->getColumn($key);
            $excelCol = SpreadsheetUtil::getColumnByIndex($col);
            $range = $excelCol; // Optimized: Style the column instead of the row range

            if ($column->getFormat()) {
                $worksheet->getStyle($range)->getNumberFormat()->setFormatCode($column->getFormat());
                ++$col;
                continue;
            }
            if ($column->getIsLargeNumber()) {
                $worksheet->getStyle($range)->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_RIGHT)
                    ->setVertical(Alignment::VERTICAL_CENTER);
                ++$col;
                continue;
            }
            switch ($column->getType()) {
                case Column::TYPE_FLOAT:
                    $format = SpreadsheetUtil::ACCOUNTING_FORMAT;
                    $worksheet->getStyle($range)->getNumberFormat()->setFormatCode($format);
                    break;
                case Column::TYPE_DATE:
                    $format = empty($column->getFormat()) ? 'dd-mm-yyyy' : $column->getFormat();
                    $worksheet->getStyle($range)->getNumberFormat()->setFormatCode($format);
                    break;
                case Column::TYPE_INTEGER:
                    $worksheet->getStyle($range)->getAlignment()
                        ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                        ->setVertical(Alignment::VERTICAL_CENTER);
                    break;
                default:
                    $worksheet->getStyle($range)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_TEXT);
            }
            ++$col;
        }
    }

    private function applyConditionalStyle(Worksheet $worksheet, ReportInterface $report, int $highestRow): void
    {
        $styles = $report->getReportConfig()->getMetaDataByArrayKey('conditionalStyle');
        if (empty($styles)) {
            return;
        }
        
        $conditionalStyles = $worksheet->getStyle("A:{$this->lastColumn}")
            ->getConditionalStyles();
            
        foreach ($styles as $style) {
            $conditional = new Conditional();
            $conditional->setConditionType($style['conditionType'])
                ->setOperatorType($style['operatorType'])
                ->addCondition($style['condition']);
            
            $conditional->getStyle()->getFill()->setFillType(Fill::FILL_SOLID);
            $conditional->getStyle()->getFill()->getStartColor()->setARGB($style['color']);
            $conditional->getStyle()->getFill()->getEndColor()->setARGB($style['color']);
            $conditional->getStyle()->getFont()->setBold(true);
            $conditionalStyles[] = $conditional;
        }
        
        $worksheet->getStyle("A:{$this->lastColumn}")->setConditionalStyles(
            $conditionalStyles
        );
    }

    private function insertLogo(Worksheet $worksheet): void
    {
        // Create public/images directory if it doesn't exist and generate a 10x10 dummy png logo
        $imgDir = dirname(__DIR__, 2) . '/public/images';
        if (!is_dir($imgDir)) {
            mkdir($imgDir, 0777, true);
        }
        
        $logoPath = $imgDir . '/logo.png';
        if (!file_exists($logoPath)) {
            // Generate dummy 100x40 image
            $im = imagecreatetruecolor(100, 40);
            $bg = imagecolorallocate($im, 74, 144, 226); // Nice blue color
            $white = imagecolorallocate($im, 255, 255, 255);
            imagefilledrectangle($im, 0, 0, 99, 39, $bg);
            imagestring($im, 4, 15, 12, 'BRAC LOGO', $white);
            imagepng($im, $logoPath);
            imagedestroy($im);
        }

        $drawing = new Drawing();
        $drawing->setName('BRAC Logo');
        $drawing->setDescription('BRAC Company Logo');
        $drawing->setPath($logoPath);
        $drawing->setHeight(30);
        $drawing->setCoordinates('A1');
        $drawing->setOffsetX(10);
        $drawing->setOffsetY(10);
        $drawing->setWorksheet($worksheet);
    }

    public function setPreviousColumnOfFirstSubTotalColumn(ReportInterface $report): void
    {
        $columnIndex = 0;
        $columnLabels = $report->getColumnLabels();
        foreach (array_keys($columnLabels) as $key) {
            $column = $report->getDefinition()->getColumn($key);
            if ($column->isSubTotal()) {
                break;
            }
            ++$columnIndex;
        }
        $this->subTotalLabelEndColumn = Coordinate::stringFromColumnIndex($columnIndex);
    }

    public function getPreviousColumnOfFirstTotalColumn(ReportInterface $report): string
    {
        $columnIndex = 0;
        $selectedColumns = $report->getColumnLabels();
        foreach ($report->getColumns() as $column) {
            if (!isset($selectedColumns[$column->getName()])) {
                continue;
            }
            if ($column->isTotal()) {
                break;
            }
            ++$columnIndex;
        }
        return Coordinate::stringFromColumnIndex($columnIndex);
    }

    private function setSheetColumnWidth(Worksheet $worksheet, array $row, ReportInterface $report): void
    {
        $col = 'A';
        foreach (array_keys($row) as $key) {
            $column = $report->getDefinition()->getColumn($key);
            if (null !== $column->getWidth()) {
                $worksheet->getColumnDimension($col)->setWidth($column->getWidth());
            } else {
                $worksheet->getColumnDimension($col)->setAutoSize(true);
            }
            $col = SpreadsheetUtil::getNextColumn($col);
        }
    }

    protected function addTableHeaderRow(Worksheet $sheet, ReportInterface $report): self
    {
        $reportHeader = $report->getDefinition()->getReportHeader();
        $col = 'A';
        $startRow = $this->currentRow;
        
        foreach ($reportHeader[0] as $key => $row) {
            $sheet->setCellValue($col . $this->currentRow, $row['label']);
            $sheet->getStyle($col . $this->currentRow)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setVertical(Alignment::VERTICAL_CENTER);
            $sheet->getStyle($col . $this->currentRow)->getFont()->setBold(true);
            
            if (isset($reportHeader[1])) {
                $sheet->setCellValue($col . ($this->currentRow + 1), $reportHeader[1][$key]['label']);
                $sheet->getStyle($col . ($this->currentRow + 1))->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                    ->setVertical(Alignment::VERTICAL_CENTER);
                $sheet->getStyle($col . ($this->currentRow + 1))->getFont()->setBold(true);
                if (isset($reportHeader[1][$key]['colspan'])) {
                    $sheet->mergeCells(
                        $col . ($this->currentRow + 1) . ':' . SpreadsheetUtil::getTargetColumn(
                            $col,
                            $reportHeader[1][$key]['colspan']
                        ) . ($this->currentRow + 1)
                    );
                }
            }
            
            if (isset($reportHeader[2])) {
                $sheet->setCellValue($col . ($this->currentRow + 2), $reportHeader[2][$key]['label']);
                $sheet->getStyle($col . ($this->currentRow + 2))->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                    ->setVertical(Alignment::VERTICAL_CENTER);
                $sheet->getStyle($col . ($this->currentRow + 2))->getFont()->setBold(true);
            }
            
            if (isset($row['rowspan'])) {
                $sheet->mergeCells($col . $this->currentRow . ':' . $col . ($this->currentRow + $row['rowspan']));
            }
            if (isset($row['colspan'])) {
                $sheet->mergeCells(
                    $col . $this->currentRow . ':' . SpreadsheetUtil::getTargetColumn(
                        $col,
                        $row['colspan']
                    ) . ($this->currentRow)
                );
            }
            $this->lastColumn = $col;
            $col = SpreadsheetUtil::getNextColumn($col);
        }
        
        $endRow = $this->currentRow + count($reportHeader) - 1;
        $bgColor = $report->getReportConfig()->getMetaDataByArrayKey('tableHeaderBg');
        if (null !== $bgColor) {
            $sheet->getStyle("A{$startRow}:{$this->lastColumn}{$endRow}")->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()
                ->setARGB($bgColor);
            $sheet->getStyle("A{$startRow}:{$this->lastColumn}{$endRow}")->getFont()->setColor(
                new Color(Color::COLOR_WHITE)
            );
        }
        
        $rotation = $report->getColumnHeaderTextRotation();
        if (null !== $rotation && $rotation > 0) {
            if (!isset($reportHeader[1])) {
                $sheet->getStyle('A' . $this->currentRow . ':' . $col . $this->currentRow)
                    ->getAlignment()->setTextRotation($rotation);
            } else {
                $sheet->getStyle('A' . ($this->currentRow + 1) . ':' . $col . ($this->currentRow + 1))
                    ->getAlignment()->setTextRotation($rotation);
            }
        }
        $this->currentRow += count($reportHeader);

        return $this;
    }
}
