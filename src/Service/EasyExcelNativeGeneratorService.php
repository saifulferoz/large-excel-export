<?php

namespace App\Service;

use App\Component\Report\Column;
use App\Component\Report\ReportInterface;
use EasyExcel\Native;

class EasyExcelNativeGeneratorService
{
    public function writeData(ReportInterface $report, string $outputPath, \Generator $dataGenerator): array
    {
        $populateStartTime = microtime(true);

        // 1. Initialize Native Workbook
        $handle = Native::newWorkbook();

        // 2. Setup Sheets
        // DataSheet
        $sheets = Native::sheets($handle);
        $dataSheet = 'DataSheet';
        Native::renameSheet($handle, $sheets[0], $dataSheet);

        // ConfigSheet
        $configSheet = 'ConfigSheet';
        Native::addSheet($handle, $configSheet);
        Native::writeRows($handle, $configSheet, 1, 1, [['Tax Rate', 0.15]]);

        // Make DataSheet active
        Native::setActiveSheet($handle, 0);

        // 3. Metadata retrieval
        $metadata = $report->getMetadata();

        // 4. Set Column Widths
        $columns = $report->getColumns();
        foreach ($columns as $index => $column) {
            $width = $column->getWidth();
            if ($width !== null) {
                // Column indexes are 1-based (Col A = 1)
                Native::setColWidth($handle, $dataSheet, $index + 1, $index + 1, $width);
            }
        }

        // 5. Pre-build and Apply Column Styles
        $borderStyleConfig = $metadata['style']['border'] ?? 'thin';
        $baseColStyle = [
            'font' => [
                'size' => 9,
                'name' => 'Arial'
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => $borderStyleConfig,
                    'color' => ['rgb' => '000000']
                ]
            ],
            'alignment' => [
                'vertical' => 'center'
            ]
        ];

        $excelCols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];

        foreach ($columns as $index => $column) {
            $align = 'left';
            $format = null;

            if ($column->getFormat()) {
                $format = $column->getFormat();
            }

            switch ($column->getType()) {
                case Column::TYPE_FLOAT:
                    $align = 'right';
                    if (!$format) {
                        $format = '$#,##0.00';
                    }
                    break;
                case Column::TYPE_INTEGER:
                    $align = 'center';
                    break;
                case Column::TYPE_DATE:
                    $align = 'left';
                    if (!$format) {
                        $format = 'yyyy-mm-dd';
                    }
                    break;
            }

            if ($column->getIsLargeNumber()) {
                $align = 'right';
            }

            $colSpec = $baseColStyle;
            $colSpec['alignment']['horizontal'] = $align;
            if ($column->isWrapText()) {
                $colSpec['alignment']['wrapText'] = true;
            }
            if ($format) {
                $colSpec['numberFormat'] = ['formatCode' => $format];
            }

            Native::applyStyle($handle, $dataSheet, "{$excelCols[$index]}:{$excelCols[$index]}", $colSpec);
        }

        // 6. Insert Logo
        $logoPath = dirname(__DIR__, 2) . '/public/images/logo.png';
        if (file_exists($logoPath)) {
            Native::addImage($handle, $dataSheet, 'A1', [
                'path' => $logoPath,
                'name' => 'BRAC Logo',
                'offsetX' => 10,
                'offsetY' => 10,
                'width' => 100,
                'height' => 30
            ]);
        }

        // 7. Write Main Headers (Row 5 to 7)
        $headerTitleStyle = [
            'font' => [
                'bold' => true,
                'size' => 9,
                'name' => 'Arial'
            ],
            'alignment' => [
                'horizontal' => 'center',
                'vertical' => 'center'
            ]
        ];

        // Row 5: BRAC
        Native::applyStyle($handle, $dataSheet, 'A5:H5', $headerTitleStyle);
        Native::writeRows($handle, $dataSheet, 5, 1, [['BRAC']]);
        Native::mergeCells($handle, $dataSheet, 'A5:H5');

        // Row 6: Title
        Native::applyStyle($handle, $dataSheet, 'A6:H6', $headerTitleStyle);
        Native::writeRows($handle, $dataSheet, 6, 1, [[$report->getDefinition()->getTitle()]]);
        Native::mergeCells($handle, $dataSheet, 'A6:H6');

        // Row 7: Parameters
        if (!empty($report->getHeaders())) {
            Native::applyStyle($handle, $dataSheet, 'A7:H7', $headerTitleStyle);
            Native::writeRows($handle, $dataSheet, 7, 1, [[$report->getHeaders()[0][0]]]);
            Native::mergeCells($handle, $dataSheet, 'A7:H7');
        }

        // Row 8: Table Header
        $reportHeader = $report->getDefinition()->getReportHeader();
        $tableHeaderBg = $report->getReportConfig()->getMetaDataByArrayKey('tableHeaderBg') ?? 'FF1F4E78';
        Native::applyStyle($handle, $dataSheet, 'A8:H8', [
            'font' => [
                'bold' => true,
                'size' => 9,
                'name' => 'Arial',
                'color' => ['rgb' => 'FFFFFF']
            ],
            'fill' => [
                'fillType' => 'solid',
                'startColor' => ['argb' => $tableHeaderBg],
                'endColor' => ['argb' => $tableHeaderBg]
            ],
            'alignment' => [
                'horizontal' => 'center',
                'vertical' => 'center'
            ]
        ]);
        Native::writeRows($handle, $dataSheet, 8, 1, [[$reportHeader[0][0]['label']]]);
        Native::mergeCells($handle, $dataSheet, 'A8:H8');

        // 8. Stream Data Rows in Batches
        $currentRow = 9;
        $contentStartRow = 9;
        $blockStartRow = 9;
        
        $batchSize = 1024; // Lower batch size to prevent CGO memory bloat and OOM
        $batch = [];
        $batchStartRow = 9;
        $rowCount = 0;
        $subtotalInterval = 50000;

        foreach ($dataGenerator as $rowData) {
            $batch[] = [
                $rowData['id'],
                $rowData['name'],
                $rowData['category'],
                $rowData['quantity'],
                $rowData['price'],
                "=D{$currentRow}*E{$currentRow}",
                "=F{$currentRow}*ConfigSheet!\$B\$1",
                $rowData['date']
            ];

            // Row Group Merging Feature Demo
            $isMergedRow = ($rowCount > 0 && $rowCount <= 100 && $rowCount % 20 === 0);
            if ($isMergedRow) {
                // Flush preceding batch
                if ($batch !== []) {
                    Native::writeRows($handle, $dataSheet, $batchStartRow, 1, $batch);
                    $batch = [];
                }
                
                // Merge A-H and make bold
                Native::mergeCells($handle, $dataSheet, "A{$currentRow}:H{$currentRow}");
                Native::applyStyle($handle, $dataSheet, "A{$currentRow}:H{$currentRow}", [
                    'font' => ['bold' => true]
                ]);
                $batchStartRow = $currentRow + 1;
            }

            $currentRow++;
            $rowCount++;

            if (count($batch) >= $batchSize) {
                Native::writeRows($handle, $dataSheet, $batchStartRow, 1, $batch);
                $batch = [];
                $batchStartRow = $currentRow;
            }

            // Intersperse periodic subtotals
            if ($rowCount % $subtotalInterval === 0) {
                if ($batch !== []) {
                    Native::writeRows($handle, $dataSheet, $batchStartRow, 1, $batch);
                    $batch = [];
                }

                $subtotalRow = [
                    'Sub-total',
                    '',
                    '',
                    "=SUM(D{$blockStartRow}:D" . ($currentRow - 1) . ")",
                    "=SUM(E{$blockStartRow}:E" . ($currentRow - 1) . ")",
                    "=SUM(F{$blockStartRow}:F" . ($currentRow - 1) . ")",
                    "=SUM(G{$blockStartRow}:G" . ($currentRow - 1) . ")",
                    ''
                ];
                Native::applyStyle($handle, $dataSheet, "A{$currentRow}:H{$currentRow}", [
                    'font' => ['bold' => true]
                ]);
                Native::writeRows($handle, $dataSheet, $currentRow, 1, [$subtotalRow]);
                Native::mergeCells($handle, $dataSheet, "A{$currentRow}:C{$currentRow}");

                $currentRow++;
                $blockStartRow = $currentRow;
                $batchStartRow = $currentRow;
                gc_collect_cycles();
            }
        }

        // Flush leftover batch
        if ($batch !== []) {
            Native::writeRows($handle, $dataSheet, $batchStartRow, 1, $batch);
        }

        // Final Subtotal
        if ($currentRow > $blockStartRow) {
            $subtotalRow = [
                'Sub-total',
                '',
                '',
                "=SUM(D{$blockStartRow}:D" . ($currentRow - 1) . ")",
                "=SUM(E{$blockStartRow}:E" . ($currentRow - 1) . ")",
                "=SUM(F{$blockStartRow}:F" . ($currentRow - 1) . ")",
                "=SUM(G{$blockStartRow}:G" . ($currentRow - 1) . ")",
                ''
            ];
            Native::applyStyle($handle, $dataSheet, "A{$currentRow}:H{$currentRow}", [
                'font' => ['bold' => true]
            ]);
            Native::writeRows($handle, $dataSheet, $currentRow, 1, [$subtotalRow]);
            Native::mergeCells($handle, $dataSheet, "A{$currentRow}:C{$currentRow}");
            $currentRow++;
        }

        $lastRow = $currentRow - 1;

        // 9. Write Grand Total Row
        if ($report->getDefinition()->hasTotal()) {
            $grandTotalRow = [
                'Grand Total',
                '',
                '',
                "=SUM(D{$contentStartRow}:D{$lastRow})/2",
                "=SUM(E{$contentStartRow}:E{$lastRow})/2",
                "=SUM(F{$contentStartRow}:F{$lastRow})/2",
                "=SUM(G{$contentStartRow}:G{$lastRow})/2",
                ''
            ];
            Native::applyStyle($handle, $dataSheet, "A{$currentRow}:H{$currentRow}", [
                'font' => ['bold' => true]
            ]);
            Native::writeRows($handle, $dataSheet, $currentRow, 1, [$grandTotalRow]);
            Native::mergeCells($handle, $dataSheet, "A{$currentRow}:C{$currentRow}");
            $currentRow++;
        }

        // 10. Write Signatory block
        $signatory = $report->getReportConfig()->getMetaDataByArrayKey('signatory');
        if (!empty($signatory)) {
            $currentRow += 3;
            Native::writeRows($handle, $dataSheet, $currentRow - 1, 1, [[$report->getUserFullName()]]);
            Native::writeRows($handle, $dataSheet, $currentRow, 1, $signatory);
            $currentRow += count($signatory);
        }

        // 11. Freeze Panes
        $freezeConfig = $metadata['freeze'] ?? 'A10';
        if ($freezeConfig) {
            Native::freezePanes($handle, $dataSheet, $freezeConfig);
        }

        // 12. Apply Row Heights (only to header rows to avoid CGO loop overhead on millions of rows)
        if (isset($metadata['rowHeight'])) {
            for ($r = 1; $r <= 8; $r++) {
                Native::setRowHeight($handle, $dataSheet, $r, 20);
            }
        }

        // 13. Apply Conditional Formatting
        $styles = $report->getReportConfig()->getMetaDataByArrayKey('conditionalStyle');
        if (!empty($styles)) {
            $conditionalRules = [];
            foreach ($styles as $style) {
                $conditionalRules[] = [
                    'type' => $style['conditionType'],
                    'operator' => $style['operatorType'],
                    'conditions' => [$style['condition']],
                    'stopIfTrue' => false,
                    'style' => [
                        'fill' => [
                            'fillType' => 'solid',
                            'startColor' => ['argb' => $style['color']],
                            'endColor' => ['argb' => $style['color']]
                        ],
                        'font' => [
                            'bold' => true
                        ]
                    ]
                ];
            }
            Native::setConditional($handle, $dataSheet, 'A:H', $conditionalRules);
        }

        // 14. Apply Page Background Color (to the written range only, matching compat layer)
        if (isset($metadata['bgColor'])) {
            Native::applyStyle($handle, $dataSheet, "A1:H" . ($currentRow - 1), [
                'fill' => [
                    'fillType' => 'solid',
                    'startColor' => ['argb' => $metadata['bgColor']],
                    'endColor' => ['argb' => $metadata['bgColor']]
                ]
            ]);
        }

        $populateEndTime = microtime(true);
        $populatePeakMemory = memory_get_peak_usage();

        $saveStartTime = microtime(true);

        // 14. Save XLSX workbook
        Native::saveXlsx($handle, $outputPath);

        // 15. Free Native Resource
        Native::close($handle);

        $saveEndTime = microtime(true);
        $savePeakMemory = memory_get_peak_usage();

        return [
            'populate_time' => $populateEndTime - $populateStartTime,
            'populate_memory' => $populatePeakMemory,
            'save_time' => $saveEndTime - $saveStartTime,
            'save_memory' => $savePeakMemory,
        ];
    }
}
