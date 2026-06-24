<?php

namespace App\Utility;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

abstract class SpreadsheetUtil
{
    public const ACCOUNTING_FORMAT = '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)';
    public const ACCOUNTING_FORMAT_WITH_ZERO = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';

    public static function writeReportDate(Worksheet $worksheet)
    {
        $worksheet->setCellValue('A4', 'Report Date: ' . date('d F Y'));
    }

    public static function applyDefaultTableHeaderStyle(Worksheet $worksheet, int $row, int $additionalColumns = 0)
    {
        $lastColumn = $worksheet->getHighestColumn($row);
        $lastColumnIndex = self::getIndexByColumn($lastColumn) + $additionalColumns;
        for ($i = 1; $i <= $lastColumnIndex; ++$i) {
            $cell = self::getCellByColumnIndex($i, $row);
            $worksheet->getStyle($cell)->applyFromArray(self::getTableHeaderStyle());
        }
    }

    public static function applyBoldCenter(Worksheet $worksheet, string $cell)
    {
        $worksheet->getStyle($cell)->applyFromArray(self::getTableHeaderStyle());
    }

    public static function applyContentStyle(Worksheet $worksheet, int $startRow)
    {
        $highestRow = $worksheet->getHighestRow('A');
        $highestColumn = $worksheet->getHighestColumn($startRow);
        $highestIndex = self::getIndexByColumn($highestColumn);
        for ($i = 1; $i <= $highestIndex; ++$i) {
            for ($j = $startRow; $j <= $highestRow; ++$j) {
                $cell = self::getCellByColumnIndex($i, $j);
                $worksheet->getStyle($cell)->applyFromArray(self::getPageDefaultStyle());
            }
        }
    }

    public static function applyBoldRowStyle(Worksheet $worksheet, int $startRow)
    {
        $style = self::getPageDefaultStyle();
        $style['font']['bold'] = true;

        $highestColumn = $worksheet->getHighestColumn($startRow);
        $highestIndex = self::getIndexByColumn($highestColumn);
        for ($i = 1; $i <= $highestIndex; ++$i) {
            $cell = self::getCellByColumnIndex($i, $startRow);
            $worksheet->getStyle($cell)->applyFromArray($style);
        }
    }

    public static function applyNumberAlignment(Worksheet $worksheet, string $cell)
    {
        $worksheet->getStyle($cell)->applyFromArray(self::getAmountAlignment());
    }

    public static function setFontBold(Worksheet $worksheet, string $cell)
    {
        $fontBoldArray = [
            'font' => [
                'bold' => true,
            ],
        ];
        $worksheet->getStyle($cell)->applyFromArray($fontBoldArray);
    }

    public static function setFontBoldWithRightAlignment(Worksheet $worksheet, string $cell)
    {
        $fontBoldArray = [
            'font' => [
                'bold' => true,
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_RIGHT,
                'vertical' => Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
        ];
        $worksheet->getStyle($cell)->applyFromArray($fontBoldArray);
    }

    public static function setFontBoldWithRightAlignmentForNumber(Worksheet $worksheet, string $cell)
    {
        $fontBoldArray = [
            'font' => [
                'bold' => true,
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_RIGHT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ];
        $worksheet->getStyle($cell)->applyFromArray($fontBoldArray);
    }

    public static function applyCellFormat(Worksheet $worksheet, string $cell, ?string $format = null)
    {
        $worksheet
            ->getStyle($cell)
            ->getNumberFormat()
            ->setFormatCode($format);
    }

    public static function getTableHeaderStyle()
    {
        return [
            'font' => [
                'name' => 'Times New Roman',
                'bold' => true,
                'italic' => false,
                'underline' => false,
                'strikethrough' => false,
                'color' => [
                    'rgb' => '000000',
                ],
            ],
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => [
                        'rgb' => '000000',
                    ],
                ],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
        ];
    }

    public static function getPageDefaultStyle()
    {
        return [
            'font' => [
                'name' => 'Times New Roman',
                'bold' => false,
                'italic' => false,
                'underline' => false,
                'strikethrough' => false,
                'color' => [
                    'rgb' => '000000',
                ],
            ],
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => [
                        'rgb' => '000000',
                    ],
                ],
            ],
        ];
    }

    public static function getAmountAlignment()
    {
        return [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_RIGHT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ];
    }

    public static function getIndexByColumn(string $column): int
    {
        return \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($column);
    }

    public static function getCellByColumnIndex(int $columnIndex, int $row)
    {
        $column = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($columnIndex);

        return $column . $row;
    }

    public static function applyTopBorder(Worksheet $worksheet, $startColumn, $endColumn, $rowNumber)
    {
        $styleArray = [
            'borders' => [
                'top' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => [
                        'rgb' => '000000',
                    ],
                ],
            ],
        ];
        $worksheet->getStyle($startColumn . $rowNumber . ':' . $endColumn . $rowNumber)->applyFromArray($styleArray);
    }

    public static function applyRangeBorder(Worksheet $worksheet, string $cells, string $position)
    {
        $styleArray = [
            'borders' => [
                $position => [
                    'borderStyle' => Border::BORDER_MEDIUM,
                    'color' => [
                        'rgb' => '000000',
                    ],
                ],
            ],
        ];
        $worksheet->getStyle($cells)->applyFromArray($styleArray);
    }

    public static function applyBorder(Worksheet $worksheet, string $cell)
    {
        $styleArray = [
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => [
                        'rgb' => '000000',
                    ],
                ],
            ],
        ];
        $worksheet->getStyle($cell)->applyFromArray($styleArray);
    }

    public static function getColumnByIndex(int $columnIndex)
    {
        return \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($columnIndex);
    }

    public static function getNextColumn(string $currentColumn)
    {
        $currentIndex = self::getIndexByColumn($currentColumn);

        return self::getColumnByIndex($currentIndex + 1);
    }

    public static function setColumnWidth(Worksheet $worksheet, string $column, int $width)
    {
        $worksheet->getColumnDimension($column)->setWidth($width);
    }

    public static function setColumnFitToText(Worksheet $worksheet, string $column)
    {
        $worksheet->getColumnDimension($column)->setAutoSize(true);
    }

    public static function setFreezePane(Worksheet $worksheet, string $cell)
    {
        $worksheet->freezePane($cell);
    }

    public static function setFontBoldWithAlignment(
        Worksheet $worksheet,
        string $cell,
        string $alignment = Alignment::HORIZONTAL_LEFT
    ) {
        $fontBoldArray = [
            'font' => [
                'bold' => true,
            ],
            'alignment' => [
                'horizontal' => $alignment,
                'vertical' => Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
        ];
        $worksheet->getStyle($cell)->applyFromArray($fontBoldArray);
    }

    public static function setPrintArea(Worksheet $worksheet, string $cellRange)
    {
        $worksheet->getPageSetup()->setPrintArea($cellRange);
        $worksheet->getPageSetup()->setOrientation(
            \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE
        );
    }

    public static function applyLeftAlignment(Worksheet $worksheet, string $cell)
    {
        $styleConfig = [
            'horizontal' => Alignment::HORIZONTAL_LEFT,
            'vertical' => Alignment::VERTICAL_CENTER,
            'wrapText' => true,
        ];

        $worksheet->getStyle($cell)->getAlignment()->applyFromArray($styleConfig);
    }

    public static function applyRightAlignment(Worksheet $worksheet, string $cell)
    {
        $styleConfig = [
            'horizontal' => Alignment::HORIZONTAL_RIGHT,
            'vertical' => Alignment::VERTICAL_CENTER,
            'wrapText' => true,
        ];

        $worksheet->getStyle($cell)->getAlignment()->applyFromArray($styleConfig);
    }

    public static function applyLeftIndent(Worksheet $worksheet, string $cell, int $pValue)
    {
        $worksheet->getStyle($cell)->getAlignment()->setHorizontal('left')->setIndent($pValue);
    }

    public static function setAutoSize(Worksheet $worksheet, string $column)
    {
        $worksheet->getColumnDimension($column)->setAutoSize(true);
    }

    public static function getPreviousColumn(string $currentColumn)
    {
        $currentIndex = self::getIndexByColumn($currentColumn);

        return self::getColumnByIndex($currentIndex - 1);
    }

    public static function setBackgroundColor(Worksheet $worksheet, string $cells, string $color = 'ffffffff')
    {
        $worksheet->getStyle($cells)->getFill()->setFillType(
            \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID
        )->getStartColor()->setARGB($color);
    }

    public static function applyRightIndent(Worksheet $worksheet, string $cell, int $value)
    {
        $worksheet->getStyle($cell)->getAlignment()->setHorizontal('right')->setIndent($value);
    }

    public static function applyPartialBorder(Worksheet $worksheet, string $cell, string $borderSide)
    {
        $styleArray = [
            'borders' => [
                $borderSide => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => [
                        'rgb' => '000000',
                    ],
                ],
            ],
        ];
        $worksheet->getStyle($cell)->applyFromArray($styleArray);
    }

    public static function getColumnList(string $lastColumn, ?string $fromColumnName = null)
    {
        $fromColumn = $fromColumnName ?? 'A';
        $startColumnIndex = self::getIndexByColumn($fromColumn);
        $lastColumnIndex = self::getIndexByColumn($lastColumn);
        $columns = [];
        while ($startColumnIndex <= $lastColumnIndex) {
            $columns[] = self::getColumnByIndex($startColumnIndex);
            ++$startColumnIndex;
        }

        return $columns;
    }

    public static function getCellActualValueAndValidate(Worksheet $worksheet, int $columnIndex, int $rowNo, $callBack)
    {
        $cellRef = self::getColumnByIndex($columnIndex) . $rowNo;

        $value = $worksheet->getCell($cellRef)->getValue();
        call_user_func_array($callBack, [$cellRef, $value]);

        return ['cell' => $cellRef, 'value' => $value];
    }

    public static function applyAlignment(
        Worksheet $worksheet,
        string $cell,
        $horizontalAlign = Alignment::HORIZONTAL_CENTER,
        $verticalAlign = Alignment::VERTICAL_CENTER
    ) {
        $styleConfig = [
            'horizontal' => $horizontalAlign,
            'vertical' => $verticalAlign,
        ];

        $worksheet->getStyle($cell)->getAlignment()->applyFromArray($styleConfig);
    }

    public static function getTargetColumn(string $column, int $targetIndex): string
    {
        $currentIndex = self::getIndexByColumn($column);

        return self::getColumnByIndex($currentIndex + $targetIndex);
    }
}
