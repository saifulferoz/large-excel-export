<?php

/*
 * This file is part of the sbiCloud Addons module.
 *
 * Copyright (c) 2019-2022, BRAC IT SERVICES LIMITED <https://www.bracits.com>
 */

namespace App\Service\PhpOffice;

class Spreadsheet
{
    private $spreadsheet;
    private $worksheet;

    public function __construct(SpreadsheetFactory $factory)
    {
        $this->spreadsheet = $factory->createSpreadsheet();
        $this->spreadsheet->addSheet(0);
    }

    public function getActiveSheet()
    {
        $this->worksheet = $this->spreadsheet->getSheet(0);
    }

    public function createHeader(array $reportHeader, string $reportTitle)
    {
        $this->worksheet->mergeCells('A1:B1');
        $this->worksheet->mergeCells('A1:A3');
        $this->worksheet->mergeCells('B1:B3');
        $numberOfReportColumn = count($reportHeader);
        $this->worksheet->mergeCells('C1:' . $this->numberToAlpha($numberOfReportColumn - 2));
        $this->createReportTitle($reportTitle);
        $this->worksheet->mergeCells('A5:B5');
        $this->worksheet->mergeCells('C5:' . $this->numberToAlpha($numberOfReportColumn - 2));
    }

    public function createReportTitle(string $reportTitle)
    {
        $this->worksheet->setCellValue('C5', $reportTitle);
    }

    public function createReportDate()
    {
        $this->worksheet->setCellValue('A5', 'Report Date: ' . date('d F Y'));
    }

    public function createRowHeader(array $rowHeader)
    {
        foreach ($rowHeader as $index => $row) {
            $this->worksheet->setCellValue($this->numberToAlpha($index + 1) . '7', $row);
        }
    }

    public function writeArrayDataToCell(array $rows)
    {
        $excelRowIndex = 8;
        foreach ($rows as $row) {
            foreach ($row as $index => $item) {
                $this->worksheet->setCellValue($this->numberToAlpha($index + 1) . $excelRowIndex, $item);
            }
            ++$excelRowIndex;
        }
    }

    public function numberToAlpha($data)
    {
        $alphabet = range('a', 'z');
        if ($data <= 25) {
            return $alphabet[$data];
        } elseif ($data > 25) {
            $dividend = $data + 1;
            $alpha = '';
            while ($dividend > 0) {
                $modulo = ($dividend - 1) % 26;
                $alpha = $alphabet[$modulo] . $alpha;
                $dividend = floor((($dividend - $modulo) / 26));
            }

            return $alpha;
        }
        return null;
    }

    public function alphaToNum($data)
    {
        $alphabet = range('a', 'z');
        $alpha_flip = array_flip($alphabet);
        $return_value = -1;
        $length = strlen((string) $data);
        for ($i = 0; $i < $length; ++$i) {
            $return_value +=
                ($alpha_flip[$data[$i]] + 1) * 26 ** ($length - $i - 1);
        }

        return $return_value;
    }
}
