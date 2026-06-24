<?php

namespace App\Component\Report\PhpOffice;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Cell\DefaultValueBinder;

class StringValueBinder extends DefaultValueBinder
{
    public const EXCEL_MAX_INT = 9007199254740992; // 2^53, max safe integer in JavaScript

    public function bindValue(Cell $cell, $value): bool
    {
        if (is_numeric($value) && !str_contains($value, '.') && $value > self::EXCEL_MAX_INT) {
            $cell->setValueExplicit((string) $value, DataType::TYPE_STRING);

            return true;
        }

        return parent::bindValue($cell, $value);
    }
}
