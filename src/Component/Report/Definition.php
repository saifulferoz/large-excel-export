<?php

namespace App\Component\Report;

class Definition
{
    private string $title;
    private array $columns = [];
    private bool $hasTotal = false;
    private array $reportHeader = [];

    public function __construct(string $title, array $columns, bool $hasTotal, array $reportHeader)
    {
        $this->title = $title;
        foreach ($columns as $column) {
            $this->columns[$column->getName()] = $column;
        }
        $this->hasTotal = $hasTotal;
        $this->reportHeader = $reportHeader;
    }

    public function getTitle(): string
    {
        return $this->title;
    }

    public function hasTotal(): bool
    {
        return $this->hasTotal;
    }

    public function getColumn(string $name): Column
    {
        if (!isset($this->columns[$name])) {
            throw new \InvalidArgumentException("Column '$name' not found in definition");
        }
        return $this->columns[$name];
    }

    public function getReportHeader(): array
    {
        return $this->reportHeader;
    }
}
