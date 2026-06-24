<?php

namespace App\Component\Report;

interface ReportInterface
{
    public function getColumns(): array;
    public function getColumnLabels(): array;
    public function getDefinition(): Definition;
    public function getHeaders(): array;
    public function getMetadata(): array;
    public function getMergedRow(): array;
    public function getSubTotalRow(): array;
    public function getReportConfig(): ReportConfig;
    public function getUserFullName(): string;
    public function getColumnHeaderTextRotation(): ?int;
    public function getAllRows();
}
