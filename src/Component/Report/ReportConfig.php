<?php

namespace App\Component\Report;

class ReportConfig
{
    private array $metadata = [];

    public function __construct(array $metadata = [])
    {
        $this->metadata = $metadata;
    }

    public function getMetaDataByArrayKey(string $key): mixed
    {
        return $this->metadata[$key] ?? null;
    }
}
