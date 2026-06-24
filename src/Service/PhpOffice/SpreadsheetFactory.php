<?php

/*
 * This file is part of the sbiCloud Addons module.
 *
 * Copyright (c) 2019-2022, BRAC IT SERVICES LIMITED <https://www.bracits.com>
 */

namespace App\Service\PhpOffice;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use Symfony\Component\HttpFoundation\StreamedResponse;

class SpreadsheetFactory
{
    public function __construct()
    {
        IOFactory::registerWriter('Html', HTMLWriter::class);
        IOFactory::registerWriter('Pdf', PDFWriter::class);
    }

    /**
     * Returns a new instance of the PhpSpreadsheet class.
     *
     * @param string|null $filename if set, uses the IOFactory to return the spreadsheet located at $filename
     *                              using automatic type resolution per \PhpOffice\PhpSpreadsheet\IOFactory
     *
     * @return Spreadsheet
     *
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function createSpreadsheet($filename = null)
    {
        return null === $filename ? new Spreadsheet() : IOFactory::load($filename);
    }


    /**
     * Returns the PhpSpreadsheet IWriter instance to save a file.
     *
     * @return IWriter
     *
     * @throws Exception
     */
    public function createWriter(Spreadsheet $spreadsheet, $type)
    {
        $writer = IOFactory::createWriter($spreadsheet, $type);

        if ('Pdf' === $type) {
            $this->applyOptionsToWriter($writer, ['setInternalWriter' => $this->pdfWriter]);
        }

        return $writer;
    }

    /**
     * @param string $type reader class to create
     *
     * @return mixed|IReader returns a IReader of the given type if found
     *
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function createReader($type)
    {
        return IOFactory::createReader($type);
    }

    /**
     * Return a StreamedResponse containing the file.
     *
     * @param string $type
     * @param int $status
     * @param array $headers
     * @param array $writerOptions
     *
     * @return StreamedResponse
     *
     * @throws Exception
     */
    public function createStreamedResponse(
        Spreadsheet $spreadsheet,
        $type,
        $status = 200,
        $headers = [],
        $writerOptions = []
    ) {
        $writer = IOFactory::createWriter($spreadsheet, $type);
        if (!empty($writerOptions)) {
            $this->applyOptionsToWriter($writer, $writerOptions);
        }

        return new StreamedResponse(
            function () use ($writer): void {
                $writer->save('php://output');
            },
            $status,
            $headers
        );
    }

    private function applyOptionsToWriter(IWriter $writer, array $options = []): IWriter
    {
        foreach ($options as $method => $arguments) {
            if (method_exists($writer, $method)) {
                if (!is_array($arguments)) {
                    $arguments = [$arguments];
                }
                call_user_func_array([$writer, $method], $arguments);
            }
        }

        return $writer;
    }
}
