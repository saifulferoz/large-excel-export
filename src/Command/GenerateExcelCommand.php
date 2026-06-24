<?php

namespace App\Command;

use App\Component\Report\Column;
use App\Component\Report\Definition;
use App\Component\Report\ReportConfig;
use App\Component\Report\ReportInterface;
use App\Service\DataGeneratorService;
use App\Service\SpreadsheetGeneratorService;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\Console\Attribute\AsCommand;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Console\Style\SymfonyStyle;

#[AsCommand(
    name: 'app:generate-excel',
    description: 'Generates a memory-optimized 400,000-row Excel spreadsheet.',
)]
class GenerateExcelCommand extends Command
{
    public function __construct(
        private DataGeneratorService $dataGenerator,
        private SpreadsheetGeneratorService $spreadsheetGenerator
    ) {
        parent::__construct();
    }

    protected function configure(): void
    {
        $this
            ->addArgument('rows', InputArgument::OPTIONAL, 'Number of rows to generate', 400000);
    }

    protected function execute(InputInterface $input, OutputInterface $output): int
    {
        $io = new SymfonyStyle($input, $output);
        $io->title('Starting Memory-Optimized Excel Generation');

        // Increase memory limit and execution time for safety (4GB for in-memory generation)
        ini_set('memory_limit', '4G');
        set_time_limit(0);

        $startTime = microtime(true);
        $startMemory = memory_get_usage();

        // 1. Create Mock Report Implementation matching all requirements
        $report = new class() implements ReportInterface {
            private Definition $definition;
            private ReportConfig $config;

            public function __construct()
            {
                $columns = [
                    new Column('id', 'Product ID', Column::TYPE_INTEGER, [
                        'width' => 12,
                    ]),
                    new Column('name', 'Product Name', Column::TYPE_STRING, [
                        'width' => 20,
                    ]),
                    new Column('category', 'Category', Column::TYPE_STRING, [
                        'width' => 18,
                        'wrapText' => true,
                    ]),
                    new Column('quantity', 'Quantity', Column::TYPE_INTEGER, [
                        'width' => 12,
                        'total' => true,
                        'subTotal' => true,
                    ]),
                    new Column('price', 'Unit Price', Column::TYPE_FLOAT, [
                        'width' => 15,
                        'total' => true,
                        'subTotal' => true,
                        'format' => '$#,##0.00',
                    ]),
                    new Column('subtotal', 'Subtotal (Formula)', Column::TYPE_FLOAT, [
                        'width' => 18,
                        'total' => true,
                        'subTotal' => true,
                        'format' => '$#,##0.00',
                    ]),
                    new Column('tax', 'Tax (Formula)', Column::TYPE_FLOAT, [
                        'width' => 15,
                        'total' => true,
                        'subTotal' => true,
                        'format' => '$#,##0.00',
                    ]),
                    new Column('date', 'Creation Date', Column::TYPE_DATE, [
                        'width' => 15,
                    ]),
                ];

                $reportHeader = [
                    [
                        ['label' => 'Product Sales Report - 400k Rows', 'colspan' => 8],
                    ],
                ];

                $this->definition = new Definition('Sales Activity Report', $columns, true, $reportHeader);

                $this->config = new ReportConfig([
                    'tableHeaderBg' => 'FF1F4E78', // Dark blue table header
                    'signatory' => [
                        ['Prepared By', 'Checked By', 'Approved By'],
                        ['System Exporter', 'Auditor Manager', 'Executive Officer'],
                    ],
                    'conditionalStyle' => [
                        [
                            'conditionType' => 'expression',
                            'operatorType' => 'equal',
                            'condition' => '$D10>90', // Highlight quantity > 90
                            'color' => 'FFE2EFDA', // Soft green background
                        ],
                    ],
                ]);
            }

            public function getColumns(): array
            {
                // Returns all columns as raw object array
                return [
                    $this->definition->getColumn('id'),
                    $this->definition->getColumn('name'),
                    $this->definition->getColumn('category'),
                    $this->definition->getColumn('quantity'),
                    $this->definition->getColumn('price'),
                    $this->definition->getColumn('subtotal'),
                    $this->definition->getColumn('tax'),
                    $this->definition->getColumn('date'),
                ];
            }

            public function getColumnLabels(): array
            {
                return [
                    'id' => 'Product ID',
                    'name' => 'Product Name',
                    'category' => 'Category',
                    'quantity' => 'Quantity',
                    'price' => 'Unit Price',
                    'subtotal' => 'Subtotal (Formula)',
                    'tax' => 'Tax (Formula)',
                    'date' => 'Creation Date',
                ];
            }

            public function getDefinition(): Definition
            {
                return $this->definition;
            }

            public function getHeaders(): array
            {
                return [
                    ['Report Parameters: Period = Q2 2026 | Scope = Global'],
                ];
            }

            public function getMetadata(): array
            {
                return [
                    'style' => ['border' => 'thin'],
                    'rowGroup' => ['serial' => false],
                    'skipSubTotalColumnMerge' => false,
                    'rowHeight' => 20,
                    'freeze' => 'A10', // Freeze table headers and above
                    'bgColor' => 'FFF2F2F2', // Soft background
                ];
            }

            public function getMergedRow(): array
            {
                // Demonstration of row merging (handled in SpreadsheetGeneratorService)
                return [];
            }

            public function getSubTotalRow(): array
            {
                // Periodical sub-totals are handled directly inside generators loop
                return [];
            }

            public function getReportConfig(): ReportConfig
            {
                return $this->config;
            }

            public function getUserFullName(): string
            {
                return 'Automated System Service';
            }

            public function getColumnHeaderTextRotation(): ?int
            {
                return 0;
            }

            public function getAllRows()
            {
                return []; // Unused since we stream via generator
            }
        };

        // 2. Setup output directories
        $outputDir = dirname(__DIR__, 2).'/public/reports';
        if (!is_dir($outputDir)) {
            mkdir($outputDir, 0777, true);
        }
        $outputPath = $outputDir.'/large_report.xlsx';

        $io->text('Initializing spreadsheet...');
        $spreadsheet = new Spreadsheet();

        // 3. Generate data stream generator
        $rowCount = (int)$input->getArgument('rows');
        if ($rowCount <= 0) {
            $rowCount = 400000;
        }
        $io->text(sprintf('Streaming %s data rows into sheet...', number_format($rowCount)));
        $generator = $this->dataGenerator->generate($rowCount);

        $populateStartTime = microtime(true);
        $populateStartMemory = memory_get_usage();

        // 4. Generate worksheet content
        $this->spreadsheetGenerator->writeData($report, $spreadsheet, $generator);

        $populateEndTime = microtime(true);
        $populatePeakMemory = memory_get_peak_usage();

        $saveStartTime = microtime(true);
        $saveStartMemory = memory_get_usage();

        // 5. Save Spreadsheet
        $io->text('Writing output to Xlsx file: '.$outputPath);
        $writer = new Xlsx($spreadsheet);

        // Key Performance Settings: Disable formula calculation on write to save time and memory
        $writer->setPreCalculateFormulas(false);
        $writer->save($outputPath);

        $saveEndTime = microtime(true);
        $savePeakMemory = memory_get_peak_usage();

        // 6. Output Statistics
        $endTime = microtime(true);
        $peakMemory = memory_get_peak_usage(true);

        $totalDuration = $endTime - $startTime;
        $populateDuration = $populateEndTime - $populateStartTime;
        $saveDuration = $saveEndTime - $saveStartTime;

        $io->success([
            'Excel file generated successfully!',
            sprintf('File Path: %s', $outputPath),
            sprintf('File Size: %s MB', number_format(filesize($outputPath) / 1024 / 1024, 2)),
            '',
            '--- BENCHMARKING DETAILS ---',
            '1. POPULATE PHASE (Streaming data rows & formatting):',
            sprintf('   - Time taken: %s seconds', number_format($populateDuration, 2)),
            sprintf('   - Peak memory at end of phase: %s MB', number_format($populatePeakMemory / 1024 / 1024, 2)),
            '',
            '2. WRITE/SAVE PHASE (Serializing and saving .xlsx file to disk):',
            sprintf('   - Time taken: %s seconds', number_format($saveDuration, 2)),
            sprintf('   - Peak memory at end of phase: %s MB', number_format($savePeakMemory / 1024 / 1024, 2)),
            '',
            '3. OVERALL METRICS:',
            sprintf('   - Total Duration: %s seconds', number_format($totalDuration, 2)),
            sprintf('   - Peak Memory Allocated: %s MB', number_format($peakMemory / 1024 / 1024, 2)),
        ]);

        return Command::SUCCESS;
    }
}
