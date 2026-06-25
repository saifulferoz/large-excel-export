# Symfony Large Excel Exporter

A high-performance, memory-optimized Symfony console application designed to generate extremely large Excel spreadsheets (e.g., 400,000+ rows) utilizing the `phpoffice/phpspreadsheet` library. 

This repository serves as a reference implementation for generating large-scale data reports with complex styling, dynamic sub-totals, cross-sheet references, conditional formatting, and image rendering, all while maintaining a low memory footprint.

---

## 🚀 Key Features & Performance Optimizations

1. **Memory-Optimized Streaming**
   - Implements PHP Generators (`yield`) via [DataGeneratorService](file:///Users/saifulislamferoz/Projects/large-excel-export/src/Service/DataGeneratorService.php) to stream data row-by-row rather than buffer all records in memory.
   
2. **Column-Level Styling Optimization**
   - Applies borders, alignments, and number formats directly to entire columns (e.g., `A:H`) instead of individual cells or large cell ranges. This reduces the number of style objects tracked by PhpSpreadsheet in memory, avoiding Out of Memory (OOM) errors.

3. **Cross-Sheet Formula References**
   - Creates a primary data sheet (`DataSheet`) and a secondary configuration/reference sheet (`ConfigSheet`).
   - Formulas in the data sheet reference the tax rate defined dynamically on `ConfigSheet` (e.g., `=F10*ConfigSheet!$B$1`).

4. **Dynamic Subtotals & Grand Totals**
   - Periodically injects custom subtotal rows (every 50,000 rows by default) to aggregate section metrics.
   - Calculates grand totals using formulas that partition the columns (e.g., `=SUM(D10:D400000)/2` to account for the interspersed subtotals).

5. **Conditional Formatting**
   - Applies conditional rules at the column scale to highlight records dynamically (e.g., highlighting quantities $> 90$ with a soft green background).

6. **Image & Header Placement**
   - Generates a lightweight, custom image dynamically and places it as the corporate logo in the sheet header using PhpSpreadsheet drawing tools.

7. **Garbage Collection & Pre-calculation Offloading**
   - Triggers explicit PHP garbage collection cycles (`gc_collect_cycles()`) every 10,000 rows to reclaim unreferenced cell memory.
   - Disables formula pre-calculation on writer save (`$writer->setPreCalculateFormulas(false)`) to delegate calculation weight to the target spreadsheet application (like Microsoft Excel or Google Sheets).

---

## 📁 Project Architecture

The core codebase is organized as follows:

- **[GenerateExcelCommand](file:///Users/saifulislamferoz/Projects/large-excel-export/src/Command/GenerateExcelCommand.php)**: The Symfony command executing the generation workflow, handling mock report specifications, benchmarks, and writing to the public reports directory.
- **[SpreadsheetGeneratorService](file:///Users/saifulislamferoz/Projects/large-excel-export/src/Service/SpreadsheetGeneratorService.php)**: The orchestration service that translates the report interface configurations, metadata, logos, columns, and data streams into the final styled Excel structure.
- **[DataGeneratorService](file:///Users/saifulislamferoz/Projects/large-excel-export/src/Service/DataGeneratorService.php)**: Generator-based service producing millions of rows on the fly with low CPU/memory utilization.
- **[Component/Report](file:///Users/saifulislamferoz/Projects/large-excel-export/src/Component/Report)**: OOP structure (`Column`, `Definition`, `ReportConfig`, `ReportInterface`) defining report layout, styling parameters, and conditional rule objects.
- **[SpreadsheetUtil](file:///Users/saifulislamferoz/Projects/large-excel-export/src/Utility/SpreadsheetUtil.php)**: Helper class containing formatting constants, alignments, border helpers, and coordinate conversion utilities.

---

## 🛠️ Installation & Setup

### Requirements
- **PHP**: `>= 8.4`
- **Extensions**: `ext-ctype`, `ext-iconv`, `ext-gd` (required for image/logo generation)
- **Composer**

### Setup Steps
1. Clone the repository and navigate to the project directory:
   ```bash
   composer install
   ```

2. Generate the directory for report output if needed (handled automatically by command):
   ```bash
   mkdir -p public/reports
   ```

---

## 💻 Usage

Run the Symfony console command to generate the Excel report:

```bash
php bin/console app:generate-excel [rows]
```

### Parameters
- `rows` *(Optional, Default: `400000`)*: The number of data rows to stream into the spreadsheet.

### Examples
*Generate a quick test sheet with 100 rows:*
```bash
php bin/console app:generate-excel 100
```

*Generate the full-scale 400,000 row production test sheet:*
```bash
php bin/console app:generate-excel 400000
```

The output file will be saved at:
📁 `public/reports/large_report.xlsx`

---

## 📊 Benchmarking & Performance Metrics

Running a test generation of **100 rows** yields the following performance telemetry:

```text
Starting Memory-Optimized Excel Generation
==========================================

 Initializing spreadsheet...
 Streaming 100 data rows into sheet...
 Writing output to Xlsx file: public/reports/large_report.xlsx

 [OK] Excel file generated successfully!                                        
                                                                                
      File Path: public/reports/large_report.xlsx                                                                
      File Size: 5.11 MB                                                        
                                                                                
      --- BENCHMARKING DETAILS ---                                              
                                                                                
      1. POPULATE PHASE (Streaming data rows & formatting):                     
         - Time taken: ~37 seconds                                            
         - Peak memory: ~638 MB                               
                                                                                
      2. WRITE/SAVE PHASE (Serializing and saving .xlsx file to disk):          
         - Time taken: ~21 seconds                                            
         - Peak memory: ~778 MB                               
                                                                                
      3. OVERALL METRICS:                                                       
         - Total Duration: ~58 seconds                                        
         - Peak Memory Allocated: ~816 MB                                     
```

> [!NOTE]
> Performance will vary based on hardware, CPU capabilities, and local PHP configurations.
