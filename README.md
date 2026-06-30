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

Run the Symfony console commands to generate the Excel report using the different drivers:

### 1. easy-excel (Compatibility Layer)
Uses the PhpSpreadsheet API facade mapped to the Go C-bindings under the hood:
```bash
php bin/console app:generate-excel [rows]
```

### 2. easy-excel (Native API)
Uses the raw flat static bindings of `EasyExcel\Native` to communicate directly with Go:
```bash
php bin/console app:generate-excel-native [rows]
```

### 3. OpenSpout
Uses the fast, lightweight OpenSpout streaming writer:
```bash
php bin/console app:generate-excel-openspout [rows]
```

---

## 📊 Benchmarking & Performance Metrics

Below is a detailed comparison of execution time, PHP peak memory usage, and file sizes across the three implementations, tested on a dataset of **10,000**, **100,000**, and **1,000,000** rows.

### Performance Table

| Driver | Row Count | Populate Phase (s) | Save Phase (s) | Total Time (s) | PHP Peak Memory | File Size |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **easy-excel (Compat)** | 10,000 | 3.26 | 0.19 | 3.46 | 10.00 MB | 0.49 MB |
| **easy-excel (Native)** | 10,000 | 0.27 | 1.40 | 1.70 | 10.00 MB | 0.47 MB |
| **OpenSpout** | 10,000 | 0.12 | 0.19 | 0.31 | 10.00 MB | 0.45 MB |
| | | | | | | |
| **easy-excel (Compat)** | 100,000 | 11.15 | 7.84 | 18.99 | 12.00 MB | 4.71 MB |
| **easy-excel (Native)** | 100,000 | 1.26 | 7.82 | 9.08 | 10.00 MB | 4.51 MB |
| **OpenSpout** | 100,000 | 1.29 | 1.71 | 3.00 | 10.00 MB | 4.46 MB |
| | | | | | | |
| **easy-excel (Compat)** | 1,000,000 | 54.10 | 30.72 | 84.83 | 12.00 MB | 46.87 MB |
| **easy-excel (Native)** | 1,000,000 | 9.18 | 102.06 | 111.26 | 10.00 MB | 44.85 MB |
| **OpenSpout** | 1,000,000 | 12.60 | 21.11 | 33.71 | 10.00 MB | 44.57 MB |

---

## 🔍 Key Insights & Architectural Trade-offs

### 1. Memory Consumption
- All three drivers maintain a **constant-memory footprint** (~10.00 MB to 12.00 MB of PHP memory) regardless of whether they write 10,000 or 1,000,000 rows.
- This is achieved via PHP Generators (`yield`) which stream rows on the fly, preventing data rows from accumulating in PHP memory.

### 2. easy-excel: Compat vs. Native
- **Populate Speed**: The native `easy-excel` driver populates data **6x faster** (9.18s vs 54.10s for 1M rows) than the compatibility layer because it writes rows in flat 2D arrays directly across the CGO boundary rather than resolving cells one-by-one in PHP objects.
- **Save Speed**: The compatibility layer saves faster for 1M rows because it streams cell definitions into Go in real-time. To replicate this in the native driver, all styling operations (fonts, alignments, borders) must be declared **before** the rows are written so that the Go-excelize StreamWriter can inline them sequentially rather than buffering sheet cells in memory.

### 3. OpenSpout
- **Speed**: OpenSpout is the fastest writer (33.71s for 1,000,000 rows) because it writes raw XML directly without the overhead of formula compilation, cell coordinate validations, or complex styling models.
- **Feature Gap**: OpenSpout does not support formulas, cross-sheet references, conditional formatting, page freeze, or complex corporate page headers. For reports needing professional formatting, `easy-excel` provides a near-identical feature set to PhpSpreadsheet at a fraction of the memory cost.
