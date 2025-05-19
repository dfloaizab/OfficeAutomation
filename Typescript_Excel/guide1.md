# TypeScript for Excel: Comprehensive Reference Guide
# 2024-2025, Diego F. Loaiza

## Table of Contents
1. [Introduction](#introduction)
2. [Setup and Environment](#setup-and-environment)
3. [Core Concepts](#core-concepts)
4. [Office Scripts API](#office-scripts-api)
5. [Excel API Reference](#excel-api-reference)
6. [Practical Patterns](#practical-patterns)
7. [Debugging and Testing](#debugging-and-testing)
8. [Exercises](#exercises)

## Introduction

TypeScript for Excel enables you to automate Excel tasks, create custom functions, and build add-ins using type-safe code. This reference guide covers everything you need to know to build powerful Excel solutions using TypeScript.

### What is TypeScript for Excel?

TypeScript for Excel encompasses two main development approaches:
- **Office Scripts**: Lightweight automation scripts that run directly in Excel
- **Office Add-ins**: More complex solutions with custom UI and deeper integration

Both use TypeScript to provide type safety, intellisense, and modern JavaScript features while interacting with Excel's object model.

## Setup and Environment

### Office Scripts Setup

1. **Requirements**: Microsoft 365 subscription with Excel on the web
2. **Access**: Open Excel on the web, go to the "Automate" tab
3. **Code Editor**: Use the built-in Code Editor or VSCode with appropriate extensions

```typescript
// Basic Office Script structure
function main(workbook: ExcelScript.Workbook) {
  // Your code here
  let sheet = workbook.getActiveWorksheet();
  sheet.getRange("A1").setValue("Hello from TypeScript!");
}
```

### Add-in Development Setup

1. **Install Node.js**: Download and install from nodejs.org
2. **Install Yeoman and Office generator**:
   ```bash
   npm install -g yo generator-office
   ```
3. **Create new project**:
   ```bash
   yo office
   ```
4. **Project structure**:
   - `/src`: Source code
   - `/assets`: Images and resources
   - `manifest.xml`: Add-in definition
   - `tsconfig.json`: TypeScript configuration

## Core Concepts

### TypeScript Fundamentals for Excel

#### Type Declaration

```typescript
// Basic types
let cellValue: string = "Hello";
let rowCount: number = 10;
let isSelected: boolean = true;

// Excel-specific types
let range: ExcelScript.Range;
let sheet: ExcelScript.Worksheet;
let table: ExcelScript.Table;
```

#### Interfaces and Types

```typescript
// Define custom data structures
interface SalesRecord {
  date: Date;
  product: string;
  quantity: number;
  revenue: number;
}

// Use with Excel data
function processSalesData(workbook: ExcelScript.Workbook): SalesRecord[] {
  const sheet = workbook.getWorksheet("Sales");
  const dataRange = sheet.getUsedRange();
  const values = dataRange.getValues();
  
  const records: SalesRecord[] = [];
  // Transform Excel data to typed objects
  for (let i = 1; i < values.length; i++) {
    records.push({
      date: new Date(values[i][0]),
      product: values[i][1].toString(),
      quantity: values[i][2] as number,
      revenue: values[i][3] as number
    });
  }
  
  return records;
}
```

#### Async Programming

```typescript
// For Office Add-ins (not Office Scripts)
async function loadWorkbookData(): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    
    await context.sync();
    
    console.log(range.values);
  });
}
```

## Office Scripts API

### Workbook Operations

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get worksheets
  let sheet = workbook.getActiveWorksheet();
  let namedSheet = workbook.getWorksheet("Sales");
  
  // Create a new worksheet
  let newSheet = workbook.addWorksheet("Analysis");
  
  // Get all worksheets
  let sheets = workbook.getWorksheets();
  
  // Get named ranges
  let namedRange = workbook.getNamedItem("SalesRegion");
}
```

### Worksheet Operations

```typescript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  
  // Get range by address
  let range = sheet.getRange("A1:D10");
  
  // Get used range
  let usedRange = sheet.getUsedRange();
  
  // Get range by row/column
  let cell = sheet.getCell(0, 0); // A1
  
  // Add table
  let table = sheet.addTable("A1:D10", true); // with headers
  
  // Get tables
  let allTables = sheet.getTables();
}
```

### Range Operations

```typescript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  let range = sheet.getRange("A1:D10");
  
  // Get/set values
  let values = range.getValues();
  range.setValues([[1, 2], [3, 4]]);
  
  // Formatting
  range.getFormat().getFill().setColor("yellow");
  range.getFormat().getFont().setBold(true);
  
  // Formulas
  range.setFormula("=SUM(A1:A10)");
  
  // Address information
  let address = range.getAddress();
  let rowCount = range.getRowCount();
  let columnCount = range.getColumnCount();
}
```

### Table Operations

```typescript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  
  // Create table
  let table = sheet.addTable("A1:D10", true);
  table.setName("SalesTable");
  
  // Get table range
  let tableRange = table.getRange();
  
  // Get table columns
  let columns = table.getColumns();
  
  // Add column
  let newColumn = table.addColumn();
  
  // Get table rows
  let rows = table.getRows();
  
  // Add row
  let newRow = table.addRow();
}
```

### Chart Operations

```typescript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  
  // Create chart
  let chart = sheet.addChart(
    ExcelScript.ChartType.line,
    sheet.getRange("A1:B10")
  );
  
  // Get all charts
  let charts = sheet.getCharts();
  
  // Configure chart
  chart.setTitle("Sales Trend");
  chart.setLegendPosition(ExcelScript.ChartLegendPosition.bottom);
}
```

## Excel API Reference

### Key Excel Objects

- **Workbook**: The root object for Excel
- **Worksheet**: Individual sheets in a workbook
- **Range**: Cell or cells in a worksheet
- **Table**: Structured data in a worksheet
- **Chart**: Visual representation of data
- **PivotTable**: Data summarization tool

### Common Properties and Methods

#### Workbook

```typescript
workbook.getActiveWorksheet()
workbook.getWorksheet(name)
workbook.getWorksheets()
workbook.addWorksheet(name)
workbook.getSelectedRange()
workbook.getNamedItem(name)
```

#### Worksheet

```typescript
worksheet.getName()
worksheet.getRange(address)
worksheet.getUsedRange()
worksheet.getCell(row, column)
worksheet.getTables()
worksheet.addTable(range, hasHeaders)
worksheet.getCharts()
worksheet.addChart(type, range)
worksheet.getAutoFilter()
worksheet.getPivotTables()
```

#### Range

```typescript
range.getAddress()
range.getValues()
range.setValues(values)
range.getFormulas()
range.setFormulas(formulas)
range.getFormat()
range.getRowCount()
range.getColumnCount()
range.getCell(row, column)
range.getRow()
range.getColumn()
range.clear()
range.select()
```

#### Table

```typescript
table.getName()
table.setName(name)
table.getRange()
table.getHeaderRowRange()
table.getColumns()
table.getColumn(name)
table.addColumn()
table.getRows()
table.addRow()
table.getStyle()
table.setStyle(style)
table.getTotalRowRange()
```

#### Chart

```typescript
chart.getType()
chart.setType(type)
chart.getTitle()
chart.setTitle(title)
chart.getLegend()
chart.getAxes()
chart.getDataLabels()
chart.getHeight()
chart.setHeight(height)
chart.getWidth()
chart.setWidth(width)
```

## Practical Patterns

### Working with Data Tables

```typescript
function processTable(workbook: ExcelScript.Workbook) {
  // Get table
  const sheet = workbook.getWorksheet("Data");
  const table = sheet.getTables()[0];
  
  // Get headers and data
  const headerRange = table.getHeaderRowRange();
  const headers = headerRange.getValues()[0] as string[];
  
  const dataRange = table.getRangeBetweenHeaderAndTotal();
  const data = dataRange.getValues();
  
  // Process data
  for (let i = 0; i < data.length; i++) {
    // Process row
    for (let j = 0; j < headers.length; j++) {
      // Access cell by header name
      const value = data[i][j];
      const header = headers[j];
      
      // Do something with the value
    }
  }
}
```

### Implementing Data Validation

```typescript
function addDataValidation(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange("A1:A10");
  
  // Add dropdown validation
  const validation = range.getDataValidation();
  validation.setRule({
    list: {
      inCellDropDown: true,
      source: "Option1,Option2,Option3"
    }
  });
}
```

### Creating Calculated Columns

```typescript
function addCalculatedColumn(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const table = sheet.getTables()[0];
  
  // Add new column
  const newColumn = table.addColumn();
  
  // Set formula for entire column
  const headerRange = table.getHeaderRowRange();
  const headers = headerRange.getValues()[0] as string[];
  
  // Find index of required columns
  const priceIndex = headers.indexOf("Price");
  const quantityIndex = headers.indexOf("Quantity");
  
  // Generate the formula
  const formula = `=[@[${headers[priceIndex]}]]*[@[${headers[quantityIndex]}]]`;
  
  // Set formula and header
  const columnRange = newColumn.getRange();
  const dataRange = columnRange.getOffsetRange(1, 0, columnRange.getRowCount() - 1, 1);
  dataRange.setFormula(formula);
  
  // Set header
  newColumn.getHeaderRowRange().setValue("Total");
}
```

### Conditional Formatting

```typescript
function applyConditionalFormatting(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  
  // Color scale
  const colorScale = range.getConditionalFormats().add(
    ExcelScript.ConditionalFormatType.colorScale
  );
  
  // Configure the rule
  const rule = colorScale.getColorScale();
  rule.setCriteria({
    minimum: { formula: null, type: ExcelScript.ConditionalFormatColorScaleCriterionType.lowestValue, color: "red" },
    midpoint: { formula: null, type: ExcelScript.ConditionalFormatColorScaleCriterionType.percent, value: 50, color: "yellow" },
    maximum: { formula: null, type: ExcelScript.ConditionalFormatColorScaleCriterionType.highestValue, color: "green" }
  });
}
```

## Debugging and Testing

### Logging and Debugging

```typescript
function debugExample(workbook: ExcelScript.Workbook) {
  // Log messages to the console
  console.log("Script started");
  
  try {
    const sheet = workbook.getActiveWorksheet();
    console.log(`Working with sheet: ${sheet.getName()}`);
    
    const range = sheet.getRange("A1:D10");
    console.log(`Selected range: ${range.getAddress()}`);
    
    // Log values
    const values = range.getValues();
    console.log("Range values:", values);
  } catch (error) {
    console.error("Error occurred:", error.message);
    throw error; // Rethrow to see in script editor
  }
  
  console.log("Script completed");
}
```

### Error Handling

```typescript
function robustExample(workbook: ExcelScript.Workbook) {
  try {
    // Check if sheet exists
    const sheetName = "Data";
    let sheet: ExcelScript.Worksheet;
    
    try {
      sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        throw new Error(`Sheet "${sheetName}" not found`);
      }
    } catch {
      // Create sheet if it doesn't exist
      sheet = workbook.addWorksheet(sheetName);
      console.log(`Created new sheet "${sheetName}"`);
    }
    
    // Check if range has data
    const dataRange = sheet.getRange("A1:C10");
    const values = dataRange.getValues();
    
    // Check for empty data
    if (values.every(row => row.every(cell => cell === ""))) {
      console.log("Range is empty, initializing with headers");
      sheet.getRange("A1:C1").setValues([["Date", "Product", "Amount"]]);
    }
    
    // Proceed with main logic
    
  } catch (error) {
    console.error("Script failed:", error.message);
    // Optionally write error to sheet for user to see
    const errorSheet = workbook.addWorksheet("Error Log");
    errorSheet.getRange("A1").setValue(`Error: ${error.message}`);
  }
}
```

## Exercises

### Exercise 1: Basic Data Formatting

**Objective**: Create a script that formats a basic dataset with alternate row colors and proper headers.

```typescript
function formatDataset(workbook: ExcelScript.Workbook) {
  // Get active sheet
  const sheet = workbook.getActiveWorksheet();
  
  // Define the data range
  const dataRange = sheet.getRange("A1:D10");
  
  // Format headers
  const headerRange = sheet.getRange("A1:D1");
  headerRange.getFormat().getFill().setColor("#4472C4");
  headerRange.getFormat().getFont().setColor("white");
  headerRange.getFormat().getFont().setBold(true);
  
  // Set default data format
  dataRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  
  // Apply alternating row colors
  for (let i = 1; i < 10; i++) {
    const rowRange = sheet.getRange(`A${i+1}:D${i+1}`);
    if (i % 2 === 0) {
      rowRange.getFormat().getFill().setColor("#D9E1F2");
    } else {
      rowRange.getFormat().getFill().setColor("#FFFFFF");
    }
  }
  
  // Add borders
  dataRange.getFormat().getBorders().getEdges().setStyle(ExcelScript.BorderLineStyle.thin);
  
  // Auto-fit columns
  dataRange.getFormat().autofitColumns();
}
```

### Exercise 2: Data Validation and Entry Form

**Objective**: Create a simple data entry form with validation rules.

```typescript
function createEntryForm(workbook: ExcelScript.Workbook) {
  // Get or create worksheets
  let formSheet = workbook.getWorksheet("Entry Form");
  if (!formSheet) {
    formSheet = workbook.addWorksheet("Entry Form");
  }
  
  let dataSheet = workbook.getWorksheet("Data");
  if (!dataSheet) {
    dataSheet = workbook.addWorksheet("Data");
    // Initialize data sheet with headers
    dataSheet.getRange("A1:D1").setValues([["Date", "Category", "Amount", "Description"]]);
  }
  
  // Clear and set up form sheet
  formSheet.getRange("A1:D20").clear();
  
  // Create form title
  formSheet.getRange("A1").setValue("DATA ENTRY FORM");
  formSheet.getRange("A1").getFormat().getFont().setBold(true);
  formSheet.getRange("A1").getFormat().getFont().setSize(14);
  
  // Create form labels
  formSheet.getRange("A3").setValue("Date:");
  formSheet.getRange("A4").setValue("Category:");
  formSheet.getRange("A5").setValue("Amount:");
  formSheet.getRange("A6").setValue("Description:");
  
  // Create input cells
  const dateCell = formSheet.getRange("B3");
  const categoryCell = formSheet.getRange("B4");
  const amountCell = formSheet.getRange("B5");
  const descriptionCell = formSheet.getRange("B6");
  
  // Set date validation
  dateCell.setNumberFormat("yyyy-mm-dd");
  
  // Set category validation (dropdown)
  const categories = ["Income", "Housing", "Food", "Transportation", "Entertainment", "Utilities", "Other"];
  const validation = categoryCell.getDataValidation();
  validation.setRule({
    list: {
      inCellDropDown: true,
      source: categories.join(",")
    }
  });
  
  // Set amount validation (number only)
  amountCell.setNumberFormat("$#,##0.00");
  amountCell.getDataValidation().setRule({
    decimal: {
      formula1: "0",
      operator: ExcelScript.DataValidationOperator.greaterThan
    }
  });
  
  // Add form button (visual only, cannot add functionality)
  formSheet.getRange("B8").setValue("SUBMIT");
  formSheet.getRange("B8").getFormat().getFill().setColor("#4472C4");
  formSheet.getRange("B8").getFormat().getFont().setColor("white");
  formSheet.getRange("B8").getFormat().getFont().setBold(true);
  formSheet.getRange("B8").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
}
```

### Exercise 3: Sales Data Analysis

**Objective**: Create a script that processes sales data, calculates metrics, and visualizes results.

```typescript
function analyzeSalesData(workbook: ExcelScript.Workbook) {
  // Get data sheet
  const dataSheet = workbook.getWorksheet("Sales Data");
  if (!dataSheet) {
    throw new Error("Sales Data worksheet not found");
  }
  
  // Get or create analysis sheet
  let analysisSheet = workbook.getWorksheet("Sales Analysis");
  if (!analysisSheet) {
    analysisSheet = workbook.addWorksheet("Sales Analysis");
  } else {
    analysisSheet.getRange("A1:Z100").clear(); // Clear existing content
  }
  
  // Get sales data range
  const dataRange = dataSheet.getUsedRange();
  const values = dataRange.getValues();
  
  // Extract headers
  const headers = values[0] as string[];
  const dateIndex = headers.indexOf("Date");
  const productIndex = headers.indexOf("Product");
  const regionIndex = headers.indexOf("Region");
  const salesIndex = headers.indexOf("Sales");
  
  if (dateIndex < 0 || productIndex < 0 || regionIndex < 0 || salesIndex < 0) {
    throw new Error("Required columns not found in data");
  }
  
  // Extract data (skip header row)
  const salesData = values.slice(1);
  
  // Calculate total sales
  let totalSales = 0;
  for (const row of salesData) {
    totalSales += row[salesIndex] as number;
  }
  
  // Calculate sales by product
  const productSales = {};
  for (const row of salesData) {
    const product = row[productIndex] as string;
    const sales = row[salesIndex] as number;
    
    if (!productSales[product]) {
      productSales[product] = 0;
    }
    productSales[product] += sales;
  }
  
  // Calculate sales by region
  const regionSales = {};
  for (const row of salesData) {
    const region = row[regionIndex] as string;
    const sales = row[salesIndex] as number;
    
    if (!regionSales[region]) {
      regionSales[region] = 0;
    }
    regionSales[region] += sales;
  }
  
  // Output summary
  analysisSheet.getRange("A1").setValue("Sales Analysis");
  analysisSheet.getRange("A1").getFormat().getFont().setBold(true);
  analysisSheet.getRange("A1").getFormat().getFont().setSize(16);
  
  // Total sales
  analysisSheet.getRange("A3").setValue("Total Sales:");
  analysisSheet.getRange("B3").setValue(totalSales);
  analysisSheet.getRange("B3").setNumberFormat("$#,##0.00");
  
  // Sales by product
  analysisSheet.getRange("A5").setValue("Sales by Product");
  analysisSheet.getRange("A5").getFormat().getFont().setBold(true);
  
  // Headers
  analysisSheet.getRange("A6").setValue("Product");
  analysisSheet.getRange("B6").setValue("Sales");
  
  // Data
  let row = 7;
  for (const product in productSales) {
    analysisSheet.getRange(`A${row}`).setValue(product);
    analysisSheet.getRange(`B${row}`).setValue(productSales[product]);
    analysisSheet.getRange(`B${row}`).setNumberFormat("$#,##0.00");
    row++;
  }
  
  // Create product sales chart
  const productChartRange = analysisSheet.getRange(`A6:B${row-1}`);
  const productChart = analysisSheet.addChart(
    ExcelScript.ChartType.columnClustered,
    productChartRange
  );
  productChart.setTitle("Sales by Product");
  productChart.setLeft(300);
  productChart.setTop(50);
  productChart.setWidth(400);
  productChart.setHeight(300);
  
  // Sales by region
  analysisSheet.getRange("A" + (row + 1)).setValue("Sales by Region");
  analysisSheet.getRange("A" + (row + 1)).getFormat().getFont().setBold(true);
  
  // Headers
  analysisSheet.getRange("A" + (row + 2)).setValue("Region");
  analysisSheet.getRange("B" + (row + 2)).setValue("Sales");
  
  // Data
  let regionRow = row + 3;
  for (const region in regionSales) {
    analysisSheet.getRange(`A${regionRow}`).setValue(region);
    analysisSheet.getRange(`B${regionRow}`).setValue(regionSales[region]);
    analysisSheet.getRange(`B${regionRow}`).setNumberFormat("$#,##0.00");
    regionRow++;
  }
  
  // Create region sales chart
  const regionChartRange = analysisSheet.getRange(`A${row+2}:B${regionRow-1}`);
  const regionChart = analysisSheet.addChart(
    ExcelScript.ChartType.pie,
    regionChartRange
  );
  regionChart.setTitle("Sales by Region");
  regionChart.setLeft(300);
  regionChart.setTop(400);
  regionChart.setWidth(400);
  regionChart.setHeight(300);
}
```

### Exercise 4: Inventory Management System

**Objective**: Create a more complex inventory management system with multiple sheets and features.

```typescript
function setupInventorySystem(workbook: ExcelScript.Workbook) {
  // Initialize required sheets
  const sheets = {
    inventory: workbook.getWorksheet("Inventory") || workbook.addWorksheet("Inventory"),
    transactions: workbook.getWorksheet("Transactions") || workbook.addWorksheet("Transactions"),
    dashboard: workbook.getWorksheet("Dashboard") || workbook.addWorksheet("Dashboard")
  };
  
  // Setup inventory sheet
  setupInventorySheet(sheets.inventory);
  
  // Setup transactions sheet
  setupTransactionsSheet(sheets.transactions);
  
  // Setup dashboard
  setupDashboard(workbook, sheets.dashboard);
  
  function setupInventorySheet(sheet: ExcelScript.Worksheet) {
    // Clear existing content
    sheet.getRange("A1:Z100").clear();
    
    // Set headers
    sheet.getRange("A1:F1").setValues([["Item ID", "Name", "Category", "Quantity", "Unit Price", "Value"]]);
    
    // Format headers
    sheet.getRange("A1:F1").getFormat().getFill().setColor("#4472C4");
    sheet.getRange("A1:F1").getFormat().getFont().setColor("white");
    sheet.getRange("A1:F1").getFormat().getFont().setBold(true);
    
    // Add formula for value calculation
    sheet.getRange("F2:F100").setFormula("=D2*E2");
    
    // Format columns
    sheet.getRange("D:D").setNumberFormat("0");
    sheet.getRange("E:E").setNumberFormat("$#,##0.00");
    sheet.getRange("F:F").setNumberFormat("$#,##0.00");
    
    // Add data validation for categories
    const categories = ["Electronics", "Office Supplies", "Furniture", "Food", "Clothing"];
    const validation = sheet.getRange("C2:C100").getDataValidation();
    validation.setRule({
      list: {
        inCellDropDown: true,
        source: categories.join(",")
      }
    });
    
    // Add sample data
    sheet.getRange("A2:E4").setValues([
      ["IT001", "Laptop", "Electronics", 10, 1200],
      ["IT002", "Chair", "Furniture", 25, 150],
      ["IT003", "Pens", "Office Supplies", 100, 1.5]
    ]);
    
    // Create table
    const tableRange = sheet.getRange("A1:F4");
    const table = sheet.addTable(tableRange, true);
    table.setName("InventoryTable");
    
    // Apply conditional formatting for low stock
    const qtyRange = sheet.getRange("D2:D100");
    const lowStock = qtyRange.getConditionalFormats().add(ExcelScript.ConditionalFormatType.cellValue);
    lowStock.getCellValue().setRule({
      formula1: "10",
      operator: ExcelScript.ConditionalCellValueOperator.lessThanOrEqual,
      format: {
        fill: { color: "#FF9999" },
        font: { bold: true }
      }
    });
  }
  
  function setupTransactionsSheet(sheet: ExcelScript.Worksheet) {
    // Clear existing content
    sheet.getRange("A1:Z100").clear();
    
    // Set headers
    sheet.getRange("A1:F1").setValues([["Date", "Item ID", "Transaction Type", "Quantity", "Unit Price", "Total Value"]]);
    
    // Format headers
    sheet.getRange("A1:F1").getFormat().getFill().setColor("#4472C4");
    sheet.getRange("A1:F1").getFormat().getFont().setColor("white");
    sheet.getRange("A1:F1").getFormat().getFont().setBold(true);
    
    // Format columns
    sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
    sheet.getRange("D:D").setNumberFormat("0");
    sheet.getRange("E:E").setNumberFormat("$#,##0.00");
    sheet.getRange("F:F").setNumberFormat("$#,##0.00");
    
    // Add formula for value calculation
    sheet.getRange("F2:F100").setFormula("=D2*E2");
    
    // Add data validation for transaction type
    const transactionTypes = ["Purchase", "Sale", "Return", "Adjustment"];
    const validation = sheet.getRange("C2:C100").getDataValidation();
    validation.setRule({
      list: {
        inCellDropDown: true,
        source: transactionTypes.join(",")
      }
    });
    
    // Add sample data
    sheet.getRange("A2:E4").setValues([
      [new Date(2023, 0, 15), "IT001", "Purchase", 5, 1200],
      [new Date(2023, 0, 20), "IT002", "Sale", 2, 150],
      [new Date(2023, 0, 25), "IT001", "Sale", 1, 1300]
    ]);
    
    // Create table
    const tableRange = sheet.getRange("A1:F4");
    const table = sheet.addTable(tableRange, true);
    table.setName("TransactionsTable");
    
    // Apply conditional formatting for transaction type
    const typeRange = sheet.getRange("C2:C100");
    
    // Format for purchases
    const purchaseFormat = typeRange.getConditionalFormats().add(ExcelScript.ConditionalFormatType.cellValue);
    purchaseFormat.getCellValue().setRule({
      formula1: "Purchase",
      operator: ExcelScript.ConditionalCellValueOperator.equalTo,
      format: { fill: { color: "#C6EFCE" } }
    });
    
    // Format for sales
    const salesFormat = typeRange.getConditionalFormats().add(ExcelScript.ConditionalFormatType.cellValue);
    salesFormat.getCellValue().setRule({
      formula1: "Sale",
      operator: ExcelScript.ConditionalCellValueOperator.equalTo,
      format: { fill: { color: "#FFEB9C" } }
    });
  }
  
  function setupDashboard(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet) {
    // Clear existing content
    sheet.getRange("A1:Z100").clear();
    
    // Set title
    sheet.getRange("A1").setValue("INVENTORY MANAGEMENT DASHBOARD");
    sheet.getRange("A1").getFormat().getFont().setBold(true);
    sheet.getRange("A1").getFormat().getFont().setSize(16);
    
    // Create summary section
    sheet.getRange("A3").setValue("Inventory Summary");
    sheet.getRange("A3").getFormat().getFont().setBold(true);
    
    sheet.getRange("A4:B7").setValues([
      ["Total Items:", "=COUNTA(Inventory!A:A)-1"],
      ["Total Value:", "=SUM(Inventory!F:F)"],
      ["Low Stock Items:", "=COUNTIFS(Inventory!D:D,\"<=10\")"],
      ["Categories:", "=COUNTA(UNIQUE(Inventory!C:C))"]
    ]);
    
    // Format summary values
    sheet.getRange("B5").setNumberFormat("$#,##0.00");
    
    // Create transaction metrics
    sheet.getRange("D3").setValue("Transaction Metrics");
    sheet.getRange("D3").getFormat().getFont().setBold(true);
    
    sheet.getRange("D4:E7").setValues([
      ["Total Purchases:", "=SUMIFS(Transactions!F:F,Transactions!C:C,\"Purchase\")"],
      ["Total Sales:", "=SUMIFS(Transactions!F:F,Transactions!C:C,\"Sale\")"],
      ["Net Movement:", "=E4-E5"],
      ["Last Transaction:", "=MAX(Transactions!A:A)"]
    ]);
    
    // Format transaction metrics
    sheet.getRange("E4:E6").setNumberFormat("$#,##0.00");
    sheet.getRange("E7").setNumberFormat("yyyy-mm-dd");
    
    // Format cells
    sheet.getRange("A4:E7").getFormat().getBorders().getEdges().setStyle(ExcelScript.BorderLineStyle.thin);
    
    // Create inventory value chart
    const inventorySheet = workbook.getWorksheet("Inventory");
    const inventoryTable = inventorySheet.getTable("InventoryTable");
    const chartDataRange = inventoryTable.getRange();
    
    const valueChart = sheet.addChart(
      ExcelScript.ChartType.columnClustered,
      chartDataRange
    );
    valueChart.setTitle("Inventory Value by Item");
    valueChart.setLeft(50);
    valueChart.setTop(150);
    valueChart.setWidth(300);
    valueChart.setHeight(200);
    
    // Modify X axis to use item names
    const xAxis = valueChart.getAxes().getCategoryAxis();
    xAxis.setTitle("Items");
    
    // Set Y axis to represent values
    const yAxis = valueChart.getAxes().getValueAxis();
    yAxis.setTitle("Value ($)");
    
    // Only show value series
    const series = valueChart.getSeries();
    for (let i = 0; i < series.getCount(); i++) {
      const currentSeries = series.getItemAt(i);
      if (currentSeries.getName() !== "Value") {
        currentSeries.setVisible(false);
      }
    }
    
    // Create category distribution chart
    const categoryChart = sheet.addChart(
      ExcelScript.ChartType.pie,
      inventoryTable.getRange()
    );
    categoryChart.setTitle("Items by Category");
    categoryChart.setLeft(400);
    categoryChart.setTop(150);
    categoryChart.setWidth(300);
    categoryChart.setHeight(200);
    
    // Configure series for categories
    const catSeries = categoryChart.getSeries();
    for (let i = 0; i < catSeries.getCount(); i++) {
      const currentSeries = catSeries.getItemAt(i);
      currentSeries.setVisible(false);
    }
    
    // Add transaction trend chart
    const transactionSheet = workbook.getWorksheet("Transactions");
    const transactionTable = transactionSheet.getTable("TransactionsTable");
    
    const trendChart = sheet.addChart(
      ExcelScript.ChartType.line,
      transactionTable.getRange()
    );
    trendChart.setTitle("Transaction History");
    trendChart.setLeft(50);
    trendChart.setTop(400);
    trendChart.setWidth(650);
    trendChart.setHeight(250);
  }
}
```

### Exercise 5: Advanced Excel Automation System

**Objective**: Build a comprehensive system that automates report generation, data processing, and visualization with complex Excel features.

```typescript
function automatedReportingSystem(workbook: ExcelScript.Workbook) {
  // Track execution time
  const startTime = new Date();
  
  // Configuration
  const config = {
    dataSources: ["Sales", "Inventory", "Customers"],
    reportPeriod: "Monthly",
    generateCharts: true,
    exportFormat: "Excel",
    emailRecipients: ["manager@company.com"]
  };
  
  // Setup logging
  const logSheet = getOrCreateWorksheet(workbook, "System Log");
  log(logSheet, "INFO", "Starting automated report generation");
  
  try {
    // Process each data source
    for (const source of config.dataSources) {
      log(logSheet, "INFO", `Processing ${source} data`);
      
      // Get or create source sheet
      const sourceSheet = getOrCreateWorksheet(workbook, source);
      
      // Get or create report sheet
      const reportSheet = getOrCreateWorksheet(workbook, `${source} Report`);
      reportSheet.getRange("A1:Z500").clear(); // Clear previous report
      
      // Process data based on source type
      switch (source) {
        case "Sales":
          processSalesData(workbook, sourceSheet, reportSheet);
          break;
        case "Inventory":
          processInventoryData(workbook, sourceSheet, reportSheet);
          break;
        case "Customers":
          processCustomerData(workbook, sourceSheet, reportSheet);
          break;
      }
      
      log(logSheet, "INFO", `Completed processing ${source} data`);
    }
    
    // Generate executive summary
    generateExecutiveSummary(workbook);
    
    // Record completion
    const endTime = new Date();
    const executionTime = (endTime.getTime() - startTime.getTime()) / 1000;
    log(logSheet, "INFO", `Report generation completed in ${executionTime} seconds`);
    
  } catch (error) {
    log(logSheet, "ERROR", `Failed: ${error.message}`);
    
    // Create error notification sheet
    const errorSheet = getOrCreateWorksheet(workbook, "Error Report");
    errorSheet.getRange("A1:B1").setValues([["Error Timestamp", "Error Message"]]);
    errorSheet.getRange("A2:B2").setValues([[new Date().toISOString(), error.message]]);
    
    // Format error sheet
    errorSheet.getRange("A1:B1").getFormat().getFill().setColor("#FF0000");
    errorSheet.getRange("A1:B1").getFormat().getFont().setColor("white");
    errorSheet.getRange("A1:B1").getFormat().getFont().setBold(true);
  }
  
  // Helper functions
  function getOrCreateWorksheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
    let sheet = workbook.getWorksheet(name);
    if (!sheet) {
      sheet = workbook.addWorksheet(name);
    }
    return sheet;
  }
  
  function log(sheet: ExcelScript.Worksheet, level: string, message: string) {
    // Find next empty row
    const range = sheet.getUsedRange();
    const nextRow = range ? range.getRowCount() + 1 : 1;
    
    // Check if headers exist
    if (nextRow === 1) {
      sheet.getRange("A1:C1").setValues([["Timestamp", "Level", "Message"]]);
      sheet.getRange("A1:C1").getFormat().getFill().setColor("#4472C4");
      sheet.getRange("A1:C1").getFormat().getFont().setColor("white");
      sheet.getRange("A1:C1").getFormat().getFont().setBold(true);
    }
    
    // Log entry
    sheet.getRange(`A${nextRow}:C${nextRow}`).setValues([
      [new Date().toISOString(), level, message]
    ]);
    
    // Apply formatting based on level
    let levelColor = "#FFFFFF";
    switch (level) {
      case "ERROR":
        levelColor = "#FF9999";
        break;
      case "WARNING":
        levelColor = "#FFEB9C";
        break;
      case "INFO":
        levelColor = "#C6EFCE";
        break;
    }
    
    sheet.getRange(`B${nextRow}`).getFormat().getFill().setColor(levelColor);
  }
  
  function processSalesData(workbook: ExcelScript.Workbook, sourceSheet: ExcelScript.Worksheet, reportSheet: ExcelScript.Worksheet) {
    // Check if source data exists
    const sourceRange = sourceSheet.getUsedRange();
    if (!sourceRange) {
      throw new Error("No sales data found");
    }
    
    // Get data
    const values = sourceRange.getValues();
    const headers = values[0] as string[];
    
    // Identify required columns
    const dateCol = headers.indexOf("Date");
    const productCol = headers.indexOf("Product");
    const quantityCol = headers.indexOf("Quantity");
    const priceCol = headers.indexOf("Price");
    const regionCol = headers.indexOf("Region");
    
    if (dateCol < 0 || productCol < 0 || quantityCol < 0 || priceCol < 0) {
      throw new Error("Sales data missing required columns");
    }
    
    // Set report title
    reportSheet.getRange("A1").setValue("SALES PERFORMANCE REPORT");
    reportSheet.getRange("A1").getFormat().getFont().setBold(true);
    reportSheet.getRange("A1").getFormat().getFont().setSize(16);
    reportSheet.getRange("A2").setValue(`Generated: ${new Date().toLocaleDateString()}`);
    
    // Extract data (skip header)
    const data = values.slice(1);
    
    // Calculate sales metrics
    const metrics = {
      totalSales: 0,
      totalItems: 0,
      avgPrice: 0,
      topProduct: "",
      topProductSales: 0
    };
    
    // Track sales by product and region
    const productSales = {};
    const regionSales = {};
    const monthlySales = {};
    
    for (const row of data) {
      const date = row[dateCol] as Date;
      const product = row[productCol] as string;
      const quantity = row[quantityCol] as number;
      const price = row[priceCol] as number;
      const region = regionCol >= 0 ? row[regionCol] as string : "Unknown";
      
      const saleAmount = quantity * price;
      
      // Update totals
      metrics.totalSales += saleAmount;
      metrics.totalItems += quantity;
      
      // Update product sales
      if (!productSales[product]) {
        productSales[product] = 0;
      }
      productSales[product] += saleAmount;
      
      // Update region sales
      if (!regionSales[region]) {
        regionSales[region] = 0;
      }
      regionSales[region] += saleAmount;
      
      // Update monthly sales
      if (date instanceof Date) {
        const yearMonth = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, "0")}`;
        if (!monthlySales[yearMonth]) {
          monthlySales[yearMonth] = 0;
        }
        monthlySales[yearMonth] += saleAmount;
      }
      
      // Check for top product
      if (productSales[product] > metrics.topProductSales) {
        metrics.topProduct = product;
        metrics.topProductSales = productSales[product];
      }
    }
    
    // Calculate average price
    metrics.avgPrice = metrics.totalItems > 0 ? metrics.totalSales / metrics.totalItems : 0;
    
    // Display summary metrics
    reportSheet.getRange("A4").setValue("Sales Summary");
    reportSheet.getRange("A4").getFormat().getFont().setBold(true);
    
    reportSheet.getRange("A5:B9").setValues([
      ["Total Sales:", metrics.totalSales],
      ["Total Items Sold:", metrics.totalItems],
      ["Average Price:", metrics.avgPrice],
      ["Top Product:", metrics.topProduct],
      ["Top Product Sales:", metrics.topProductSales]
    ]);
    
    reportSheet.getRange("B5").setNumberFormat("$#,##0.00");
    reportSheet.getRange("B7").setNumberFormat("$#,##0.00");
    reportSheet.getRange("B9").setNumberFormat("$#,##0.00");
    
    // Sales by product table
    reportSheet.getRange("A11").setValue("Sales by Product");
    reportSheet.getRange("A11").getFormat().getFont().setBold(true);
    
    reportSheet.getRange("A12:B12").setValues([["Product", "Sales"]]);
    reportSheet.getRange("A12:B12").getFormat().getFill().setColor("#4472C4");
    reportSheet.getRange("A12:B12").getFormat().getFont().setColor("white");
    reportSheet.getRange("A12:B12").getFormat().getFont().setBold(true);
    
    let row = 13;
    for (const product in productSales) {
      reportSheet.getRange(`A${row}:B${row}`).setValues([[product, productSales[product]]]);
      reportSheet.getRange(`B${row}`).setNumberFormat("$#,##0.00");
      row++;
    }
    
    // Create product sales table
    const productTableRange = reportSheet.getRange(`A12:B${row-1}`);
    reportSheet.addTable(productTableRange, true);
    
    // Create product sales chart if enabled
    if (config.generateCharts) {
      const productChart = reportSheet.addChart(
        ExcelScript.ChartType.columnClustered,
        productTableRange
      );
      productChart.setTitle("Sales by Product");
      productChart.setLeft(400);
      productChart.setTop(50);
      productChart.setWidth(450);
      productChart.setHeight(250);
    }
    
    // Sales by region table
    const regionStartRow = row + 2;
    reportSheet.getRange(`A${regionStartRow}`).setValue("Sales by Region");
    reportSheet.getRange(`A${regionStartRow}`).getFormat().getFont().setBold(true);
    
    reportSheet.getRange(`A${regionStartRow+1}:B${regionStartRow+1}`).setValues([["Region", "Sales"]]);
    reportSheet.getRange(`A${regionStartRow+1}:B${regionStartRow+1}`).getFormat().getFill().setColor("#4472C4");
    reportSheet.getRange(`A${regionStartRow+1}:B${regionStartRow+1}`).getFormat().getFont().setColor("white");
    reportSheet.getRange(`A${regionStartRow+1}:B${regionStartRow+1}`).getFormat().getFont().setBold(true);
    
    row = regionStartRow + 2;
    for (const region in regionSales) {
      reportSheet.getRange(`A${row}:B${row}`).setValues([[region, regionSales[region]]]);
      reportSheet.getRange(`B${row}`).setNumberFormat("$#,##0.00");
      row++;
    }
    
    // Create region sales table
    const regionTableRange = reportSheet.getRange(`A${regionStartRow+1}:B${row-1}`);
    reportSheet.addTable(regionTableRange, true);
    
    // Create region sales chart if enabled
    if (config.generateCharts) {
      const regionChart = reportSheet.addChart(
        ExcelScript.ChartType.pie,
        regionTableRange
      );
      regionChart.setTitle("Sales by Region");
      regionChart.setLeft(400);
      regionChart.setTop(350);
      regionChart.setWidth(450);
      regionChart.setHeight(250);
    }
    
    // Monthly trend table
    const monthlyStartRow = row + 2;
    reportSheet.getRange(`A${monthlyStartRow}`).setValue("Monthly Sales Trend");
    reportSheet.getRange(`A${monthlyStartRow}`).getFormat().getFont().setBold(true);
    
    reportSheet.getRange(`A${monthlyStartRow+1}:B${monthlyStartRow+1}`).setValues([["Month", "Sales"]]);
    reportSheet.getRange(`A${monthlyStartRow+1}:B${monthlyStartRow+1}`).getFormat().getFill().setColor("#4472C4");
    reportSheet.getRange(`A${monthlyStartRow+1}:B${monthlyStartRow+1}`).getFormat().getFont().setColor("white");
    reportSheet.getRange(`A${monthlyStartRow+1}:B${monthlyStartRow+1}`).getFormat().getFont().setBold(true);
    
    row = monthlyStartRow + 2;
    
    // Sort months chronologically
    const months = Object.keys(monthlySales).sort();
    for (const month of months) {
      reportSheet.getRange(`A${row}:B${row}`).setValues([[month, monthlySales[month]]]);
      reportSheet.getRange(`B${row}`).setNumberFormat("$#,##0.00");
      row++;
    }
    
    // Create monthly trend table
    const monthlyTableRange = reportSheet.getRange(`A${monthlyStartRow+1}:B${row-1}`);
    reportSheet.addTable(monthlyTableRange, true);
    
    // Create monthly trend chart if enabled
    if (config.generateCharts) {
      const trendChart = reportSheet.addChart(
        ExcelScript.ChartType.line,
        monthlyTableRange
      );
      trendChart.setTitle("Monthly Sales Trend");
      trendChart.setLeft(50);
      trendChart.setTop(600);
      trendChart.setWidth(800);
      trendChart.setHeight(250);
    }
  }
  
  function processInventoryData(workbook: ExcelScript.Workbook, sourceSheet: ExcelScript.Worksheet, reportSheet: ExcelScript.Worksheet) {
    // Similar implementation as processSalesData but for inventory
    // Check if source data exists
    const sourceRange = sourceSheet.getUsedRange();
    if (!sourceRange) {
      throw new Error("No inventory data found");
    }
    
    // Set report title
    reportSheet.getRange("A1").setValue("INVENTORY STATUS REPORT");
    reportSheet.getRange("A1").getFormat().getFont().setBold(true);
    reportSheet.getRange("A1").getFormat().getFont().setSize(16);
    reportSheet.getRange("A2").setValue(`Generated: ${new Date().toLocaleDateString()}`);
    
    // Add placeholder implementation
    reportSheet.getRange("A4").setValue("Inventory analysis functionality implemented");
    reportSheet.getRange("A5").setValue("See Sales report for detailed implementation pattern");
  }
  
  function processCustomerData(workbook: ExcelScript.Workbook, sourceSheet: ExcelScript.Worksheet, reportSheet: ExcelScript.Worksheet) {
    // Similar implementation as processSalesData but for customers
    // Check if source data exists
    const sourceRange = sourceSheet.getUsedRange();
    if (!sourceRange) {
      throw new Error("No customer data found");
    }
    
    // Set report title
    reportSheet.getRange("A1").setValue("CUSTOMER ANALYSIS REPORT");
    reportSheet.getRange("A1").getFormat().getFont().setBold(true);
    reportSheet.getRange("A1").getFormat().getFont().setSize(16);
    reportSheet.getRange("A2").setValue(`Generated: ${new Date().toLocaleDateString()}`);
    
    // Add placeholder implementation
    reportSheet.getRange("A4").setValue("Customer analysis functionality implemented");
    reportSheet.getRange("A5").setValue("See Sales report for detailed implementation pattern");
  }
  
  function generateExecutiveSummary(workbook: ExcelScript.Workbook) {
    // Create executive summary sheet
    const summarySheet = getOrCreateWorksheet(workbook, "Executive Summary");
    summarySheet.getRange("A1:Z100").clear();
    
    // Set title
    summarySheet.getRange("A1").setValue("EXECUTIVE SUMMARY");
    summarySheet.getRange("A1").getFormat().getFont().setBold(true);
    summarySheet.getRange("A1").getFormat().getFont().setSize(18);
    summarySheet.getRange("A2").setValue(`Report Generated: ${new Date().toLocaleDateString()}`);
    
    // Create overview section
    summarySheet.getRange("A4").setValue("Performance Overview");
    summarySheet.getRange("A4").getFormat().getFont().setBold(true);
    summarySheet.getRange("A4").getFormat().getFont().setSize(14);
    
    // Add automated references to source reports
    summarySheet.getRange("A5:B9").setValues([
      ["Total Sales:", "='Sales Report'!B5"],
      ["Top Product:", "='Sales Report'!B8"],
      ["Inventory Value:", "=SUM('Inventory'!E:E)"],
      ["Low Stock Items:", "=COUNTIFS('Inventory'!D:D,\"<=10\")"],
      ["Active Customers:", "=COUNTA('Customers'!A:A)-1"]
    ]);
    
    // Format summary
    summarySheet.getRange("A5:A9").getFormat().getFont().setBold(true);
    summarySheet.getRange("B6").setNumberFormat("$#,##0.00");
    
    // Create key metrics summary table
    summarySheet.getRange("A11").setValue("Key Performance Indicators");
    summarySheet.getRange("A11").getFormat().getFont().setBold(true);
    summarySheet.getRange("A11").getFormat().getFont().setSize(14);
    
    summarySheet.getRange("A12:C12").setValues([["Metric", "Current", "Target"]]);
    summarySheet.getRange("A12:C12").getFormat().getFill().setColor("#4472C4");
    summarySheet.getRange("A12:C12").getFormat().getFont().setColor("white");
    summarySheet.getRange("A12:C12").getFormat().getFont().setBold(true);
    
    // Add some dynamic KPIs
    summarySheet.getRange("A13:C17").setValues([
      ["Sales Growth", "=('Sales Report'!B5-10000)/10000", "15%"],
      ["Inventory Turnover", "='Sales Report'!B5/SUM('Inventory'!E:E)", "4"],
      ["Average Order Value", "='Sales Report'!B5/'Sales Report'!B6", "$100"],
      ["Return Rate", "5%", "<3%"],
      ["Customer Retention", "82%", ">85%"]
    ]);
    
    // Format KPIs
    summarySheet.getRange("B13:B17").setNumberFormat("0.0%");
    summarySheet.getRange("C13:C14").setNumberFormat("0.0");
    summarySheet.getRange("C15").setNumberFormat("$#,##0.00");
    summarySheet.getRange("C16:C17").setNumberFormat("0.0%");
    
    // Add conditional formatting for KPIs
    const kpiRange = summarySheet.getRange("B13:B17");
    const targetRange = summarySheet.getRange("C13:C17");
    
    // Apply conditional formatting
    const kpiFormat = kpiRange.getConditionalFormats().add(ExcelScript.ConditionalFormatType.iconSet);
    kpiFormat.getIconSet().setStyle(ExcelScript.IconSet.threeTriangles);
    
    // Add summary chart if enabled
    if (config.generateCharts) {
      // Create a combination chart for summary metrics
      const summaryChart = summarySheet.addChart(
        ExcelScript.ChartType.columnClustered,
        summarySheet.getRange("A12:C17") 
      );
      summaryChart.setTitle("KPI Summary");
      summaryChart.setLeft(400);
      summaryChart.setTop(50);
      summaryChart.setWidth(500);
      summaryChart.setHeight(300);
    }
    
    // Create report links section
    summarySheet.getRange("A20").setValue("Detailed Reports");
    summarySheet.getRange("A20").getFormat().getFont().setBold(true);
    summarySheet.getRange("A20").getFormat().getFont().setSize(14);
    
    // Add hyperlinks to reports (visual only - can't create actual hyperlinks)
    summarySheet.getRange("A21").setValue("→ Sales Performance Report");
    summarySheet.getRange("A21").getFormat().getFont().setColor("#0563C1");
    summarySheet.getRange("A21").getFormat().getFont().setUnderline(true);
    
    summarySheet.getRange("A22").setValue("→ Inventory Status Report");
    summarySheet.getRange("A22").getFormat().getFont().setColor("#0563C1");
    summarySheet.getRange("A22").getFormat().getFont().setUnderline(true);
    
    summarySheet.getRange("A23").setValue("→ Customer Analysis Report");
    summarySheet.getRange("A23").getFormat().getFont().setColor("#0563C1");
    summarySheet.getRange("A23").getFormat().getFont().setUnderline(true);
  }
}
```

## Additional Topics in TypeScript for Excel

### Working with Events (Office Add-ins)

```typescript
// In an Office Add-in
Office.onReady(() => {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Handle worksheet selection change event
    sheet.onSelectionChanged.add(handleSelectionChange);
    
    await context.sync();
  });
});

async function handleSelectionChange(event: Excel.WorksheetSelectionChangedEventArgs) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    
    await context.sync();
    
    console.log(`Selection changed to ${range.address}`);
  });
}
```

### Custom Functions (Office Add-ins)

```typescript
/**
 * Calculates compound interest.
 * @customfunction
 * @param principal The principal amount
 * @param rate The annual interest rate
 * @param periods The number of periods
 * @param [frequency=1] Compounding frequency per year
 * @returns The future value
 */
function COMPOUND_INTEREST(
  principal: number,
  rate: number,
  periods: number,
  frequency: number = 1
): number {
  const r = rate/frequency;
  const n = periods * frequency;
  return principal * Math.pow(1 + r, n);
}

/**
 * Gets the current stock price (example).
 * @customfunction
 * @param symbol The stock symbol
 * @returns The current stock price
 */
async function STOCK_PRICE(symbol: string): Promise<number> {
  // This would need actual implementation to fetch stock price
  return 150.75;
}
```

### Power Query Integration

```typescript
// TypeScript interface for a Power Query result
interface PowerQueryResult {
  tableName: string;
  columns: string[];
  rows: any[][];
}

// Example of handling Power Query data
function processPowerQueryData(workbook: ExcelScript.Workbook, queryName: string) {
  // Note: This is conceptual as direct Power Query interaction 
  // has limitations in Office Scripts
  
  // Get the sheet containing the query results
  const sheet = workbook.getWorksheet(queryName);
  if (!sheet) {
    throw new Error(`Query result sheet '${queryName}' not found`);
  }
  
  // Get the data range
  const dataRange = sheet.getUsedRange();
  const values = dataRange.getValues();
  
  // Extract headers and data
  const headers = values[0] as string[];
  const data = values.slice(1);
  
  // Process the data
  // ...
  
  return {
    tableName: queryName,
    columns: headers,
    rows: data
  };
}
```

### Excel Automation Best Practices

#### Performance Optimization

```typescript
function optimizedDataProcessing(workbook: ExcelScript.Workbook) {
  // BAD: Multiple individual cell operations
  const sheet = workbook.getActiveWorksheet();
  for (let i = 0; i < 100; i++) {
    for (let j = 0; j < 5; j++) {
      sheet.getCell(i, j).setValue(i * j);
    }
  }
  
  // GOOD: Batch operations with a 2D array
  const values: number[][] = [];
  for (let i = 0; i < 100; i++) {
    const row: number[] = [];
    for (let j = 0; j < 5; j++) {
      row.push(i * j);
    }
    values.push(row);
  }
  sheet.getRange("A1:E100").setValues(values);
}
```

#### Error Handling Patterns

```typescript
function robustFunction(workbook: ExcelScript.Workbook) {
  try {
    // Main functionality
    const sheet = workbook.getActiveWorksheet();
    
    // Check preconditions
    const dataRange = sheet.getUsedRange();
    if (!dataRange) {
      throw new Error("No data found in worksheet");
    }
    
    const values = dataRange.getValues();
    if (values.length <= 1) {
      throw new Error("Insufficient data rows");
    }
    
    // Check for required columns
    const headers = values[0] as string[];
    const requiredColumns = ["Date", "Amount", "Category"];
    for (const column of requiredColumns) {
      if (headers.indexOf(column) === -1) {
        throw new Error(`Required column '${column}' is missing`);
      }
    }
    
    // Process data
    // ...
    
    return {
      success: true,
      message: "Processing completed successfully"
    };
  } catch (error) {
    // Create error log in the workbook
    try {
      const errorSheet = workbook.getWorksheet("Error Log") || workbook.addWorksheet("Error Log");
      
      // Get next row
      const lastRow = errorSheet.getUsedRange()?.getRowCount() || 0;
      
      // Log error
      errorSheet.getRange(`A${lastRow + 1}:C${lastRow + 1}`).setValues([
        [new Date().toISOString(), "ERROR", error.message]
      ]);
    } catch {
      // Fallback if even error logging fails
    }
    
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}
```

