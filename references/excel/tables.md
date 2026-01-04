# Excel Tables

Structured tables with headers, filtering, and sorting.

## Create a Table

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const table = sheet.tables.add("A1:C10", true);  // true = has headers
table.name = "SalesTable";
await context.sync();
return "Table created";
```

## Get Existing Table

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const table = sheet.tables.getItem("SalesTable");
table.load("name,showHeaders,showTotals");
await context.sync();
return { name: table.name, headers: table.showHeaders, totals: table.showTotals };
```

## List All Tables

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const tables = sheet.tables;
tables.load("items/name,items/showHeaders");
await context.sync();
return tables.items.map(t => t.name);
```

## Table Columns

### Get Column Data

```javascript
const table = sheet.tables.getItem("SalesTable");
const column = table.columns.getItem("Revenue");
const dataRange = column.getDataBodyRange();
dataRange.load("values");
await context.sync();
return dataRange.values.flat();
```

### Add Column

```javascript
const table = sheet.tables.getItem("SalesTable");
table.columns.add(null, [
  ["Status"],
  ["Active"],
  ["Pending"],
  ["Complete"]
]);
await context.sync();
return "Column added";
```

### Delete Column

```javascript
const table = sheet.tables.getItem("SalesTable");
const column = table.columns.getItem("OldColumn");
column.delete();
await context.sync();
```

## Table Rows

### Get All Rows

```javascript
const table = sheet.tables.getItem("SalesTable");
const bodyRange = table.getDataBodyRange();
bodyRange.load("values");
await context.sync();
return bodyRange.values;
```

### Add Row

```javascript
const table = sheet.tables.getItem("SalesTable");
table.rows.add(null, [["New Item", 100, "Active"]]);
await context.sync();
return "Row added";
```

### Delete Row

```javascript
const table = sheet.tables.getItem("SalesTable");
const row = table.rows.getItemAt(0);  // First data row
row.delete();
await context.sync();
```

## Sorting

```javascript
const table = sheet.tables.getItem("SalesTable");
table.sort.apply([
  { key: 0, ascending: true },   // First column ascending
  { key: 1, ascending: false }   // Second column descending
]);
await context.sync();
return "Table sorted";
```

## Filtering

### Apply Filter

```javascript
const table = sheet.tables.getItem("SalesTable");
const column = table.columns.getItem("Status");
column.filter.applyValuesFilter(["Active", "Pending"]);
await context.sync();
return "Filter applied";
```

### Filter Types

```javascript
// Values filter
column.filter.applyValuesFilter(["Value1", "Value2"]);

// Top/Bottom filter
column.filter.applyTopItemsFilter(10);  // Top 10

// Custom filter
column.filter.applyCustomFilter(">100");

// Date filter
column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.thisMonth);
```

### Clear Filters

```javascript
const table = sheet.tables.getItem("SalesTable");
table.clearFilters();
await context.sync();
```

## Total Row

```javascript
const table = sheet.tables.getItem("SalesTable");
table.showTotals = true;

const revenueColumn = table.columns.getItem("Revenue");
revenueColumn.getTotalRowRange().load("values");
await context.sync();

// Set total function
revenueColumn.getTotalRowRange().formulas = [["=SUBTOTAL(109,[Revenue])"]];
await context.sync();
```

## Table Styles

```javascript
const table = sheet.tables.getItem("SalesTable");
table.style = "TableStyleMedium2";
table.showBandedRows = true;
table.showBandedColumns = false;
table.showFilterButton = true;
await context.sync();
```

## Convert to Range

```javascript
const table = sheet.tables.getItem("SalesTable");
const range = table.convertToRange();
await context.sync();
return "Converted to range";
```

## Delete Table

```javascript
const table = sheet.tables.getItem("SalesTable");
table.delete();
await context.sync();
return "Table deleted";
```

## Tips

- Table names must be unique in the workbook
- Column names come from the header row
- `getDataBodyRange()` excludes headers and totals
- Filter criteria are case-insensitive
- Styles: `TableStyleLight1-21`, `TableStyleMedium1-28`, `TableStyleDark1-11`
