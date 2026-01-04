# Excel Formulas

Work with formulas, functions, and named ranges.

## Setting Formulas

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const cell = sheet.getRange("C1");
cell.formulas = [["=A1+B1"]];
await context.sync();
return "Formula set";
```

## Getting Formulas

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1:C5");
range.load("formulas");
await context.sync();
return range.formulas;  // 2D array of formula strings
```

## Common Functions

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();

// SUM
sheet.getRange("A10").formulas = [["=SUM(A1:A9)"]];

// AVERAGE
sheet.getRange("B10").formulas = [["=AVERAGE(B1:B9)"]];

// VLOOKUP
sheet.getRange("D1").formulas = [["=VLOOKUP(C1,A1:B9,2,FALSE)"]];

// IF
sheet.getRange("E1").formulas = [["=IF(A1>100,\"High\",\"Low\")"]];

// COUNTIF
sheet.getRange("F1").formulas = [["=COUNTIF(A1:A9,\">50\")"]];

await context.sync();
```

## Named Ranges

### Create Named Range

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1:B10");
context.workbook.names.add("SalesData", range);
await context.sync();
return "Named range created";
```

### Use Named Range

```javascript
const namedItem = context.workbook.names.getItem("SalesData");
const range = namedItem.getRange();
range.load("values");
await context.sync();
return range.values;
```

### List Named Ranges

```javascript
const names = context.workbook.names;
names.load("items/name,items/type");
await context.sync();
return names.items.map(n => ({ name: n.name, type: n.type }));
```

### Delete Named Range

```javascript
const namedItem = context.workbook.names.getItem("SalesData");
namedItem.delete();
await context.sync();
return "Named range deleted";
```

## Array Formulas

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("D1:D10");
// Dynamic array formula (Excel 365)
range.formulas = [["=SORT(A1:A10)"]];
await context.sync();
```

## R1C1 Reference Style

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const cell = sheet.getRange("C1");
cell.formulasR1C1 = [["=RC[-2]+RC[-1]"]];  // Same as =A1+B1
await context.sync();
```

## Calculate Workbook

```javascript
// Recalculate all formulas
context.workbook.application.calculate(Excel.CalculationType.full);
await context.sync();
return "Calculated";
```

## Tips

- Formulas must start with `=`
- Use double quotes inside formulas: `"=IF(A1>0,\"Yes\",\"No\")"`
- Formulas are locale-independent (use commas, not semicolons)
- Array formulas return to a spill range in Excel 365
