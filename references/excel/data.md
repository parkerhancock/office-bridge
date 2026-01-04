# Excel Data Operations

Data validation, import/export, and data manipulation.

## Data Validation

### Dropdown List

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1:A10");
range.dataValidation.rule = {
  list: {
    inCellDropDown: true,
    source: "Option1,Option2,Option3"
  }
};
await context.sync();
return "Dropdown added";
```

### From Range

```javascript
const range = sheet.getRange("B1:B10");
range.dataValidation.rule = {
  list: {
    inCellDropDown: true,
    source: "=Sheet2!$A$1:$A$5"  // Reference another range
  }
};
await context.sync();
```

### Numeric Validation

```javascript
const range = sheet.getRange("C1:C10");
range.dataValidation.rule = {
  wholeNumber: {
    formula1: 0,
    formula2: 100,
    operator: Excel.DataValidationOperator.between
  }
};
range.dataValidation.errorAlert = {
  message: "Value must be between 0 and 100",
  showAlert: true,
  style: Excel.DataValidationAlertStyle.stop,
  title: "Invalid Input"
};
await context.sync();
```

### Date Validation

```javascript
const range = sheet.getRange("D1:D10");
range.dataValidation.rule = {
  date: {
    formula1: "2024-01-01",
    formula2: "2024-12-31",
    operator: Excel.DataValidationOperator.between
  }
};
await context.sync();
```

### Clear Validation

```javascript
const range = sheet.getRange("A1:D10");
range.dataValidation.clear();
await context.sync();
```

## Copy and Paste

### Copy Range

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const sourceRange = sheet.getRange("A1:B5");
const destRange = sheet.getRange("D1:E5");
destRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
await context.sync();
return "Copied";
```

### Copy Options

```javascript
// Values only (no formulas)
destRange.copyFrom(sourceRange, Excel.RangeCopyType.values);

// Formulas only
destRange.copyFrom(sourceRange, Excel.RangeCopyType.formulas);

// Formats only
destRange.copyFrom(sourceRange, Excel.RangeCopyType.formats);
```

## Find and Replace

### Find Values

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const usedRange = sheet.getUsedRange();
usedRange.load("values,address");
await context.sync();

const searchTerm = "error";
const matches = [];
usedRange.values.forEach((row, rowIndex) => {
  row.forEach((cell, colIndex) => {
    if (String(cell).toLowerCase().includes(searchTerm)) {
      matches.push({ row: rowIndex, col: colIndex, value: cell });
    }
  });
});
return matches;
```

### Replace Values

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1:Z100");
range.load("values");
await context.sync();

const oldValue = "old";
const newValue = "new";
const newValues = range.values.map(row =>
  row.map(cell =>
    typeof cell === "string" ? cell.replace(oldValue, newValue) : cell
  )
);
range.values = newValues;
await context.sync();
return "Replaced";
```

## Clear Operations

```javascript
const range = sheet.getRange("A1:Z100");

// Clear everything
range.clear(Excel.ClearApplyTo.all);

// Clear values only (keep formatting)
range.clear(Excel.ClearApplyTo.contents);

// Clear formatting only (keep values)
range.clear(Excel.ClearApplyTo.formats);

await context.sync();
```

## Insert and Delete

### Insert Rows

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("3:5");  // Rows 3-5
range.insert(Excel.InsertShiftDirection.down);
await context.sync();
return "Rows inserted";
```

### Insert Columns

```javascript
const range = sheet.getRange("B:C");  // Columns B-C
range.insert(Excel.InsertShiftDirection.right);
await context.sync();
```

### Delete Rows

```javascript
const range = sheet.getRange("3:5");
range.delete(Excel.DeleteShiftDirection.up);
await context.sync();
```

## Special Cells

### Get Blank Cells

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const usedRange = sheet.getUsedRange();
const blanks = usedRange.getSpecialCells(Excel.SpecialCellType.blanks);
blanks.load("address");
await context.sync();
return blanks.address;
```

### Get Cells with Formulas

```javascript
const formulas = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
formulas.load("address");
await context.sync();
return formulas.address;
```

### Get Cells with Constants

```javascript
const constants = usedRange.getSpecialCells(Excel.SpecialCellType.constants);
constants.load("address");
await context.sync();
return constants.address;
```

## Freeze Panes

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();

// Freeze first row
sheet.freezePanes.freezeRows(1);

// Freeze first column
sheet.freezePanes.freezeColumns(1);

// Freeze at specific cell
sheet.freezePanes.freezeAt(sheet.getRange("B2"));

// Unfreeze
sheet.freezePanes.unfreeze();

await context.sync();
```

## Tips

- Data validation messages support basic formatting
- Use `Excel.ClearApplyTo` enum for targeted clearing
- Special cells may throw if no matching cells found - wrap in try/catch
- Freeze panes affects the view, not the data
