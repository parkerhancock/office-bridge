# Excel Formatting

Cell formatting, styles, and conditional formatting.

## Basic Formatting

### Font Properties

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1");
range.format.font.bold = true;
range.format.font.italic = true;
range.format.font.size = 14;
range.format.font.color = "#FF0000";
range.format.font.name = "Arial";
range.format.font.underline = Excel.RangeUnderlineStyle.single;
await context.sync();
```

### Fill Color

```javascript
const range = sheet.getRange("A1:B5");
range.format.fill.color = "#FFFF00";  // Yellow background
await context.sync();
```

### Borders

```javascript
const range = sheet.getRange("A1:C5");
const border = range.format.borders.getItem(Excel.BorderIndex.edgeBottom);
border.style = Excel.BorderLineStyle.continuous;
border.color = "#000000";
border.weight = Excel.BorderWeight.medium;
await context.sync();
```

### All Borders

```javascript
const range = sheet.getRange("A1:C5");
const borders = [
  Excel.BorderIndex.edgeTop,
  Excel.BorderIndex.edgeBottom,
  Excel.BorderIndex.edgeLeft,
  Excel.BorderIndex.edgeRight,
  Excel.BorderIndex.insideHorizontal,
  Excel.BorderIndex.insideVertical
];
for (const index of borders) {
  const border = range.format.borders.getItem(index);
  border.style = Excel.BorderLineStyle.continuous;
  border.weight = Excel.BorderWeight.thin;
}
await context.sync();
```

### Alignment

```javascript
const range = sheet.getRange("A1");
range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
range.format.verticalAlignment = Excel.VerticalAlignment.center;
range.format.wrapText = true;
await context.sync();
```

## Number Formats

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();

// Currency
sheet.getRange("A1").numberFormat = [["$#,##0.00"]];

// Percentage
sheet.getRange("A2").numberFormat = [["0.00%"]];

// Date
sheet.getRange("A3").numberFormat = [["yyyy-mm-dd"]];

// Custom
sheet.getRange("A4").numberFormat = [["#,##0.00;[Red]-#,##0.00"]];

await context.sync();
```

## Column Width and Row Height

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();

// Set column width
sheet.getRange("A:A").format.columnWidth = 150;

// Set row height
sheet.getRange("1:1").format.rowHeight = 30;

// Auto-fit column
sheet.getRange("B:B").format.autofitColumns();

// Auto-fit rows
sheet.getRange("A1:A10").format.autofitRows();

await context.sync();
```

## Merge Cells

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1:C1");
range.merge(true);  // true = merge across
await context.sync();
```

## Conditional Formatting

### Color Scale

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("A1:A10");
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
cf.colorScale.threeColorScale.minimum = {
  color: "#FF0000",
  type: Excel.ConditionalFormatColorCriterionType.lowestValue
};
cf.colorScale.threeColorScale.midpoint = {
  color: "#FFFF00",
  type: Excel.ConditionalFormatColorCriterionType.percentile,
  formula: "50"
};
cf.colorScale.threeColorScale.maximum = {
  color: "#00FF00",
  type: Excel.ConditionalFormatColorCriterionType.highestValue
};
await context.sync();
```

### Data Bars

```javascript
const range = sheet.getRange("B1:B10");
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
cf.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
cf.dataBar.positiveFormat.fillColor = "#0066CC";
await context.sync();
```

### Cell Value Rule

```javascript
const range = sheet.getRange("C1:C10");
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
cf.cellValue.format.fill.color = "#FFCCCC";
cf.cellValue.format.font.color = "#CC0000";
cf.cellValue.rule = {
  formula1: "=50",
  operator: Excel.ConditionalCellValueOperator.greaterThan
};
await context.sync();
```

### Icon Sets

```javascript
const range = sheet.getRange("D1:D10");
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
cf.iconSet.style = Excel.IconSet.threeArrows;
await context.sync();
```

## Clear Formatting

```javascript
const range = sheet.getRange("A1:Z100");
range.format.fill.clear();
range.format.font.bold = false;
range.format.font.italic = false;
range.format.font.color = "#000000";
range.conditionalFormats.clearAll();
await context.sync();
```

## Tips

- Colors use hex format: `"#RRGGBB"`
- Number formats follow Excel format codes
- Conditional formats stack (first match wins)
- Use `autofitColumns()` after setting content
