# Excel Charts

Create and manipulate charts.

## Create a Chart

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const dataRange = sheet.getRange("A1:B5");
const chart = sheet.charts.add(
  Excel.ChartType.columnClustered,
  dataRange,
  Excel.ChartSeriesBy.auto
);
chart.name = "SalesChart";
await context.sync();
return "Chart created";
```

## Chart Types

Common chart types:
- `Excel.ChartType.columnClustered`
- `Excel.ChartType.columnStacked`
- `Excel.ChartType.barClustered`
- `Excel.ChartType.line`
- `Excel.ChartType.lineMarkers`
- `Excel.ChartType.pie`
- `Excel.ChartType.doughnut`
- `Excel.ChartType.area`
- `Excel.ChartType.scatter`
- `Excel.ChartType.bubble`

## Set Chart Title

```javascript
const chart = sheet.charts.getItem("SalesChart");
chart.title.text = "Monthly Sales";
chart.title.visible = true;
await context.sync();
```

## Position and Size

```javascript
const chart = sheet.charts.getItem("SalesChart");
chart.left = 300;    // Points from left
chart.top = 100;     // Points from top
chart.width = 400;
chart.height = 300;
await context.sync();
```

## Chart Legend

```javascript
const chart = sheet.charts.getItem("SalesChart");
chart.legend.visible = true;
chart.legend.position = Excel.ChartLegendPosition.right;
await context.sync();
```

## Axes

```javascript
const chart = sheet.charts.getItem("SalesChart");
const valueAxis = chart.axes.valueAxis;
const categoryAxis = chart.axes.categoryAxis;

valueAxis.title.text = "Revenue ($)";
valueAxis.minimum = 0;
valueAxis.maximum = 1000;

categoryAxis.title.text = "Month";

await context.sync();
```

## Data Labels

```javascript
const chart = sheet.charts.getItem("SalesChart");
const series = chart.series.getItemAt(0);
series.hasDataLabels = true;
series.dataLabels.showValue = true;
series.dataLabels.position = Excel.ChartDataLabelPosition.outsideEnd;
await context.sync();
```

## List All Charts

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const charts = sheet.charts;
charts.load("items/name,items/chartType");
await context.sync();
return charts.items.map(c => ({ name: c.name, type: c.chartType }));
```

## Delete a Chart

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const chart = sheet.charts.getItem("SalesChart");
chart.delete();
await context.sync();
return "Chart deleted";
```

## Update Chart Data

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const chart = sheet.charts.getItem("SalesChart");
chart.setData(sheet.getRange("A1:B10"), Excel.ChartSeriesBy.auto);
await context.sync();
return "Chart data updated";
```

## Export Chart as Image

```javascript
const chart = sheet.charts.getItem("SalesChart");
const image = chart.getImage(
  Excel.ImageFittingMode.fit,
  400,  // width
  300   // height
);
await context.sync();
return image.value;  // Base64 PNG
```

## Tips

- Charts are named automatically if not specified
- Use `charts.getItemAt(index)` for unnamed charts
- Position uses points (72 points = 1 inch)
- Image export returns base64-encoded PNG
