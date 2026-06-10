# ExcelChartExporter — ODC External Library

An [OutSystems Developer Cloud (ODC) external library](https://success.outsystems.com/documentation/outsystems_developer_cloud/building_apps/extend_your_apps_with_external_logic/)
that generates Excel (`.xlsx`) files containing a data table **and a native Excel chart**
built from the input data.

It is built on [NPOI](https://github.com/nissl-lab/npoi) **2.7.4**, which is
**free under the Apache 2.0 license** (no commercial licensing required — unlike
EPPlus 5+, Syncfusion, or NPOI 2.8.0+ which introduced the OSMF maintenance-fee EULA).

## Supported chart types

| Chart type | Notes |
|------------|-------|
| `Column`   | Vertical bars |
| `Bar`      | Horizontal bars |
| `Line`     | One line per series |
| `Pie`      | Uses the **first series only** (Excel pie charts plot a single series) |
| `Area`     | One area per series |
| `Scatter`  | Category labels that parse as numbers are used as X values; otherwise the 1-based index is used |

## Exposed server actions

### `ExportChartToExcel`

Writes the data as a table in a worksheet and plots a chart of the requested
type next to it. Returns the `.xlsx` file as **Binary Data** — ready to feed
into a Download node, attach to an email, or store.

| Parameter | Type | Mandatory | Description |
|-----------|------|-----------|-------------|
| `ChartType` | Text | Yes | One of `Column`, `Bar`, `Line`, `Pie`, `Area`, `Scatter` (case-insensitive) |
| `Categories` | List of Text | Yes | Category labels (X axis) |
| `Series` | List of `ChartSeries` | Yes | The data series to plot |
| `ChartTitle` | Text | No | Title shown above the chart |
| `SheetName` | Text | No | Worksheet name (default `Chart Data`; invalid characters are replaced) |
| `CategoryAxisTitle` | Text | No | X-axis title (ignored for Pie) |
| `ValueAxisTitle` | Text | No | Y-axis title (ignored for Pie) |
| `ShowLegend` | Boolean | No | Show the chart legend (default `True`) |

Returns: `ExcelFile` (Binary Data).

### `GetSupportedChartTypes`

Returns the list of accepted chart type names — useful to populate a dropdown.

### `ChartSeries` structure

| Attribute | Type | Description |
|-----------|------|-------------|
| `Name` | Text | Series name (legend entry / column header) |
| `Values` | List of Decimal | One value per category |

Every series must have exactly one value per category; the library validates
this and raises a descriptive error otherwise.

## Using it in ODC

1. Build the upload package:
   ```bash
   ./scripts/package.sh
   ```
   This produces `artifacts/ExcelChartExporter.zip`. (The GitHub Actions
   workflow also publishes this zip as a build artifact on every push.)
2. In the **ODC Portal**, go to **External logic** and upload the zip.
3. In **ODC Studio**, add the `ExcelChartExporter` server actions to your app
   (Add public elements → search for `ExportChartToExcel`).
4. Build the `Categories` and `Series` lists from your data (e.g. an
   aggregate result), call `ExportChartToExcel`, and pass the returned binary
   to a **Download** node with a `.xlsx` filename and MIME type
   `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.

## Example

For categories `["Q1","Q2","Q3","Q4"]` and two series (`Revenue`, `Costs`),
calling `ExportChartToExcel("Column", ...)` produces a worksheet like:

| Quarter | Revenue | Costs |
|---------|---------|-------|
| Q1 | 120.5 | 80 |
| Q2 | 150 | 95.5 |
| Q3 | 90.25 | 70 |
| Q4 | 200 | 110 |

with a native, editable Excel column chart plotted beside the table. Because
the chart references the cells (not pasted as an image), users can restyle or
edit it in Excel.

## Development

```bash
dotnet test            # run the test suite
./scripts/package.sh   # produce the ODC upload zip
```

Requirements: .NET 8 SDK.

### Why NPOI?

| Library | Charts | License / cost |
|---------|--------|----------------|
| **NPOI 2.7.4** ✅ | Column, Bar, Line, Pie, Area, Scatter | Apache 2.0 — free, incl. commercial use |
| NPOI ≥ 2.8.0 | Same + more | Requires OSMF maintenance-fee EULA for commercial use |
| EPPlus ≥ 5 | Full | Polyform Noncommercial — paid license for commercial use |
| ClosedXML | ❌ cannot create charts | MIT |
| Open XML SDK | Possible but very low-level | MIT |
| Syncfusion | Full | Commercial (excluded) |
