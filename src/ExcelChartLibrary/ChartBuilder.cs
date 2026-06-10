using ExcelChartLibrary.Structures;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ExcelChartLibrary;

internal enum ChartKind
{
    Column,
    Bar,
    Line,
    Pie,
    Doughnut,
    Area,
    Scatter,
    Heatmap,
    Gantt
}

/// <summary>
/// Writes the data table into a worksheet and plots a native Excel chart
/// referencing those cells, using NPOI.
/// </summary>
internal static class ChartBuilder
{
    public static readonly string[] SupportedChartTypes =
        ["Column", "Bar", "Line", "Pie", "Doughnut", "Area", "Scatter", "Heatmap", "Gantt"];

    private const int ChartWidthInColumns = 9;
    private const int ChartHeightInRows = 22;
    private const int MaxSheetNameLength = 31;

    public static ChartKind ParseChartType(string chartType)
    {
        // Accept common synonyms and ignore spaces/hyphens/underscores,
        // e.g. "donut", "Gantt Chart", "heat-map".
        var normalized = (chartType ?? string.Empty)
            .Replace(" ", string.Empty)
            .Replace("-", string.Empty)
            .Replace("_", string.Empty);

        normalized = normalized.ToLowerInvariant() switch
        {
            "donut" or "donutchart" or "doughnutchart" => "Doughnut",
            "ganttchart" => "Gantt",
            "heatmapchart" => "Heatmap",
            _ => normalized
        };

        if (Enum.TryParse<ChartKind>(normalized, ignoreCase: true, out var kind))
            return kind;

        throw new ArgumentException(
            $"Unknown chart type '{chartType}'. Supported types: {string.Join(", ", SupportedChartTypes)}.",
            nameof(chartType));
    }

    public static byte[] Build(
        ChartKind kind,
        List<string> categories,
        List<ChartSeries> series,
        string chartTitle,
        string sheetName,
        string categoryAxisTitle,
        string valueAxisTitle,
        bool showLegend)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = (XSSFSheet)workbook.CreateSheet(SanitizeSheetName(sheetName));

        if (kind == ChartKind.Heatmap)
        {
            BuildHeatmap(workbook, sheet, categories, series, chartTitle, categoryAxisTitle);
        }
        else
        {
            WriteDataTable(workbook, sheet, kind, categories, series, categoryAxisTitle, startRow: 0);
            var chart = CreateChartFrame(sheet, series.Count, chartTitle, showLegend);
            PlotChart(chart, sheet, kind, categories.Count, series.Count, categoryAxisTitle, valueAxisTitle);
        }

        using var stream = new MemoryStream();
        workbook.Write(stream, leaveOpen: true);
        return stream.ToArray();
    }

    /// <summary>
    /// Lays the data out as: header row with the category axis title and the
    /// series names, then one row per category. Scatter charts get a numeric
    /// category column so it can be used as the X value range.
    /// </summary>
    private static void WriteDataTable(
        XSSFWorkbook workbook,
        XSSFSheet sheet,
        ChartKind kind,
        List<string> categories,
        List<ChartSeries> series,
        string categoryAxisTitle,
        int startRow)
    {
        var headerStyle = workbook.CreateCellStyle();
        var headerFont = workbook.CreateFont();
        headerFont.IsBold = true;
        headerStyle.SetFont(headerFont);

        var headerRow = sheet.CreateRow(startRow);
        var categoryHeader = headerRow.CreateCell(0);
        categoryHeader.SetCellValue(string.IsNullOrWhiteSpace(categoryAxisTitle) ? "Category" : categoryAxisTitle);
        categoryHeader.CellStyle = headerStyle;

        for (var s = 0; s < series.Count; s++)
        {
            var cell = headerRow.CreateCell(s + 1);
            cell.SetCellValue(SeriesName(series[s], s));
            cell.CellStyle = headerStyle;
        }

        var numericCategories = kind == ChartKind.Scatter ? ToNumericCategories(categories) : null;

        for (var r = 0; r < categories.Count; r++)
        {
            var row = sheet.CreateRow(startRow + r + 1);
            var categoryCell = row.CreateCell(0);
            if (numericCategories != null)
                categoryCell.SetCellValue(numericCategories[r]);
            else
                categoryCell.SetCellValue(categories[r] ?? string.Empty);

            for (var s = 0; s < series.Count; s++)
                row.CreateCell(s + 1).SetCellValue((double)series[s].Values[r]);
        }

        for (var c = 0; c <= series.Count; c++)
            sheet.SetColumnWidth(c, 14 * 256);
    }

    private static XSSFChart CreateChartFrame(XSSFSheet sheet, int seriesCount, string chartTitle, bool showLegend)
    {
        var drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
        var firstColumn = seriesCount + 2;
        var anchor = drawing.CreateAnchor(0, 0, 0, 0, firstColumn, 0, firstColumn + ChartWidthInColumns, ChartHeightInRows);
        var chart = (XSSFChart)drawing.CreateChart(anchor);

        if (!string.IsNullOrWhiteSpace(chartTitle))
            chart.SetTitle(chartTitle);

        if (showLegend)
        {
            var legend = chart.GetOrCreateLegend();
            legend.Position = LegendPosition.Bottom;
        }

        return chart;
    }

    private static void PlotChart(
        XSSFChart chart,
        XSSFSheet sheet,
        ChartKind kind,
        int categoryCount,
        int seriesCount,
        string categoryAxisTitle,
        string valueAxisTitle)
    {
        var firstDataRow = 1;
        var lastDataRow = categoryCount;
        var categoryRange = new CellRangeAddress(firstDataRow, lastDataRow, 0, 0);

        if (kind is ChartKind.Pie or ChartKind.Doughnut)
        {
            // Excel pie charts plot a single series; doughnut charts render
            // each additional series as a concentric ring.
            var pieData = chart.ChartDataFactory.CreatePieChartData<string, double>();
            var categorySource = DataSources.FromStringCellRange(sheet, categoryRange);
            var pieSeriesCount = kind == ChartKind.Pie ? 1 : seriesCount;
            for (var s = 0; s < pieSeriesCount; s++)
            {
                var valueSource = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(firstDataRow, lastDataRow, s + 1, s + 1));
                var pieSeries = pieData.AddSeries(categorySource, valueSource);
                pieSeries.SetTitle(HeaderCellReference(sheet, s));
            }
            chart.Plot(pieData);

            if (kind == ChartKind.Doughnut)
                ConvertPieToDoughnut(chart);
            return;
        }

        var bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
        var leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
        leftAxis.Crosses = AxisCrosses.AutoZero;

        switch (kind)
        {
            case ChartKind.Column:
            {
                var data = chart.ChartDataFactory.CreateColumnChartData<string, double>();
                AddCategorySeries(data.AddSeries, sheet, categoryRange, firstDataRow, lastDataRow, seriesCount);
                chart.Plot(data, bottomAxis, leftAxis);
                break;
            }
            case ChartKind.Bar:
            case ChartKind.Gantt:
            {
                var data = chart.ChartDataFactory.CreateBarChartData<string, double>();
                AddCategorySeries(data.AddSeries, sheet, categoryRange, firstDataRow, lastDataRow, seriesCount);
                chart.Plot(data, bottomAxis, leftAxis);
                if (kind == ChartKind.Gantt)
                    ApplyGanttFormatting(chart);
                break;
            }
            case ChartKind.Line:
            {
                var data = chart.ChartDataFactory.CreateLineChartData<string, double>();
                AddCategorySeries(data.AddSeries, sheet, categoryRange, firstDataRow, lastDataRow, seriesCount);
                chart.Plot(data, bottomAxis, leftAxis);
                break;
            }
            case ChartKind.Area:
            {
                var data = chart.ChartDataFactory.CreateAreaChartData<string, double>();
                AddCategorySeries(data.AddSeries, sheet, categoryRange, firstDataRow, lastDataRow, seriesCount);
                chart.Plot(data, bottomAxis, leftAxis);
                break;
            }
            case ChartKind.Scatter:
            {
                var data = chart.ChartDataFactory.CreateScatterChartData<double, double>();
                var xSource = DataSources.FromNumericCellRange(sheet, categoryRange);
                for (var s = 0; s < seriesCount; s++)
                {
                    var ySource = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(firstDataRow, lastDataRow, s + 1, s + 1));
                    var scatterSeries = data.AddSeries(xSource, ySource);
                    scatterSeries.SetTitle(HeaderCellReference(sheet, s));
                }
                chart.Plot(data, bottomAxis, leftAxis);
                break;
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported chart type.");
        }

        SetAxisTitles(chart, categoryAxisTitle, valueAxisTitle);
    }

    /// <summary>
    /// NPOI has no doughnut chart factory, so the plotted pie chart is
    /// rewritten as a doughnutChart element in the underlying OOXML part.
    /// </summary>
    private static void ConvertPieToDoughnut(XSSFChart chart)
    {
        var plotArea = chart.GetCTChart().plotArea;
        var pie = plotArea.pieChart[0];

        var doughnut = new CT_DoughnutChart
        {
            varyColors = pie.varyColors,
            ser = pie.ser,
            holeSize = new CT_HoleSize { val = 50 }
        };

        plotArea.doughnutChart ??= [];
        plotArea.doughnutChart.Add(doughnut);
        plotArea.pieChart.Clear();
    }

    /// <summary>
    /// Turns the plotted horizontal bar chart into the classic Excel Gantt:
    /// stacked bars where the first series (start offsets) is invisible, the
    /// remaining series are the phase durations, and tasks read top-down.
    /// </summary>
    private static void ApplyGanttFormatting(XSSFChart chart)
    {
        var ctChart = chart.GetCTChart();
        var bar = ctChart.plotArea.barChart[0];

        bar.grouping ??= bar.AddNewGrouping();
        bar.grouping.val = ST_BarGrouping.stacked;
        bar.overlap = new CT_Overlap { val = 100 };

        bar.ser[0].spPr = new CT_ShapeProperties { noFill = new NPOI.OpenXmlFormats.Dml.CT_NoFillProperties() };

        // First task at the top, like a project plan.
        var scaling = ctChart.plotArea.catAx[0].scaling;
        scaling.orientation ??= scaling.AddNewOrientation();
        scaling.orientation.val = ST_Orientation.maxMin;

        // Note: the offset series still gets a legend entry. NPOI 2.7.4 cannot
        // serialize <c:legendEntry><c:delete/> correctly, so hiding it would
        // produce invalid chart XML.
    }

    /// <summary>
    /// Excel has no native heatmap chart type: the standard representation is
    /// the data grid itself with a 3-color scale conditional format (red →
    /// yellow → green), which is what Excel produces for a heatmap as well.
    /// </summary>
    private static void BuildHeatmap(
        XSSFWorkbook workbook,
        XSSFSheet sheet,
        List<string> categories,
        List<ChartSeries> series,
        string chartTitle,
        string categoryAxisTitle)
    {
        var startRow = 0;
        if (!string.IsNullOrWhiteSpace(chartTitle))
        {
            var titleFont = workbook.CreateFont();
            titleFont.IsBold = true;
            titleFont.FontHeightInPoints = 14;
            var titleStyle = workbook.CreateCellStyle();
            titleStyle.SetFont(titleFont);

            var titleCell = sheet.CreateRow(0).CreateCell(0);
            titleCell.SetCellValue(chartTitle);
            titleCell.CellStyle = titleStyle;
            if (series.Count > 0)
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, series.Count));
            startRow = 1;
        }

        WriteDataTable(workbook, sheet, ChartKind.Heatmap, categories, series, categoryAxisTitle, startRow);

        var dataRegion = new CellRangeAddress(startRow + 1, startRow + categories.Count, 1, series.Count);
        var formatting = sheet.SheetConditionalFormatting;
        var rule = formatting.CreateConditionalFormattingColorScaleRule();
        var colorScale = rule.ColorScaleFormatting;

        var low = colorScale.CreateThreshold();
        low.RangeType = RangeType.MIN;
        var mid = colorScale.CreateThreshold();
        mid.RangeType = RangeType.PERCENTILE;
        mid.Value = 50;
        var high = colorScale.CreateThreshold();
        high.RangeType = RangeType.MAX;
        colorScale.Thresholds = [low, mid, high];

        // Excel's default red → yellow → green scale.
        colorScale.Colors =
        [
            CreateColor("FFF8696B"),
            CreateColor("FFFFEB84"),
            CreateColor("FF63BE7B")
        ];

        formatting.AddConditionalFormatting([dataRegion], rule);

        // NPOI stamps dxfId="0" on the rule, but color scale rules carry no
        // differential style and the workbook defines none, so strict readers
        // reject the reference. Remove it.
        foreach (var conditionalFormatting in sheet.GetCTWorksheet().conditionalFormatting)
        foreach (var cfRule in conditionalFormatting.cfRule)
            cfRule.dxfIdSpecified = false;
    }

    private static XSSFColor CreateColor(string argbHex)
    {
        var color = new XSSFColor();
        color.ARGBHex = argbHex;
        return color;
    }

    private delegate IChartSeries AddSeriesFunc(IChartDataSource<string> categories, IChartDataSource<double> values);

    private static void AddCategorySeries(
        AddSeriesFunc addSeries,
        XSSFSheet sheet,
        CellRangeAddress categoryRange,
        int firstDataRow,
        int lastDataRow,
        int seriesCount)
    {
        var categorySource = DataSources.FromStringCellRange(sheet, categoryRange);
        for (var s = 0; s < seriesCount; s++)
        {
            var valueSource = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(firstDataRow, lastDataRow, s + 1, s + 1));
            var chartSeries = addSeries(categorySource, valueSource);
            chartSeries.SetTitle(HeaderCellReference(sheet, s));
        }
    }

    /// <summary>
    /// Fully qualified reference to a series header cell (e.g. 'Chart
    /// Data'!$B$1). A bare reference like "B1" cannot be resolved by Excel
    /// from within the chart part.
    /// </summary>
    private static CellReference HeaderCellReference(XSSFSheet sheet, int seriesIndex) =>
        new(sheet.SheetName, 0, seriesIndex + 1, true, true);

    /// <summary>
    /// NPOI's classic chart API has no axis title support, so the titles are
    /// written directly into the underlying OOXML chart parts.
    /// </summary>
    private static void SetAxisTitles(XSSFChart chart, string categoryAxisTitle, string valueAxisTitle)
    {
        var plotArea = chart.GetCTChart().plotArea;

        if (!string.IsNullOrWhiteSpace(categoryAxisTitle) && plotArea.catAx is { Count: > 0 })
            plotArea.catAx[0].title = CreateAxisTitle(categoryAxisTitle);

        if (!string.IsNullOrWhiteSpace(valueAxisTitle) && plotArea.valAx is { Count: > 0 })
            plotArea.valAx[0].title = CreateAxisTitle(valueAxisTitle);
    }

    private static CT_Title CreateAxisTitle(string text)
    {
        var title = new CT_Title();
        var rich = title.AddNewTx().AddNewRich();
        rich.AddNewBodyPr();
        rich.AddNewP().AddNewR().t = text;
        title.overlay = new CT_Boolean { val = 0 };
        return title;
    }

    /// <summary>
    /// Scatter charts need numeric X values. Category labels that parse as
    /// numbers are used as-is; otherwise the 1-based category index is used.
    /// </summary>
    private static double[] ToNumericCategories(List<string> categories)
    {
        var result = new double[categories.Count];
        for (var i = 0; i < categories.Count; i++)
        {
            result[i] = double.TryParse(categories[i], System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var value)
                ? value
                : i + 1;
        }
        return result;
    }

    private static string SeriesName(ChartSeries series, int index) =>
        string.IsNullOrWhiteSpace(series.Name) ? $"Series {index + 1}" : series.Name;

    private static string SanitizeSheetName(string sheetName)
    {
        var name = string.IsNullOrWhiteSpace(sheetName) ? "Chart Data" : sheetName;
        foreach (var invalid in new[] { ':', '\\', '/', '?', '*', '[', ']' })
            name = name.Replace(invalid, ' ');
        name = name.Trim();
        if (name.Length == 0)
            name = "Chart Data";
        return name.Length > MaxSheetNameLength ? name[..MaxSheetNameLength] : name;
    }
}
