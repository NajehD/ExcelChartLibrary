using System.IO.Compression;
using System.Text;
using ExcelChartLibrary;
using ExcelChartLibrary.Structures;
using NPOI.XSSF.UserModel;
using Xunit;

namespace ExcelChartLibrary.Tests;

public class ExcelChartExporterTests
{
    private readonly ExcelChartExporter _exporter = new();

    private static List<string> Categories => ["Q1", "Q2", "Q3", "Q4"];

    private static List<ChartSeries> Series =>
    [
        new ChartSeries { Name = "Revenue", Values = [120.5m, 150m, 90.25m, 200m] },
        new ChartSeries { Name = "Costs", Values = [80m, 95.5m, 70m, 110m] }
    ];

    [Theory]
    [InlineData("Column", "barChart")]
    [InlineData("Bar", "barChart")]
    [InlineData("Line", "lineChart")]
    [InlineData("Pie", "pieChart")]
    [InlineData("Doughnut", "doughnutChart")]
    [InlineData("Area", "areaChart")]
    [InlineData("Scatter", "scatterChart")]
    [InlineData("Gantt", "barChart")]
    public void ExportChartToExcel_GeneratesWorkbookWithChart(string chartType, string expectedChartElement)
    {
        var bytes = _exporter.ExportChartToExcel(
            chartType, Categories, Series,
            chartTitle: $"{chartType} Test",
            sheetName: "Report",
            categoryAxisTitle: "Quarter",
            valueAxisTitle: "Amount");

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0, "Generated file is empty.");

        var chartXml = ReadChartXml(bytes);
        Assert.Contains($"<c:{expectedChartElement}>", chartXml);

        // The workbook must be readable and contain the data table.
        using var stream = new MemoryStream(bytes);
        var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.GetSheet("Report");
        Assert.NotNull(sheet);
        Assert.Equal("Revenue", sheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal(120.5, sheet.GetRow(1).GetCell(1).NumericCellValue, 3);
    }

    [Fact]
    public void ExportChartToExcel_IsCaseInsensitiveOnChartType()
    {
        var bytes = _exporter.ExportChartToExcel("pIe", Categories, Series);
        Assert.Contains("<c:pieChart>", ReadChartXml(bytes));
    }

    [Theory]
    [InlineData("Donut", "doughnutChart")]
    [InlineData("donut chart", "doughnutChart")]
    [InlineData("GanttChart", "barChart")]
    [InlineData("Gantt Chart", "barChart")]
    public void ExportChartToExcel_AcceptsChartTypeSynonyms(string chartType, string expectedChartElement)
    {
        var bytes = _exporter.ExportChartToExcel(chartType, Categories, Series);
        Assert.Contains($"<c:{expectedChartElement}>", ReadChartXml(bytes));
    }

    [Fact]
    public void ExportChartToExcel_Doughnut_HasHoleAndAllSeriesAsRings()
    {
        var bytes = _exporter.ExportChartToExcel("Doughnut", Categories, Series);
        var chartXml = ReadChartXml(bytes);
        Assert.Contains("<c:holeSize", chartXml);
        Assert.DoesNotContain("<c:pieChart>", chartXml);
        // Both series present as rings, with fully qualified name references.
        Assert.Contains("$B$1", chartXml);
        Assert.Contains("$C$1", chartXml);
        Assert.Equal(2, chartXml.Split("<c:ser>").Length - 1);
    }

    [Fact]
    public void ExportChartToExcel_Gantt_IsStackedWithHiddenFirstSeries()
    {
        var bytes = _exporter.ExportChartToExcel("Gantt",
            ["Design", "Build", "Test"],
            [
                new ChartSeries { Name = "Start", Values = [0m, 5m, 12m] },
                new ChartSeries { Name = "Duration", Values = [5m, 7m, 4m] }
            ]);
        var chartXml = ReadChartXml(bytes);
        Assert.Contains("grouping val=\"stacked\"", chartXml);
        Assert.Contains("overlap val=\"100\"", chartXml);
        Assert.Contains("<a:noFill", chartXml);
        Assert.Contains("orientation val=\"maxMin\"", chartXml);
        Assert.Contains("barDir val=\"bar\"", chartXml);
    }

    [Fact]
    public void ExportChartToExcel_Gantt_RequiresTwoSeries()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _exporter.ExportChartToExcel("Gantt", Categories,
                [new ChartSeries { Name = "Only", Values = [1m, 2m, 3m, 4m] }]));
        Assert.Contains("at least two series", ex.Message);
    }

    [Fact]
    public void ExportChartToExcel_Heatmap_AppliesColorScaleInsteadOfChart()
    {
        var bytes = _exporter.ExportChartToExcel("Heatmap", Categories, Series,
            chartTitle: "Intensity", sheetName: "Report");

        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        Assert.DoesNotContain(archive.Entries, e => e.FullName.StartsWith("xl/charts/"));
        using var reader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var sheetXml = reader.ReadToEnd();
        Assert.Contains("colorScale", sheetXml);
        Assert.Contains("FFF8696B", sheetXml);

        // Title row shifts the table down by one row.
        using var stream = new MemoryStream(bytes);
        var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.GetSheet("Report");
        Assert.Equal("Intensity", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Revenue", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal(120.5, sheet.GetRow(2).GetCell(1).NumericCellValue, 3);
    }

    [Fact]
    public void ExportChartToExcel_Heatmap_WithoutTitleStartsAtFirstRow()
    {
        var bytes = _exporter.ExportChartToExcel("Heatmap", Categories, Series, sheetName: "Report");
        using var stream = new MemoryStream(bytes);
        var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.GetSheet("Report");
        Assert.Equal("Revenue", sheet.GetRow(0).GetCell(1).StringCellValue);
    }

    [Fact]
    public void ExportChartToExcel_WritesAxisTitles()
    {
        var bytes = _exporter.ExportChartToExcel("Line", Categories, Series,
            categoryAxisTitle: "Quarter", valueAxisTitle: "Amount (EUR)");
        var chartXml = ReadChartXml(bytes);
        Assert.Contains("Quarter", chartXml);
        Assert.Contains("Amount (EUR)", chartXml);
    }

    [Fact]
    public void ExportChartToExcel_SanitizesInvalidSheetName()
    {
        var bytes = _exporter.ExportChartToExcel("Column", Categories, Series,
            sheetName: "Sales/2026:Q1*");
        using var stream = new MemoryStream(bytes);
        var workbook = new XSSFWorkbook(stream);
        Assert.NotNull(workbook.GetSheet("Sales 2026 Q1"));
    }

    [Fact]
    public void ExportChartToExcel_RejectsUnknownChartType()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _exporter.ExportChartToExcel("Bubble", Categories, Series));
        Assert.Contains("Bubble", ex.Message);
        Assert.Contains("Column", ex.Message);
    }

    [Fact]
    public void ExportChartToExcel_RejectsMismatchedSeriesLength()
    {
        var badSeries = new List<ChartSeries>
        {
            new ChartSeries { Name = "Short", Values = [1m, 2m] }
        };
        var ex = Assert.Throws<ArgumentException>(() =>
            _exporter.ExportChartToExcel("Column", Categories, badSeries));
        Assert.Contains("Short", ex.Message);
    }

    [Fact]
    public void ExportChartToExcel_RejectsEmptyInput()
    {
        Assert.Throws<ArgumentException>(() => _exporter.ExportChartToExcel("Column", [], Series));
        Assert.Throws<ArgumentException>(() => _exporter.ExportChartToExcel("Column", Categories, []));
    }

    [Fact]
    public void GetSupportedChartTypes_ReturnsAllTypes()
    {
        var types = _exporter.GetSupportedChartTypes();
        Assert.Equal(["Column", "Bar", "Line", "Pie", "Doughnut", "Area", "Scatter", "Heatmap", "Gantt"], types);
    }

    private static string ReadChartXml(byte[] xlsxBytes)
    {
        using var archive = new ZipArchive(new MemoryStream(xlsxBytes), ZipArchiveMode.Read);
        var entry = archive.Entries.FirstOrDefault(e => e.FullName.StartsWith("xl/charts/chart"));
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
