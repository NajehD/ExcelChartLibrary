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
    [InlineData("Area", "areaChart")]
    [InlineData("Scatter", "scatterChart")]
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
        Assert.Equal(["Column", "Bar", "Line", "Pie", "Area", "Scatter"], types);
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
