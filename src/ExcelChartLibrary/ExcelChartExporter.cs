using ExcelChartLibrary.Structures;

namespace ExcelChartLibrary;

/// <summary>
/// Implementation of the ExcelChartExporter external library.
/// </summary>
public class ExcelChartExporter : IExcelChartExporter
{
    public byte[] ExportChartToExcel(
        string chartType,
        List<string> categories,
        List<ChartSeries> series,
        string chartTitle = "",
        string sheetName = "Chart Data",
        string categoryAxisTitle = "",
        string valueAxisTitle = "",
        bool showLegend = true)
    {
        var kind = ChartBuilder.ParseChartType(chartType);
        Validate(categories, series);

        return ChartBuilder.Build(
            kind,
            categories,
            series,
            chartTitle ?? string.Empty,
            sheetName ?? string.Empty,
            categoryAxisTitle ?? string.Empty,
            valueAxisTitle ?? string.Empty,
            showLegend);
    }

    public List<string> GetSupportedChartTypes() => ChartBuilder.SupportedChartTypes.ToList();

    private static void Validate(List<string> categories, List<ChartSeries> series)
    {
        if (categories == null || categories.Count == 0)
            throw new ArgumentException("At least one category is required.", nameof(categories));

        if (series == null || series.Count == 0)
            throw new ArgumentException("At least one data series is required.", nameof(series));

        for (var i = 0; i < series.Count; i++)
        {
            var values = series[i].Values;
            if (values == null || values.Count != categories.Count)
            {
                var name = string.IsNullOrWhiteSpace(series[i].Name) ? $"#{i + 1}" : $"'{series[i].Name}'";
                throw new ArgumentException(
                    $"Series {name} has {values?.Count ?? 0} value(s) but there are {categories.Count} categories. " +
                    "Each series must have exactly one value per category.",
                    nameof(series));
            }
        }
    }
}
