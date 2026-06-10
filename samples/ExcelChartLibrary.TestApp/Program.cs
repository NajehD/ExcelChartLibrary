using ExcelChartLibrary;
using ExcelChartLibrary.Structures;

var builder = WebApplication.CreateBuilder(args);

// The exact same class ODC instantiates for the external library.
builder.Services.AddSingleton<IExcelChartExporter, ExcelChartExporter>();

var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapGet("/api/chart-types", (IExcelChartExporter exporter) =>
    Results.Ok(exporter.GetSupportedChartTypes()));

app.MapPost("/api/export", (ExportRequest request, IExcelChartExporter exporter) =>
{
    try
    {
        var series = request.Series
            .Select(s => new ChartSeries { Name = s.Name, Values = s.Values })
            .ToList();

        var bytes = exporter.ExportChartToExcel(
            request.ChartType,
            request.Categories,
            series,
            request.ChartTitle ?? string.Empty,
            string.IsNullOrWhiteSpace(request.SheetName) ? "Chart Data" : request.SheetName,
            request.CategoryAxisTitle ?? string.Empty,
            request.ValueAxisTitle ?? string.Empty,
            request.ShowLegend ?? true);

        var fileName = $"{(string.IsNullOrWhiteSpace(request.ChartTitle) ? request.ChartType : request.ChartTitle)}.xlsx";
        return Results.File(bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            fileName);
    }
    catch (ArgumentException ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
});

// Quick one-click download with built-in demo data, e.g. /api/sample/Pie
app.MapGet("/api/sample/{chartType}", (string chartType, IExcelChartExporter exporter) =>
{
    try
    {
        var isGantt = chartType.Replace(" ", "").Replace("-", "")
            .Contains("gantt", StringComparison.OrdinalIgnoreCase);

        var bytes = isGantt
            ? exporter.ExportChartToExcel(chartType,
                ["Design", "Build", "Test", "Deploy"],
                [
                    new ChartSeries { Name = "Start", Values = [0m, 5m, 12m, 16m] },
                    new ChartSeries { Name = "Duration", Values = [5m, 7m, 4m, 2m] }
                ],
                chartTitle: "Project plan", sheetName: "Plan",
                categoryAxisTitle: "Task", valueAxisTitle: "Day")
            : exporter.ExportChartToExcel(chartType,
                ["Q1", "Q2", "Q3", "Q4"],
                [
                    new ChartSeries { Name = "Revenue", Values = [120.5m, 150m, 90.25m, 200m] },
                    new ChartSeries { Name = "Costs", Values = [80m, 95.5m, 70m, 110m] }
                ],
                chartTitle: $"Quarterly results ({chartType})", sheetName: "Report",
                categoryAxisTitle: "Quarter", valueAxisTitle: "Amount (kEUR)");

        return Results.File(bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            $"Sample_{chartType}.xlsx");
    }
    catch (ArgumentException ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
});

app.Run();

record ExportRequest(
    string ChartType,
    List<string> Categories,
    List<SeriesDto> Series,
    string? ChartTitle,
    string? SheetName,
    string? CategoryAxisTitle,
    string? ValueAxisTitle,
    bool? ShowLegend);

record SeriesDto(string Name, List<decimal> Values);
