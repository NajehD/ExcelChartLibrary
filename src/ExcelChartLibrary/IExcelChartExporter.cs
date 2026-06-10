using ExcelChartLibrary.Structures;
using OutSystems.ExternalLibraries.SDK;

namespace ExcelChartLibrary;

/// <summary>
/// ODC external library that generates Excel (.xlsx) files containing a data
/// table and a native Excel chart built from that data.
/// </summary>
[OSInterface(
    Name = "ExcelChartExporter",
    Description = "Generates Excel (.xlsx) files containing a data table and a native Excel chart (Column, Bar, Line, Pie, Area or Scatter) built from the input data. Powered by NPOI (Apache 2.0).")]
public interface IExcelChartExporter
{
    /// <summary>
    /// Builds an .xlsx workbook with the provided data written as a table and
    /// a chart of the requested type plotted next to it.
    /// </summary>
    [OSAction(
        Description = "Creates an Excel file with the input data written as a table and a chart of the requested type plotted next to it. Returns the .xlsx file as binary data.",
        ReturnName = "ExcelFile",
        ReturnDescription = "The generated .xlsx file as binary data.",
        ReturnType = OSDataType.BinaryData)]
    byte[] ExportChartToExcel(
        [OSParameter(Description = "Chart type. One of: Column, Bar, Line, Pie, Area, Scatter (case-insensitive).")]
        string chartType,
        [OSParameter(Description = "Category labels (X axis). Each series must have one value per category.")]
        List<string> categories,
        [OSParameter(Description = "Data series to plot. Pie charts use only the first series.")]
        List<ChartSeries> series,
        [OSParameter(Description = "Title displayed above the chart. Optional.")]
        string chartTitle = "",
        [OSParameter(Description = "Name of the worksheet. Defaults to 'Chart Data'.")]
        string sheetName = "Chart Data",
        [OSParameter(Description = "Title of the category (X) axis. Optional, ignored for Pie charts.")]
        string categoryAxisTitle = "",
        [OSParameter(Description = "Title of the value (Y) axis. Optional, ignored for Pie charts.")]
        string valueAxisTitle = "",
        [OSParameter(Description = "Whether to show the chart legend. Defaults to True.")]
        bool showLegend = true);

    /// <summary>
    /// Lists the chart type names accepted by ExportChartToExcel.
    /// </summary>
    [OSAction(
        Description = "Returns the list of chart type names accepted by ExportChartToExcel.",
        ReturnName = "ChartTypes",
        ReturnDescription = "Supported chart type names.")]
    List<string> GetSupportedChartTypes();
}
