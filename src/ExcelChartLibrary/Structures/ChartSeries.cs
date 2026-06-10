using OutSystems.ExternalLibraries.SDK;

namespace ExcelChartLibrary.Structures;

/// <summary>
/// One named data series to plot on the chart.
/// </summary>
[OSStructure(Description = "A named data series to plot on the chart. The number of Values must match the number of Categories.")]
public struct ChartSeries
{
    [OSStructureField(Description = "Name of the series, shown in the chart legend and as the column header in the data sheet.", IsMandatory = true)]
    public string Name;

    [OSStructureField(Description = "Numeric values of the series, one per category.", IsMandatory = true)]
    public List<decimal> Values;
}
