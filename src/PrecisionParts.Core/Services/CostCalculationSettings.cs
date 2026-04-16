namespace PrecisionParts.Core.Services;

/// <summary>
/// Application settings — replaces VB6 modGlobals.bas hardcoded constants.
/// Loaded from appsettings.json via IOptions pattern.
/// </summary>
public class CostCalculationSettings
{
    public const string SectionName = "CostCalculation";

    /// <summary>Labor rate per hour (VB6 hardcoded $45/hr)</summary>
    public decimal LaborRatePerHour { get; set; } = 45.00m;

    /// <summary>Estimated hours per part (VB6 hardcoded 0.5)</summary>
    public decimal HoursPerPart { get; set; } = 0.5m;

    /// <summary>Overhead percentage (VB6 hardcoded 15%)</summary>
    public decimal OverheadPercentage { get; set; } = 0.15m;
}
