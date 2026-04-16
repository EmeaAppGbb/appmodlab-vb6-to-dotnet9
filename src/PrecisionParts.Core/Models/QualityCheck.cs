using PrecisionParts.Core.Enums;

namespace PrecisionParts.Core.Models;

/// <summary>
/// Quality check inspection — migrated from VB6 QualityChecks table.
/// Business rule: Fail result auto-sets work order to OnHold.
/// </summary>
public class QualityCheck
{
    public int CheckId { get; set; }
    public int WorkOrderId { get; set; }
    public string Inspector { get; set; } = string.Empty;
    public DateTime CheckDate { get; set; } = DateTime.UtcNow;
    public QualityCheckResult Result { get; set; } = QualityCheckResult.Pass;
    public string? Measurements { get; set; }
    public string? Notes { get; set; }
    public int DefectCount { get; set; }
    public int? SampleSize { get; set; }
    public double? Temperature { get; set; }
    public double? Humidity { get; set; }
    public string? CreatedBy { get; set; }
    public DateTime CreatedDate { get; set; } = DateTime.UtcNow;

    // Navigation properties
    public WorkOrder WorkOrder { get; set; } = null!;
}
