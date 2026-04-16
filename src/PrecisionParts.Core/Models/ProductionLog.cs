namespace PrecisionParts.Core.Models;

/// <summary>
/// Production log entry — migrated from VB6 ProductionLog table.
/// Tracks per-shift production output, rejects, and downtime.
/// </summary>
public class ProductionLog
{
    public int LogId { get; set; }
    public int WorkOrderId { get; set; }
    public DateTime LogDate { get; set; } = DateTime.UtcNow;
    public string? Shift { get; set; }
    public string? Operator { get; set; }
    public string? MachineId { get; set; }
    public double QuantityProduced { get; set; }
    public double QuantityRejected { get; set; }
    public int DowntimeMinutes { get; set; }
    public string? Notes { get; set; }

    // Navigation properties
    public WorkOrder WorkOrder { get; set; } = null!;
    public MachineStatus? Machine { get; set; }
}
