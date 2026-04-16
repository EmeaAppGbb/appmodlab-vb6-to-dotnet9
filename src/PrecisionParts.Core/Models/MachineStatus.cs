namespace PrecisionParts.Core.Models;

/// <summary>
/// Machine status for real-time monitoring — migrated from VB6 MachineStatus table.
/// </summary>
public class MachineStatus
{
    public string MachineId { get; set; } = string.Empty;
    public string MachineName { get; set; } = string.Empty;
    public string Status { get; set; } = "Idle";
    public int? CurrentWorkOrderId { get; set; }
    public string? CurrentOperator { get; set; }
    public double? LastCycleTime { get; set; }
    public long TotalCyclesCompleted { get; set; }
    public DateTime? LastMaintenanceDate { get; set; }
    public DateTime? NextMaintenanceDate { get; set; }
    public string? Location { get; set; }
    public bool IsActive { get; set; } = true;

    // Navigation properties
    public WorkOrder? CurrentWorkOrder { get; set; }
}
