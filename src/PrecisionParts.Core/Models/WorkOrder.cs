using PrecisionParts.Core.Enums;

namespace PrecisionParts.Core.Models;

/// <summary>
/// Work order — migrated from VB6 clsWorkOrder.cls and WorkOrders table.
/// Cost calculation logic extracted from hardcoded VB6 values into configurable settings.
/// </summary>
public class WorkOrder
{
    public int WorkOrderId { get; set; }
    public string WorkOrderNumber { get; set; } = string.Empty;
    public string PartNumber { get; set; } = string.Empty;
    public double Quantity { get; set; }
    public double QuantityCompleted { get; set; }
    public WorkOrderStatus Status { get; set; } = WorkOrderStatus.New;
    public WorkOrderPriority Priority { get; set; } = WorkOrderPriority.Normal;
    public int? CustomerId { get; set; }
    public string? CustomerPO { get; set; }
    public DateTime? DueDate { get; set; }
    public DateTime? StartDate { get; set; }
    public DateTime? CompletedDate { get; set; }
    public string? AssignedTo { get; set; }
    public decimal? EstimatedCost { get; set; }
    public decimal? ActualCost { get; set; }
    public string? Notes { get; set; }
    public string? CreatedBy { get; set; }
    public DateTime CreatedDate { get; set; } = DateTime.UtcNow;
    public DateTime? ModifiedDate { get; set; }

    // Navigation properties
    public Customer? Customer { get; set; }
    public Part? Part { get; set; }
    public ICollection<QualityCheck> QualityChecks { get; set; } = new List<QualityCheck>();
    public ICollection<ShippingManifest> ShippingManifests { get; set; } = new List<ShippingManifest>();
    public ICollection<ProductionLog> ProductionLogs { get; set; } = new List<ProductionLog>();

    /// <summary>
    /// Replaces VB6 clsWorkOrder.GetPercentComplete
    /// </summary>
    public double PercentComplete =>
        Quantity > 0 ? (QuantityCompleted / Quantity) * 100 : 0;

    /// <summary>
    /// Replaces VB6 clsWorkOrder.IsOverdue
    /// </summary>
    public bool IsOverdue =>
        Status is not (WorkOrderStatus.Completed or WorkOrderStatus.Cancelled)
        && DueDate.HasValue && DateTime.UtcNow > DueDate.Value;

    /// <summary>
    /// Replaces VB6 clsWorkOrder.GetDaysRemaining
    /// </summary>
    public int DaysRemaining =>
        DueDate.HasValue ? (DueDate.Value - DateTime.UtcNow).Days : 0;
}
