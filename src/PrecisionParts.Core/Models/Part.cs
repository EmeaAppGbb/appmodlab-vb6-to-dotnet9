namespace PrecisionParts.Core.Models;

/// <summary>
/// Part master data — migrated from VB6 clsPart.cls and Parts table.
/// Format: PP-XXXX-NNN with category prefix validation.
/// </summary>
public class Part
{
    public int PartId { get; set; }
    public string PartNumber { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string? Material { get; set; }
    public string? MaterialSpec { get; set; }
    public decimal UnitCost { get; set; }
    public double Weight { get; set; }
    public string? DrawingNumber { get; set; }
    public string RevisionLevel { get; set; } = "A";
    public string? Category { get; set; }
    public bool IsActive { get; set; } = true;
    public DateTime CreatedDate { get; set; } = DateTime.UtcNow;
    public DateTime? ModifiedDate { get; set; }

    // Navigation properties
    public ICollection<WorkOrder> WorkOrders { get; set; } = new List<WorkOrder>();
}
