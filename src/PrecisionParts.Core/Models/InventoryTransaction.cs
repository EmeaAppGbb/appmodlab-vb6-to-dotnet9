namespace PrecisionParts.Core.Models;

/// <summary>
/// Inventory transaction log — replaces VB6 clsInventory.LogInventoryTransaction.
/// Every stock change is recorded for audit trail.
/// </summary>
public class InventoryTransaction
{
    public int TransactionId { get; set; }
    public int MaterialId { get; set; }
    public double QuantityChange { get; set; }
    public double QuantityAfter { get; set; }
    public string Reason { get; set; } = string.Empty;
    public DateTime TransactionDate { get; set; } = DateTime.UtcNow;
    public int? UserId { get; set; }

    // Navigation properties
    public RawMaterial Material { get; set; } = null!;
}
