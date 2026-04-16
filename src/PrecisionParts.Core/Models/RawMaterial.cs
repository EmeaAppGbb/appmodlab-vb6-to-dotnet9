using PrecisionParts.Core.Enums;

namespace PrecisionParts.Core.Models;

/// <summary>
/// Raw material inventory — migrated from VB6 clsInventory.cls and RawMaterials table.
/// Includes reorder point logic previously hardcoded in VB6.
/// </summary>
public class RawMaterial
{
    public int MaterialId { get; set; }
    public string MaterialCode { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string? Supplier { get; set; }
    public double QuantityOnHand { get; set; }
    public double QuantityOnOrder { get; set; }
    public double QuantityReserved { get; set; }
    public double ReorderPoint { get; set; } = 100;
    public double ReorderQuantity { get; set; } = 500;
    public decimal UnitCost { get; set; }
    public string UnitOfMeasure { get; set; } = "EA";
    public string? Location { get; set; }
    public DateTime? LastOrderDate { get; set; }
    public DateTime? LastReceivedDate { get; set; }
    public int LeadTimeDays { get; set; } = 7;
    public bool IsActive { get; set; } = true;
    public string? Notes { get; set; }
    public DateTime CreatedDate { get; set; } = DateTime.UtcNow;
    public DateTime? ModifiedDate { get; set; }

    // Navigation properties
    public ICollection<InventoryTransaction> Transactions { get; set; } = new List<InventoryTransaction>();

    /// <summary>
    /// Available = OnHand + OnOrder - Reserved
    /// Replaces VB6 clsInventory.GetAvailableQuantity
    /// </summary>
    public double AvailableQuantity => Math.Max(0, QuantityOnHand - QuantityReserved);

    /// <summary>
    /// Total inventory value = qty × unit cost
    /// Replaces VB6 clsInventory.GetTotalValue
    /// </summary>
    public decimal TotalValue => (decimal)QuantityOnHand * UnitCost;

    /// <summary>
    /// Replaces VB6 clsInventory.CheckReorderStatus
    /// </summary>
    public StockStatus GetStockStatus()
    {
        var available = QuantityOnHand + QuantityOnOrder - QuantityReserved;
        if (QuantityOnHand <= 0) return StockStatus.Critical;
        if (available <= ReorderPoint) return StockStatus.ReorderRequired;
        if (available <= ReorderPoint * 1.5) return StockStatus.LowStock;
        return StockStatus.Ok;
    }

    public bool NeedsReorder()
    {
        var status = GetStockStatus();
        return status is StockStatus.ReorderRequired or StockStatus.Critical;
    }

    public DateTime GetExpectedDeliveryDate()
    {
        var baseDate = LastOrderDate ?? DateTime.UtcNow;
        return baseDate.AddDays(LeadTimeDays);
    }
}
