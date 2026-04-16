namespace PrecisionParts.Core.Models;

/// <summary>
/// Shipping manifest — migrated from VB6 ShippingManifests table.
/// Business rule: shipping auto-completes the work order.
/// </summary>
public class ShippingManifest
{
    public int ManifestId { get; set; }
    public string ManifestNumber { get; set; } = string.Empty;
    public int WorkOrderId { get; set; }
    public DateTime ShipDate { get; set; } = DateTime.UtcNow;
    public string? Carrier { get; set; }
    public string? TrackingNumber { get; set; }
    public string? ShipToName { get; set; }
    public string? ShipToAddress { get; set; }
    public string? ShipToCity { get; set; }
    public string? ShipToState { get; set; }
    public string? ShipToZipCode { get; set; }
    public double? Weight { get; set; }
    public int Boxes { get; set; } = 1;
    public decimal? FreightCost { get; set; }
    public string? Notes { get; set; }
    public string? CreatedBy { get; set; }
    public DateTime CreatedDate { get; set; } = DateTime.UtcNow;

    // Navigation properties
    public WorkOrder WorkOrder { get; set; } = null!;
}
