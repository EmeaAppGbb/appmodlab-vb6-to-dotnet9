namespace PrecisionParts.Core.Models;

/// <summary>
/// Customer record — migrated from VB6 Customers table.
/// </summary>
public class Customer
{
    public int CustomerId { get; set; }
    public string CustomerCode { get; set; } = string.Empty;
    public string CompanyName { get; set; } = string.Empty;
    public string? ContactName { get; set; }
    public string? Phone { get; set; }
    public string? Email { get; set; }
    public string? Address { get; set; }
    public string? City { get; set; }
    public string? State { get; set; }
    public string? ZipCode { get; set; }
    public string Country { get; set; } = "USA";
    public string PaymentTerms { get; set; } = "Net 30";
    public bool IsActive { get; set; } = true;

    // Navigation properties
    public ICollection<WorkOrder> WorkOrders { get; set; } = new List<WorkOrder>();
}
