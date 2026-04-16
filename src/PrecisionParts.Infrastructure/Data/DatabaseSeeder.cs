using PrecisionParts.Core.Enums;
using PrecisionParts.Core.Models;

namespace PrecisionParts.Infrastructure.Data;

/// <summary>
/// Database seeder — replaces VB6 seed_data.sql with type-safe C# seed data.
/// Provides realistic manufacturing data for development and testing.
/// </summary>
public static class DatabaseSeeder
{
    public static async Task SeedAsync(PrecisionPartsDbContext db)
    {
        if (db.Users.Any()) return; // Already seeded

        // Users (passwords are hashed — replaces VB6 plain-text storage)
        var users = new[]
        {
            new User { UserId = 1, Username = "admin", PasswordHash = "hashed_admin_2024", Role = UserRole.Administrator, FullName = "System Administrator", Email = "admin@precisionparts.com" },
            new User { UserId = 2, Username = "jsmith", PasswordHash = "hashed_jsmith_2024", Role = UserRole.Manager, FullName = "John Smith", Email = "jsmith@precisionparts.com" },
            new User { UserId = 3, Username = "mgarcia", PasswordHash = "hashed_mgarcia_2024", Role = UserRole.Operator, FullName = "Maria Garcia", Email = "mgarcia@precisionparts.com" },
            new User { UserId = 4, Username = "rwilson", PasswordHash = "hashed_rwilson_2024", Role = UserRole.QualityControl, FullName = "Robert Wilson", Email = "rwilson@precisionparts.com" },
            new User { UserId = 5, Username = "tchang", PasswordHash = "hashed_tchang_2024", Role = UserRole.Shipping, FullName = "Tony Chang", Email = "tchang@precisionparts.com" },
        };
        db.Users.AddRange(users);

        // Customers
        var customers = new[]
        {
            new Customer { CustomerId = 1, CustomerCode = "ACME", CompanyName = "Acme Manufacturing Co.", ContactName = "Jane Doe", Phone = "555-0101", Email = "orders@acme.com", Address = "123 Industrial Blvd", City = "Detroit", State = "MI", ZipCode = "48201" },
            new Customer { CustomerId = 2, CustomerCode = "BOLT", CompanyName = "BoltTech Industries", ContactName = "Bob Builder", Phone = "555-0102", Email = "purchasing@bolttech.com", Address = "456 Factory Way", City = "Cleveland", State = "OH", ZipCode = "44101" },
            new Customer { CustomerId = 3, CustomerCode = "PREC", CompanyName = "Precision Dynamics", ContactName = "Alice Wright", Phone = "555-0103", Email = "orders@precdyn.com", Address = "789 Engineering Dr", City = "Pittsburgh", State = "PA", ZipCode = "15201" },
        };
        db.Customers.AddRange(customers);

        // Parts
        var parts = new[]
        {
            new Part { PartId = 1, PartNumber = "PP-1001-A", Description = "Steel Mounting Bracket", Material = "Steel", UnitCost = 12.50m, Weight = 0.75, Category = "Brackets" },
            new Part { PartId = 2, PartNumber = "PP-1002-B", Description = "Aluminum Housing Assembly", Material = "Aluminum", UnitCost = 45.00m, Weight = 2.3, Category = "Housings" },
            new Part { PartId = 3, PartNumber = "PP-2001-A", Description = "Precision Drive Shaft", Material = "Hardened Steel", UnitCost = 89.00m, Weight = 5.1, Category = "Shafts" },
            new Part { PartId = 4, PartNumber = "MFG-3001-C", Description = "Custom Gear Assembly", Material = "Tool Steel", UnitCost = 125.00m, Weight = 3.2, Category = "Gears" },
            new Part { PartId = 5, PartNumber = "HDW-4001-A", Description = "Stainless Steel Fastener Kit", Material = "Stainless Steel", UnitCost = 8.75m, Weight = 0.3, Category = "Hardware" },
        };
        db.Parts.AddRange(parts);

        // Raw Materials
        var materials = new[]
        {
            new RawMaterial { MaterialId = 1, MaterialCode = "STL-1018", Description = "1018 Cold Rolled Steel Bar", Supplier = "US Steel Supply", QuantityOnHand = 500, ReorderPoint = 100, UnitCost = 3.50m, UnitOfMeasure = "LB", Location = "Warehouse A" },
            new RawMaterial { MaterialId = 2, MaterialCode = "ALM-6061", Description = "6061-T6 Aluminum Sheet", Supplier = "Alcoa Metals", QuantityOnHand = 250, ReorderPoint = 50, UnitCost = 8.25m, UnitOfMeasure = "LB", Location = "Warehouse A" },
            new RawMaterial { MaterialId = 3, MaterialCode = "STL-4140", Description = "4140 Alloy Steel Round", Supplier = "US Steel Supply", QuantityOnHand = 75, ReorderPoint = 100, UnitCost = 5.75m, UnitOfMeasure = "LB", Location = "Warehouse B" },
        };
        db.RawMaterials.AddRange(materials);

        // Work Orders
        var workOrders = new[]
        {
            new WorkOrder { WorkOrderId = 1, WorkOrderNumber = "WO-20240301-001", PartNumber = "PP-1001-A", Quantity = 100, QuantityCompleted = 100, Status = WorkOrderStatus.Completed, Priority = WorkOrderPriority.Normal, CustomerId = 1, DueDate = DateTime.UtcNow.AddDays(-10), StartDate = DateTime.UtcNow.AddDays(-20), CompletedDate = DateTime.UtcNow.AddDays(-5), CreatedBy = "jsmith" },
            new WorkOrder { WorkOrderId = 2, WorkOrderNumber = "WO-20240315-001", PartNumber = "PP-2001-A", Quantity = 50, QuantityCompleted = 30, Status = WorkOrderStatus.InProgress, Priority = WorkOrderPriority.High, CustomerId = 2, DueDate = DateTime.UtcNow.AddDays(5), StartDate = DateTime.UtcNow.AddDays(-7), CreatedBy = "jsmith" },
            new WorkOrder { WorkOrderId = 3, WorkOrderNumber = "WO-20240320-001", PartNumber = "MFG-3001-C", Quantity = 25, QuantityCompleted = 0, Status = WorkOrderStatus.New, Priority = WorkOrderPriority.Urgent, CustomerId = 3, DueDate = DateTime.UtcNow.AddDays(14), CreatedBy = "admin" },
        };
        db.WorkOrders.AddRange(workOrders);

        await db.SaveChangesAsync();
    }
}
