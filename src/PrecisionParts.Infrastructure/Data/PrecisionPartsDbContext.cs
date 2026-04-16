using Microsoft.EntityFrameworkCore;
using PrecisionParts.Core.Models;

namespace PrecisionParts.Infrastructure.Data;

/// <summary>
/// EF Core DbContext — replaces VB6 modDatabase.bas global ADO connection.
/// Scoped per request instead of a single global connection.
/// </summary>
public class PrecisionPartsDbContext : DbContext
{
    public PrecisionPartsDbContext(DbContextOptions<PrecisionPartsDbContext> options)
        : base(options) { }

    public DbSet<User> Users => Set<User>();
    public DbSet<Part> Parts => Set<Part>();
    public DbSet<RawMaterial> RawMaterials => Set<RawMaterial>();
    public DbSet<Customer> Customers => Set<Customer>();
    public DbSet<WorkOrder> WorkOrders => Set<WorkOrder>();
    public DbSet<QualityCheck> QualityChecks => Set<QualityCheck>();
    public DbSet<ShippingManifest> ShippingManifests => Set<ShippingManifest>();
    public DbSet<ProductionLog> ProductionLogs => Set<ProductionLog>();
    public DbSet<MachineStatus> MachineStatuses => Set<MachineStatus>();
    public DbSet<ShiftSchedule> ShiftSchedules => Set<ShiftSchedule>();
    public DbSet<InventoryTransaction> InventoryTransactions => Set<InventoryTransaction>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // User
        modelBuilder.Entity<User>(e =>
        {
            e.HasKey(u => u.UserId);
            e.Property(u => u.Username).HasMaxLength(50).IsRequired();
            e.Property(u => u.PasswordHash).HasMaxLength(256).IsRequired();
            e.Property(u => u.Role).HasConversion<string>().HasMaxLength(50);
            e.Property(u => u.FullName).HasMaxLength(100);
            e.Property(u => u.Email).HasMaxLength(100);
            e.HasIndex(u => u.Username).IsUnique();
        });

        // Part
        modelBuilder.Entity<Part>(e =>
        {
            e.HasKey(p => p.PartId);
            e.Property(p => p.PartNumber).HasMaxLength(50).IsRequired();
            e.Property(p => p.Description).HasMaxLength(255).IsRequired();
            e.Property(p => p.Material).HasMaxLength(100);
            e.Property(p => p.MaterialSpec).HasMaxLength(100);
            e.Property(p => p.UnitCost).HasColumnType("decimal(18,2)");
            e.Property(p => p.DrawingNumber).HasMaxLength(50);
            e.Property(p => p.RevisionLevel).HasMaxLength(10);
            e.Property(p => p.Category).HasMaxLength(50);
            e.HasIndex(p => p.PartNumber).IsUnique();
        });

        // RawMaterial
        modelBuilder.Entity<RawMaterial>(e =>
        {
            e.HasKey(r => r.MaterialId);
            e.Property(r => r.MaterialCode).HasMaxLength(50).IsRequired();
            e.Property(r => r.Description).HasMaxLength(255).IsRequired();
            e.Property(r => r.Supplier).HasMaxLength(100);
            e.Property(r => r.UnitCost).HasColumnType("decimal(18,2)");
            e.Property(r => r.UnitOfMeasure).HasMaxLength(20);
            e.Property(r => r.Location).HasMaxLength(50);
            e.HasIndex(r => r.MaterialCode).IsUnique();
            e.HasIndex(r => r.Supplier);
            e.Ignore(r => r.AvailableQuantity);
            e.Ignore(r => r.TotalValue);
        });

        // Customer
        modelBuilder.Entity<Customer>(e =>
        {
            e.HasKey(c => c.CustomerId);
            e.Property(c => c.CustomerCode).HasMaxLength(50).IsRequired();
            e.Property(c => c.CompanyName).HasMaxLength(100).IsRequired();
            e.Property(c => c.ContactName).HasMaxLength(100);
            e.Property(c => c.Phone).HasMaxLength(20);
            e.Property(c => c.Email).HasMaxLength(100);
            e.Property(c => c.Address).HasMaxLength(255);
            e.Property(c => c.City).HasMaxLength(50);
            e.Property(c => c.State).HasMaxLength(2);
            e.Property(c => c.ZipCode).HasMaxLength(10);
            e.Property(c => c.Country).HasMaxLength(50);
            e.Property(c => c.PaymentTerms).HasMaxLength(50);
            e.HasIndex(c => c.CustomerCode).IsUnique();
        });

        // WorkOrder
        modelBuilder.Entity<WorkOrder>(e =>
        {
            e.HasKey(w => w.WorkOrderId);
            e.Property(w => w.WorkOrderNumber).HasMaxLength(50).IsRequired();
            e.Property(w => w.PartNumber).HasMaxLength(50).IsRequired();
            e.Property(w => w.Status).HasConversion<string>().HasMaxLength(50);
            e.Property(w => w.Priority).HasConversion<string>().HasMaxLength(20);
            e.Property(w => w.CustomerPO).HasMaxLength(50);
            e.Property(w => w.AssignedTo).HasMaxLength(50);
            e.Property(w => w.EstimatedCost).HasColumnType("decimal(18,2)");
            e.Property(w => w.ActualCost).HasColumnType("decimal(18,2)");
            e.Property(w => w.CreatedBy).HasMaxLength(50);
            e.HasIndex(w => w.WorkOrderNumber).IsUnique();
            e.HasIndex(w => w.Status);
            e.HasIndex(w => w.DueDate);
            e.HasIndex(w => w.PartNumber);
            e.HasOne(w => w.Customer).WithMany(c => c.WorkOrders)
                .HasForeignKey(w => w.CustomerId).OnDelete(DeleteBehavior.SetNull);
            e.HasOne(w => w.Part).WithMany(p => p.WorkOrders)
                .HasForeignKey(w => w.PartNumber).HasPrincipalKey(p => p.PartNumber)
                .OnDelete(DeleteBehavior.Restrict);
            e.Ignore(w => w.PercentComplete);
            e.Ignore(w => w.IsOverdue);
            e.Ignore(w => w.DaysRemaining);
        });

        // QualityCheck
        modelBuilder.Entity<QualityCheck>(e =>
        {
            e.HasKey(q => q.CheckId);
            e.Property(q => q.Inspector).HasMaxLength(100).IsRequired();
            e.Property(q => q.Result).HasConversion<string>().HasMaxLength(50);
            e.Property(q => q.Measurements).HasMaxLength(255);
            e.Property(q => q.CreatedBy).HasMaxLength(50);
            e.HasIndex(q => q.WorkOrderId);
            e.HasIndex(q => q.Result);
            e.HasOne(q => q.WorkOrder).WithMany(w => w.QualityChecks)
                .HasForeignKey(q => q.WorkOrderId).OnDelete(DeleteBehavior.Cascade);
        });

        // ShippingManifest
        modelBuilder.Entity<ShippingManifest>(e =>
        {
            e.HasKey(s => s.ManifestId);
            e.Property(s => s.ManifestNumber).HasMaxLength(50).IsRequired();
            e.Property(s => s.Carrier).HasMaxLength(50);
            e.Property(s => s.TrackingNumber).HasMaxLength(100);
            e.Property(s => s.ShipToName).HasMaxLength(100);
            e.Property(s => s.ShipToAddress).HasMaxLength(255);
            e.Property(s => s.ShipToCity).HasMaxLength(50);
            e.Property(s => s.ShipToState).HasMaxLength(2);
            e.Property(s => s.ShipToZipCode).HasMaxLength(10);
            e.Property(s => s.FreightCost).HasColumnType("decimal(18,2)");
            e.Property(s => s.CreatedBy).HasMaxLength(50);
            e.HasIndex(s => s.ManifestNumber).IsUnique();
            e.HasIndex(s => s.WorkOrderId);
            e.HasOne(s => s.WorkOrder).WithMany(w => w.ShippingManifests)
                .HasForeignKey(s => s.WorkOrderId).OnDelete(DeleteBehavior.Cascade);
        });

        // ProductionLog
        modelBuilder.Entity<ProductionLog>(e =>
        {
            e.HasKey(p => p.LogId);
            e.Property(p => p.Shift).HasMaxLength(20);
            e.Property(p => p.Operator).HasMaxLength(100);
            e.Property(p => p.MachineId).HasMaxLength(50);
            e.HasIndex(p => p.WorkOrderId);
            e.HasOne(p => p.WorkOrder).WithMany(w => w.ProductionLogs)
                .HasForeignKey(p => p.WorkOrderId).OnDelete(DeleteBehavior.Cascade);
            e.HasOne(p => p.Machine).WithMany()
                .HasForeignKey(p => p.MachineId).OnDelete(DeleteBehavior.SetNull);
        });

        // MachineStatus
        modelBuilder.Entity<MachineStatus>(e =>
        {
            e.HasKey(m => m.MachineId);
            e.Property(m => m.MachineId).HasMaxLength(50);
            e.Property(m => m.MachineName).HasMaxLength(100).IsRequired();
            e.Property(m => m.Status).HasMaxLength(50);
            e.Property(m => m.CurrentOperator).HasMaxLength(100);
            e.Property(m => m.Location).HasMaxLength(50);
            e.HasOne(m => m.CurrentWorkOrder).WithMany()
                .HasForeignKey(m => m.CurrentWorkOrderId).OnDelete(DeleteBehavior.SetNull);
        });

        // ShiftSchedule
        modelBuilder.Entity<ShiftSchedule>(e =>
        {
            e.HasKey(s => s.ScheduleId);
            e.Property(s => s.ShiftName).HasMaxLength(50).IsRequired();
            e.Property(s => s.Supervisor).HasMaxLength(100);
        });

        // InventoryTransaction
        modelBuilder.Entity<InventoryTransaction>(e =>
        {
            e.HasKey(i => i.TransactionId);
            e.Property(i => i.Reason).HasMaxLength(255).IsRequired();
            e.HasOne(i => i.Material).WithMany(r => r.Transactions)
                .HasForeignKey(i => i.MaterialId).OnDelete(DeleteBehavior.Cascade);
        });
    }
}
