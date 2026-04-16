using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PrecisionParts.Core.Enums;
using PrecisionParts.Core.Models;
using PrecisionParts.Core.Services;
using PrecisionParts.Infrastructure.Data;

namespace PrecisionParts.Infrastructure.Services;

/// <summary>
/// Work order service — replaces VB6 clsWorkOrder.cls.
/// Separates business logic from data access using EF Core.
/// </summary>
public class WorkOrderService : IWorkOrderService
{
    private readonly PrecisionPartsDbContext _db;
    private readonly CostCalculationSettings _costSettings;
    private readonly ILogger<WorkOrderService> _logger;

    public WorkOrderService(
        PrecisionPartsDbContext db,
        IOptions<CostCalculationSettings> costSettings,
        ILogger<WorkOrderService> logger)
    {
        _db = db;
        _costSettings = costSettings.Value;
        _logger = logger;
    }

    public async Task<IEnumerable<WorkOrder>> GetAllAsync(WorkOrderStatus? statusFilter = null)
    {
        var query = _db.WorkOrders
            .Include(w => w.Customer)
            .Include(w => w.Part)
            .AsQueryable();

        if (statusFilter.HasValue)
            query = query.Where(w => w.Status == statusFilter.Value);

        return await query.OrderByDescending(w => w.CreatedDate).ToListAsync();
    }

    public async Task<WorkOrder?> GetByIdAsync(int workOrderId)
    {
        return await _db.WorkOrders
            .Include(w => w.Customer)
            .Include(w => w.Part)
            .Include(w => w.QualityChecks)
            .Include(w => w.ShippingManifests)
            .FirstOrDefaultAsync(w => w.WorkOrderId == workOrderId);
    }

    public async Task<WorkOrder?> GetByNumberAsync(string workOrderNumber)
    {
        return await _db.WorkOrders
            .Include(w => w.Customer)
            .Include(w => w.Part)
            .FirstOrDefaultAsync(w => w.WorkOrderNumber == workOrderNumber);
    }

    public async Task<WorkOrder> CreateAsync(WorkOrder workOrder)
    {
        if (string.IsNullOrWhiteSpace(workOrder.WorkOrderNumber))
            workOrder.WorkOrderNumber = GenerateWorkOrderNumber();

        workOrder.CreatedDate = DateTime.UtcNow;
        workOrder.EstimatedCost = await CalculateCostAsync(workOrder.PartNumber, workOrder.Quantity);

        _db.WorkOrders.Add(workOrder);
        await _db.SaveChangesAsync();

        _logger.LogInformation("Work Order {Number} created (ID: {Id})",
            workOrder.WorkOrderNumber, workOrder.WorkOrderId);

        return workOrder;
    }

    public async Task<WorkOrder> UpdateAsync(WorkOrder workOrder)
    {
        workOrder.ModifiedDate = DateTime.UtcNow;

        // Auto-set completion date when status changes to Completed
        if (workOrder.Status == WorkOrderStatus.Completed && !workOrder.CompletedDate.HasValue)
            workOrder.CompletedDate = DateTime.UtcNow;

        _db.WorkOrders.Update(workOrder);
        await _db.SaveChangesAsync();

        _logger.LogInformation("Work Order {Number} updated", workOrder.WorkOrderNumber);
        return workOrder;
    }

    public async Task<bool> DeleteAsync(int workOrderId)
    {
        var workOrder = await _db.WorkOrders
            .Include(w => w.QualityChecks)
            .Include(w => w.ShippingManifests)
            .FirstOrDefaultAsync(w => w.WorkOrderId == workOrderId);

        if (workOrder is null) return false;

        // Business rule: cannot delete if QC or shipping records exist
        if (workOrder.QualityChecks.Any())
        {
            _logger.LogWarning("Cannot delete WO {Id} — quality check records exist", workOrderId);
            return false;
        }

        if (workOrder.ShippingManifests.Any())
        {
            _logger.LogWarning("Cannot delete WO {Id} — shipping records exist", workOrderId);
            return false;
        }

        _db.WorkOrders.Remove(workOrder);
        await _db.SaveChangesAsync();

        _logger.LogInformation("Work Order {Id} deleted", workOrderId);
        return true;
    }

    public async Task<bool> UpdateStatusAsync(int workOrderId, WorkOrderStatus newStatus)
    {
        var workOrder = await _db.WorkOrders.FindAsync(workOrderId);
        if (workOrder is null) return false;

        workOrder.Status = newStatus;
        workOrder.ModifiedDate = DateTime.UtcNow;

        if (newStatus == WorkOrderStatus.Completed)
            workOrder.CompletedDate ??= DateTime.UtcNow;

        await _db.SaveChangesAsync();

        _logger.LogInformation("Work Order {Id} status changed to {Status}",
            workOrderId, newStatus);
        return true;
    }

    /// <summary>
    /// Replaces VB6 clsWorkOrder.CalculateCost with configurable rates.
    /// Formula: Total = (MaterialCost + LaborCost) × (1 + OverheadPct)
    /// </summary>
    public async Task<decimal> CalculateCostAsync(string partNumber, double quantity)
    {
        var part = await _db.Parts.FirstOrDefaultAsync(p => p.PartNumber == partNumber);
        var unitCost = part?.UnitCost ?? 0m;

        var materialCost = unitCost * (decimal)quantity;
        var laborCost = (decimal)quantity * _costSettings.HoursPerPart * _costSettings.LaborRatePerHour;
        var total = (materialCost + laborCost) * (1 + _costSettings.OverheadPercentage);

        return Math.Round(total, 2);
    }

    /// <summary>
    /// Replaces VB6 modUtilities.GenerateWorkOrderNumber.
    /// Format: WO-YYYYMMDD-NNN
    /// </summary>
    public string GenerateWorkOrderNumber()
    {
        var datePart = DateTime.UtcNow.ToString("yyyyMMdd");
        var todayCount = _db.WorkOrders
            .Count(w => w.WorkOrderNumber.StartsWith($"WO-{datePart}"));
        return $"WO-{datePart}-{(todayCount + 1):D3}";
    }
}
