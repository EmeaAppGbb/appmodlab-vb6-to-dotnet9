using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PrecisionParts.Core.Enums;
using PrecisionParts.Core.Models;
using PrecisionParts.Core.Services;
using PrecisionParts.Infrastructure.Data;

namespace PrecisionParts.Infrastructure.Services;

/// <summary>
/// Shipping service — replaces VB6 frmShipping business logic.
/// Key rule: shipping auto-completes the work order.
/// </summary>
public class ShippingService : IShippingService
{
    private readonly PrecisionPartsDbContext _db;
    private readonly IWorkOrderService _workOrderService;
    private readonly ILogger<ShippingService> _logger;

    public ShippingService(
        PrecisionPartsDbContext db,
        IWorkOrderService workOrderService,
        ILogger<ShippingService> logger)
    {
        _db = db;
        _workOrderService = workOrderService;
        _logger = logger;
    }

    public async Task<IEnumerable<ShippingManifest>> GetAllAsync()
    {
        return await _db.ShippingManifests
            .Include(s => s.WorkOrder)
            .OrderByDescending(s => s.ShipDate)
            .ToListAsync();
    }

    public async Task<ShippingManifest?> GetByIdAsync(int manifestId)
    {
        return await _db.ShippingManifests
            .Include(s => s.WorkOrder)
            .FirstOrDefaultAsync(s => s.ManifestId == manifestId);
    }

    /// <summary>
    /// Creates a shipping manifest and auto-completes the work order.
    /// Business rule: only completed work orders can be shipped.
    /// </summary>
    public async Task<ShippingManifest> CreateAsync(ShippingManifest manifest)
    {
        if (string.IsNullOrWhiteSpace(manifest.ManifestNumber))
            manifest.ManifestNumber = GenerateManifestNumber();

        manifest.CreatedDate = DateTime.UtcNow;

        _db.ShippingManifests.Add(manifest);

        // Auto-complete the work order on shipping
        await _workOrderService.UpdateStatusAsync(
            manifest.WorkOrderId, WorkOrderStatus.Completed);

        await _db.SaveChangesAsync();

        _logger.LogInformation("Shipping manifest {Number} created for WO {WoId}",
            manifest.ManifestNumber, manifest.WorkOrderId);

        return manifest;
    }

    public async Task<IEnumerable<ShippingManifest>> GetByWorkOrderAsync(int workOrderId)
    {
        return await _db.ShippingManifests
            .Where(s => s.WorkOrderId == workOrderId)
            .OrderByDescending(s => s.ShipDate)
            .ToListAsync();
    }

    public string GenerateManifestNumber()
    {
        var datePart = DateTime.UtcNow.ToString("yyyyMMdd");
        var todayCount = _db.ShippingManifests
            .Count(s => s.ManifestNumber.StartsWith($"SM-{datePart}"));
        return $"SM-{datePart}-{(todayCount + 1):D3}";
    }
}
