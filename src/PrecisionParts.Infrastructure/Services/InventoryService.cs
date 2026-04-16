using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PrecisionParts.Core.Enums;
using PrecisionParts.Core.Models;
using PrecisionParts.Core.Services;
using PrecisionParts.Infrastructure.Data;

namespace PrecisionParts.Infrastructure.Services;

/// <summary>
/// Inventory service — replaces VB6 clsInventory.cls.
/// Manages stock levels, reorder alerts, and transaction logging.
/// </summary>
public class InventoryService : IInventoryService
{
    private readonly PrecisionPartsDbContext _db;
    private readonly ILogger<InventoryService> _logger;

    public InventoryService(PrecisionPartsDbContext db, ILogger<InventoryService> logger)
    {
        _db = db;
        _logger = logger;
    }

    public async Task<IEnumerable<RawMaterial>> GetAllAsync(string? supplierFilter = null)
    {
        var query = _db.RawMaterials.AsQueryable();

        if (!string.IsNullOrWhiteSpace(supplierFilter))
            query = query.Where(r => r.Supplier == supplierFilter);

        return await query.OrderBy(r => r.MaterialName).ToListAsync();
    }

    public async Task<RawMaterial?> GetByIdAsync(int materialId)
    {
        return await _db.RawMaterials
            .Include(r => r.Transactions.OrderByDescending(t => t.TransactionDate).Take(10))
            .FirstOrDefaultAsync(r => r.MaterialId == materialId);
    }

    public async Task<RawMaterial> CreateAsync(RawMaterial material)
    {
        material.CreatedDate = DateTime.UtcNow;
        _db.RawMaterials.Add(material);
        await _db.SaveChangesAsync();

        _logger.LogInformation("Material {Name} created (ID: {Id})",
            material.MaterialName, material.MaterialId);
        return material;
    }

    public async Task<RawMaterial> UpdateAsync(RawMaterial material)
    {
        material.ModifiedDate = DateTime.UtcNow;
        _db.RawMaterials.Update(material);
        await _db.SaveChangesAsync();

        _logger.LogInformation("Material {Id} updated", material.MaterialId);
        return material;
    }

    public async Task<bool> DeleteAsync(int materialId)
    {
        var material = await _db.RawMaterials.FindAsync(materialId);
        if (material is null) return false;

        _db.RawMaterials.Remove(material);
        await _db.SaveChangesAsync();

        _logger.LogInformation("Material {Id} deleted", materialId);
        return true;
    }

    /// <summary>
    /// Replaces VB6 clsInventory.UpdateQuantity with transaction logging.
    /// Validates non-negative inventory and logs the change.
    /// </summary>
    public async Task<bool> UpdateQuantityAsync(int materialId, double quantityChange, string reason, int? userId = null)
    {
        var material = await _db.RawMaterials.FindAsync(materialId);
        if (material is null) return false;

        var newQty = material.QuantityOnHand + quantityChange;
        if (newQty < 0)
        {
            _logger.LogWarning("Insufficient stock for Material {Id}: have {Qty}, requested {Change}",
                materialId, material.QuantityOnHand, Math.Abs(quantityChange));
            return false;
        }

        material.QuantityOnHand = newQty;
        material.ModifiedDate = DateTime.UtcNow;

        // Log the transaction (replaces VB6 LogInventoryTransaction)
        var transaction = new InventoryTransaction
        {
            MaterialId = materialId,
            QuantityChange = quantityChange,
            QuantityAfter = newQty,
            Reason = reason,
            TransactionDate = DateTime.UtcNow,
            UserId = userId
        };
        _db.InventoryTransactions.Add(transaction);

        await _db.SaveChangesAsync();

        var status = material.GetStockStatus();
        if (status is StockStatus.ReorderRequired or StockStatus.Critical)
        {
            _logger.LogWarning("Material {Name} stock alert: {Status} (Qty: {Qty})",
                material.MaterialName, status, material.QuantityOnHand);
        }

        return true;
    }

    public async Task<IEnumerable<RawMaterial>> GetLowStockItemsAsync()
    {
        var materials = await _db.RawMaterials.Where(r => r.IsActive).ToListAsync();
        return materials.Where(m => m.NeedsReorder());
    }

    public async Task<IEnumerable<InventoryTransaction>> GetTransactionsAsync(int materialId)
    {
        return await _db.InventoryTransactions
            .Where(t => t.MaterialId == materialId)
            .OrderByDescending(t => t.TransactionDate)
            .ToListAsync();
    }
}
