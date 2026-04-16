using PrecisionParts.Core.Enums;
using PrecisionParts.Core.Models;

namespace PrecisionParts.Core.Services;

public interface IInventoryService
{
    Task<IEnumerable<RawMaterial>> GetAllAsync(string? supplierFilter = null);
    Task<RawMaterial?> GetByIdAsync(int materialId);
    Task<RawMaterial> CreateAsync(RawMaterial material);
    Task<RawMaterial> UpdateAsync(RawMaterial material);
    Task<bool> DeleteAsync(int materialId);
    Task<bool> UpdateQuantityAsync(int materialId, double quantityChange, string reason, int? userId = null);
    Task<IEnumerable<RawMaterial>> GetLowStockItemsAsync();
    Task<IEnumerable<InventoryTransaction>> GetTransactionsAsync(int materialId);
}
