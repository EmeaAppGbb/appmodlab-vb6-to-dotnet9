using PrecisionParts.Core.Models;

namespace PrecisionParts.Core.Services;

public interface IShippingService
{
    Task<IEnumerable<ShippingManifest>> GetAllAsync();
    Task<ShippingManifest?> GetByIdAsync(int manifestId);
    Task<ShippingManifest> CreateAsync(ShippingManifest manifest);
    Task<IEnumerable<ShippingManifest>> GetByWorkOrderAsync(int workOrderId);
    string GenerateManifestNumber();
}
