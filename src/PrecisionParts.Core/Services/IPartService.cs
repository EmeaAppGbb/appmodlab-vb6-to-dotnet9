using PrecisionParts.Core.Models;

namespace PrecisionParts.Core.Services;

public interface IPartService
{
    Task<IEnumerable<Part>> GetAllAsync(bool activeOnly = true);
    Task<Part?> GetByIdAsync(int partId);
    Task<Part?> GetByNumberAsync(string partNumber);
    Task<IEnumerable<Part>> SearchAsync(string searchTerm);
    Task<Part> CreateAsync(Part part);
    Task<Part> UpdateAsync(Part part);
    Task<bool> DeleteAsync(int partId);
    bool ValidatePartNumber(string partNumber, out string errorMessage);
}
