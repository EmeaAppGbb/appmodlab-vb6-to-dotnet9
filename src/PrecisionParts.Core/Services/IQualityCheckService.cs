using PrecisionParts.Core.Enums;
using PrecisionParts.Core.Models;

namespace PrecisionParts.Core.Services;

public interface IQualityCheckService
{
    Task<IEnumerable<QualityCheck>> GetByWorkOrderAsync(int workOrderId);
    Task<QualityCheck?> GetByIdAsync(int checkId);
    Task<QualityCheck> CreateAsync(QualityCheck qualityCheck);
    Task<IEnumerable<QualityCheck>> GetByDateRangeAsync(DateTime from, DateTime to);
}
