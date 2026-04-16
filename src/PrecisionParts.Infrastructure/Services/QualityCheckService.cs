using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PrecisionParts.Core.Enums;
using PrecisionParts.Core.Models;
using PrecisionParts.Core.Services;
using PrecisionParts.Infrastructure.Data;

namespace PrecisionParts.Infrastructure.Services;

/// <summary>
/// Quality check service — replaces VB6 frmQualityCheck business logic.
/// Key rule: Fail result auto-sets work order to OnHold.
/// </summary>
public class QualityCheckService : IQualityCheckService
{
    private readonly PrecisionPartsDbContext _db;
    private readonly IWorkOrderService _workOrderService;
    private readonly ILogger<QualityCheckService> _logger;

    public QualityCheckService(
        PrecisionPartsDbContext db,
        IWorkOrderService workOrderService,
        ILogger<QualityCheckService> logger)
    {
        _db = db;
        _workOrderService = workOrderService;
        _logger = logger;
    }

    public async Task<IEnumerable<QualityCheck>> GetByWorkOrderAsync(int workOrderId)
    {
        return await _db.QualityChecks
            .Where(q => q.WorkOrderId == workOrderId)
            .OrderByDescending(q => q.CheckDate)
            .ToListAsync();
    }

    public async Task<QualityCheck?> GetByIdAsync(int checkId)
    {
        return await _db.QualityChecks
            .Include(q => q.WorkOrder)
            .FirstOrDefaultAsync(q => q.CheckId == checkId);
    }

    /// <summary>
    /// Creates a QC record and applies business rules.
    /// If result is Fail, the work order is automatically put OnHold.
    /// </summary>
    public async Task<QualityCheck> CreateAsync(QualityCheck qualityCheck)
    {
        qualityCheck.CreatedDate = DateTime.UtcNow;

        _db.QualityChecks.Add(qualityCheck);
        await _db.SaveChangesAsync();

        // Business rule: Fail → auto-hold work order
        if (qualityCheck.Result == QualityCheckResult.Fail)
        {
            await _workOrderService.UpdateStatusAsync(
                qualityCheck.WorkOrderId, WorkOrderStatus.OnHold);

            _logger.LogWarning("QC Fail for WO {Id} — work order set to On Hold",
                qualityCheck.WorkOrderId);
        }

        _logger.LogInformation("QC Check {Id} recorded: {Result} for WO {WoId}",
            qualityCheck.CheckId, qualityCheck.Result, qualityCheck.WorkOrderId);

        return qualityCheck;
    }

    public async Task<IEnumerable<QualityCheck>> GetByDateRangeAsync(DateTime from, DateTime to)
    {
        return await _db.QualityChecks
            .Include(q => q.WorkOrder)
            .Where(q => q.CheckDate >= from && q.CheckDate <= to)
            .OrderByDescending(q => q.CheckDate)
            .ToListAsync();
    }
}
