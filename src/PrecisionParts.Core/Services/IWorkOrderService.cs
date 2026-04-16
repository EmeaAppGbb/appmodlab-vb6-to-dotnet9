using PrecisionParts.Core.Enums;
using PrecisionParts.Core.Models;

namespace PrecisionParts.Core.Services;

public interface IWorkOrderService
{
    Task<IEnumerable<WorkOrder>> GetAllAsync(WorkOrderStatus? statusFilter = null);
    Task<WorkOrder?> GetByIdAsync(int workOrderId);
    Task<WorkOrder?> GetByNumberAsync(string workOrderNumber);
    Task<WorkOrder> CreateAsync(WorkOrder workOrder);
    Task<WorkOrder> UpdateAsync(WorkOrder workOrder);
    Task<bool> DeleteAsync(int workOrderId);
    Task<bool> UpdateStatusAsync(int workOrderId, WorkOrderStatus newStatus);
    Task<decimal> CalculateCostAsync(string partNumber, double quantity);
    string GenerateWorkOrderNumber();
}
