namespace PrecisionParts.Core.Models;

/// <summary>
/// Shift schedule — migrated from VB6 ShiftSchedule table.
/// </summary>
public class ShiftSchedule
{
    public int ScheduleId { get; set; }
    public string ShiftName { get; set; } = string.Empty;
    public TimeSpan StartTime { get; set; }
    public TimeSpan EndTime { get; set; }
    public DayOfWeek? DayOfWeek { get; set; }
    public string? Supervisor { get; set; }
    public bool IsActive { get; set; } = true;
}
