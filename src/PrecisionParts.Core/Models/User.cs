using PrecisionParts.Core.Enums;

namespace PrecisionParts.Core.Models;

/// <summary>
/// Application user — replaces VB6 Users table with plain-text passwords.
/// In .NET 9 this maps to ASP.NET Core Identity for secure auth.
/// </summary>
public class User
{
    public int UserId { get; set; }
    public string Username { get; set; } = string.Empty;
    public string PasswordHash { get; set; } = string.Empty;
    public UserRole Role { get; set; } = UserRole.Operator;
    public string FullName { get; set; } = string.Empty;
    public string? Email { get; set; }
    public bool IsActive { get; set; } = true;
    public DateTime CreatedDate { get; set; } = DateTime.UtcNow;
    public DateTime? LastLogin { get; set; }
}
