using System.Text.RegularExpressions;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PrecisionParts.Core.Models;
using PrecisionParts.Core.Services;
using PrecisionParts.Infrastructure.Data;

namespace PrecisionParts.Infrastructure.Services;

/// <summary>
/// Part service — replaces VB6 clsPart.cls.
/// Part number validation extracted from VB6 ValidatePartNumber.
/// </summary>
public partial class PartService : IPartService
{
    private readonly PrecisionPartsDbContext _db;
    private readonly ILogger<PartService> _logger;

    // Valid part number prefixes (from VB6 business rules)
    private static readonly HashSet<string> KnownPrefixes =
        ["PP", "PPM", "PPA", "MFG", "ASM", "RAW", "HDW", "ELC"];

    public PartService(PrecisionPartsDbContext db, ILogger<PartService> logger)
    {
        _db = db;
        _logger = logger;
    }

    public async Task<IEnumerable<Part>> GetAllAsync(bool activeOnly = true)
    {
        var query = _db.Parts.AsQueryable();
        if (activeOnly) query = query.Where(p => p.IsActive);
        return await query.OrderBy(p => p.PartNumber).ToListAsync();
    }

    public async Task<Part?> GetByIdAsync(int partId)
    {
        return await _db.Parts.FindAsync(partId);
    }

    public async Task<Part?> GetByNumberAsync(string partNumber)
    {
        return await _db.Parts.FirstOrDefaultAsync(p => p.PartNumber == partNumber);
    }

    /// <summary>
    /// Replaces VB6 frmPartLookup search functionality.
    /// Searches part number, description, and material.
    /// </summary>
    public async Task<IEnumerable<Part>> SearchAsync(string searchTerm)
    {
        if (string.IsNullOrWhiteSpace(searchTerm))
            return await GetAllAsync();

        var term = searchTerm.ToUpperInvariant();
        return await _db.Parts
            .Where(p => p.PartNumber.Contains(term)
                     || p.Description.Contains(term)
                     || (p.Material != null && p.Material.Contains(term)))
            .OrderBy(p => p.PartNumber)
            .ToListAsync();
    }

    public async Task<Part> CreateAsync(Part part)
    {
        part.PartNumber = part.PartNumber.ToUpperInvariant().Trim();
        part.CreatedDate = DateTime.UtcNow;

        // Check for duplicate
        if (await _db.Parts.AnyAsync(p => p.PartNumber == part.PartNumber))
            throw new InvalidOperationException($"Part Number '{part.PartNumber}' already exists.");

        _db.Parts.Add(part);
        await _db.SaveChangesAsync();

        _logger.LogInformation("Part {Number} created", part.PartNumber);
        return part;
    }

    public async Task<Part> UpdateAsync(Part part)
    {
        part.ModifiedDate = DateTime.UtcNow;
        _db.Parts.Update(part);
        await _db.SaveChangesAsync();

        _logger.LogInformation("Part {Number} updated", part.PartNumber);
        return part;
    }

    public async Task<bool> DeleteAsync(int partId)
    {
        var part = await _db.Parts
            .Include(p => p.WorkOrders)
            .FirstOrDefaultAsync(p => p.PartId == partId);

        if (part is null) return false;

        // Business rule: cannot delete if work orders reference this part
        if (part.WorkOrders.Any())
        {
            _logger.LogWarning("Cannot delete Part {Number} — work orders exist", part.PartNumber);
            return false;
        }

        _db.Parts.Remove(part);
        await _db.SaveChangesAsync();

        _logger.LogInformation("Part {Number} deleted", part.PartNumber);
        return true;
    }

    /// <summary>
    /// Replaces VB6 clsPart.ValidatePartNumber.
    /// Format: 2-4 letter prefix, hyphen, 1-8 char number, optional additional segments.
    /// Total max 20 chars.
    /// </summary>
    public bool ValidatePartNumber(string partNumber, out string errorMessage)
    {
        errorMessage = string.Empty;

        if (string.IsNullOrWhiteSpace(partNumber))
        {
            errorMessage = "Part Number is required.";
            return false;
        }

        partNumber = partNumber.ToUpperInvariant().Trim();

        if (!partNumber.Contains('-'))
        {
            errorMessage = "Part Number must contain at least one hyphen (e.g., PP-1234-001).";
            return false;
        }

        var segments = partNumber.Split('-');
        if (segments.Length < 2 || segments.Length > 4)
        {
            errorMessage = "Part Number must have 2 to 4 segments separated by hyphens.";
            return false;
        }

        // Prefix validation: 2-4 alpha characters
        var prefix = segments[0];
        if (prefix.Length < 2 || prefix.Length > 4)
        {
            errorMessage = "Part Number prefix must be 2-4 characters.";
            return false;
        }

        if (!prefix.All(char.IsLetter))
        {
            errorMessage = "Part Number prefix must contain only letters.";
            return false;
        }

        if (!KnownPrefixes.Contains(prefix))
            _logger.LogWarning("Unknown part number prefix: {Prefix}", prefix);

        // Second segment: 1-8 characters
        if (segments[1].Length < 1 || segments[1].Length > 8)
        {
            errorMessage = "Part Number numeric segment must be 1-8 characters.";
            return false;
        }

        if (partNumber.Length > 20)
        {
            errorMessage = "Part Number cannot exceed 20 characters.";
            return false;
        }

        return true;
    }
}
