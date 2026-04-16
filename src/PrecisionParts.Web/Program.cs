using Microsoft.EntityFrameworkCore;
using PrecisionParts.Core.Services;
using PrecisionParts.Infrastructure.Data;
using PrecisionParts.Infrastructure.Services;
using PrecisionParts.Web.Components;

var builder = WebApplication.CreateBuilder(args);

// Configuration — replaces VB6 modGlobals registry-based config
builder.Services.Configure<CostCalculationSettings>(
    builder.Configuration.GetSection(CostCalculationSettings.SectionName));

// Database — replaces VB6 modDatabase global ADO connection
builder.Services.AddDbContext<PrecisionPartsDbContext>(options =>
    options.UseInMemoryDatabase("PrecisionParts"));

// Services — replaces VB6 global modules and classes with DI
builder.Services.AddScoped<IWorkOrderService, WorkOrderService>();
builder.Services.AddScoped<IInventoryService, InventoryService>();
builder.Services.AddScoped<IPartService, PartService>();
builder.Services.AddScoped<IQualityCheckService, QualityCheckService>();
builder.Services.AddScoped<IShippingService, ShippingService>();

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

var app = builder.Build();

// Seed database — replaces VB6 seed_data.sql
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<PrecisionPartsDbContext>();
    await DatabaseSeeder.SeedAsync(db);
}

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
