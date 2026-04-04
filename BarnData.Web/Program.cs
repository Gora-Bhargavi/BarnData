using BarnData.Data;
using BarnData.Core.Services;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

// ── 1. Database ───────────────────────────────────────────────────────────
builder.Services.AddDbContext<BarnDataContext>(options =>
    options.UseSqlServer(
        builder.Configuration.GetConnectionString("BarnData"),
        sqlOptions =>
        {
            sqlOptions.EnableRetryOnFailure(
                maxRetryCount: 3,
                maxRetryDelay: TimeSpan.FromSeconds(5),
                errorNumbersToAdd: null
            );
        }
    )
);

// ── 2. Services (business logic) ──────────────────────────────────────────
builder.Services.AddScoped<IAnimalService, AnimalService>();
builder.Services.AddScoped<IVendorService, VendorService>();

// ── 3. MVC ────────────────────────────────────────────────────────────────
builder.Services.AddControllersWithViews();

var app = builder.Build();

// ── 4. Middleware pipeline ────────────────────────────────────────────────
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

// ── 5. Routes ─────────────────────────────────────────────────────────────
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Animal}/{action=Index}/{id?}"
);

// ── 6. Verify DB connection on startup ────────────────────────────────────
using (var scope = app.Services.CreateScope())
{
    var context = scope.ServiceProvider.GetRequiredService<BarnDataContext>();
    try
    {
        context.Database.CanConnect();
        Console.WriteLine("[BarnData] Database connection OK.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[BarnData] Database connection FAILED: {ex.Message}");
    }
}

app.Run();
