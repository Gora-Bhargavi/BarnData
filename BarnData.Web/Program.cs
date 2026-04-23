using BarnData.Data;
using BarnData.Core.Services;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

// 1. Database 
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
            // Compatibility level 130 = SQL Server 2016
            // Prevents EF Core 8 from using OPENJSON / $ JSON path syntax
            //sqlOptions.UseCompatibilityLevel(130);
        }
    )
    //.LogTo(Console.WriteLine, LogLevel.Information)        
    //.EnableSensitiveDataLogging() 
);

builder.Services.AddDistributedMemoryCache();

// Raise form field limit (safety net for large grids)
builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(o =>
{
    o.ValueCountLimit  = 32768;
    o.ValueLengthLimit = int.MaxValue;
    o.MultipartBodyLengthLimit = int.MaxValue;
});
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromMinutes(30);
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

//  2. Services (business logic) 
builder.Services.AddScoped<IAnimalService, AnimalService>();
builder.Services.AddScoped<IVendorService, VendorService>();

//  3. MVC 
builder.Services
    .AddControllersWithViews()
    .AddSessionStateTempDataProvider();

var app = builder.Build();

//  4. Middleware pipeline 
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseStaticFiles();
app.UseRouting();
app.UseSession();
app.UseAuthorization();

// Rotativa PDF engine 
Rotativa.AspNetCore.RotativaConfiguration.Setup(
    app.Environment.WebRootPath,
    wkhtmltopdfRelativePath: "Rotativa"
);

// 5. Routes 
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Animal}/{action=Index}/{id?}"
);

//  6. Verify DB connection on startup 
using (var scope = app.Services.CreateScope())
{
    var ctx = scope.ServiceProvider.GetRequiredService<BarnDataContext>();
    try
    {
        // Use raw SQL test instead of CanConnect() to avoid EF Core internal queries
        ctx.Database.ExecuteSqlRaw("SELECT 1");
        Console.WriteLine("[BarnData] Database connection OK.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[BarnData] Database connection FAILED: {ex.Message}");
    }
}

app.Run();
