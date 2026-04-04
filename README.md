# BarnData — Barn Data Entry System
Access replacement for Trax-IT-Slaughter | Phase 3

---

## Project structure

```
BarnData.sln
├── BarnData.Web/       ASP.NET Core MVC — UI, controllers, views
├── BarnData.Data/      EF Core — entities, DbContext, SQL Server connection
└── BarnData.Core/      Business logic — AnimalService, VendorService, validation
```

---

## Setup steps (do in this exact order)

### 1. Open the solution
Open `BarnData.sln` in Visual Studio 2022.

### 2. Update the connection string
Open `BarnData.Web/appsettings.json` and replace `YOUR_PASSWORD_HERE` with your actual SQL Server password:
```json
"BarnData": "Server=192.168.50.126;Database=Trax-IT-Slaughter;User Id=symcod;Password=ACTUAL_PASSWORD;TrustServerCertificate=True;"
```

### 3. Restore NuGet packages
In Visual Studio: right-click solution → Restore NuGet Packages.
Or in terminal:
```
dotnet restore
```

### 4. Build the solution
```
dotnet build
```
Should build with 0 errors. If EF Core packages fail, run `dotnet restore` again.

### 5. Verify DB connection
Run the project (F5). Check the console output for:
```
[BarnData] Database connection OK.
```
If you see FAILED — check the connection string and that your IP can reach 192.168.50.126.

### 6. Set BarnData.Web as startup project
Right-click BarnData.Web → Set as Startup Project.

---

## Phase 3 — what is left to build

After Step 1 (this setup) is complete, the remaining work is:

| Step | What to build | File |
|------|--------------|------|
| Step 2 | Animal list view (Index) | Animal/Index.cshtml |
| Step 3 | Animal entry form (Create) | Animal/Create.cshtml |
| Step 4 | Animal edit form (Edit) | Animal/Edit.cshtml |
| Step 5 | AnimalController wiring | AnimalController.cs |
| Step 6 | Tally report view | Report/Tally.cshtml |
| Step 7 | ReportController | ReportController.cs |
| Step 8 | PDF + Excel export | ReportController.cs |

---

## Key business rules (implemented in AnimalService.cs)

- **Tag Number 1 must be unique** per vendor per kill date — duplicate blocks save
- **Live weight range** 300–2500 lbs — outside this shows a warning (does not block)
- **Kill date** must be >= Purchase date — violation blocks save
- **Soft delete only** — KillStatus set to 'Flagged', never hard deleted
- **Grade 2** is entered manually by pricing staff later — NOT auto-filled from HotScale

---

## Notes for Adam (production deployment)

When ready to go live, run all CREATE TABLE scripts against the production server.
Connection string for prod will need to be updated in appsettings.json (or via environment variable).
Do NOT run migrations against production — use the hand-written SQL scripts from Phase 1.
