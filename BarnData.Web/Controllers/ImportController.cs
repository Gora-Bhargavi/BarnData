using BarnData.Core.Services;
using BarnData.Data.Entities;
using BarnData.Web.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.Text.Json;
namespace BarnData.Web.Controllers
{
    public class ImportController : Controller
    {
        private readonly IAnimalService _animalService;
        private readonly IVendorService _vendorService;

        // KillDate placeholder — 2000-01-02 means "not yet killed" in the export
        private static readonly DateTime PENDING_DATE = new DateTime(2000, 1, 2);

        public ImportController(IAnimalService animalService, IVendorService vendorService)
        {
            _animalService = animalService;
            _vendorService = vendorService;
        }

        // ── SALE BILL IMPORT — GET ────────────────────────────────────────
        public IActionResult SaleBill()
        {
            return View();
        }

        // ── SALE BILL IMPORT — POST ───────────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SaleBill(IFormFile? file)
        {
            if (file == null || file.Length == 0)
            {
                ModelState.AddModelError("", "Please select an Excel file.");
                return View();
            }

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".xlsx")
            {
                ModelState.AddModelError("", "Only .xlsx files are supported. Please save as .xlsx in Excel first.");
                return View();
            }

            var vm      = new SaleBillImportViewModel { ImportedFile = file.FileName };
            var toImport = new List<Animal>();
            var vendors  = (await _vendorService.GetAllActiveAsync()).ToList();

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using var wb = new XLWorkbook(stream);
                var ws = wb.Worksheets.First();

                // ── Map headers by name ───────────────────────────────────
                var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 30;

                for (int c = 1; c <= lastCol; c++)
                {
                    var h = ws.Cell(1, c).GetString().Trim()
                               .Replace(":", "").ToLowerInvariant();
                    if (!string.IsNullOrEmpty(h))
                        colMap[h] = c;
                }

                // Helper to get column number by possible names
                int Col(params string[] names)
                {
                    foreach (var n in names)
                        if (colMap.TryGetValue(n.ToLowerInvariant(), out int c)) return c;
                    return -1;
                }

                // Map all columns
                int colAnimalType    = Col("animal type");
                int colTag1          = Col("tag number one", "tag one", "tag 1");
                int colTag2          = Col("tag number two", "tag two", "tag 2");
                int colTag3          = Col("tag 3", "tag3");
                int colPurchDate     = Col("purchase date");
                int colPurchType     = Col("purchase type");
                int colVendor        = Col("vendor");
                int colLiveWeight    = Col("live weight");
                int colLiveRate      = Col("live rate");
                int colKillDate      = Col("kill date");
                int colHotWeight     = Col("hot weight");
                int colGrade         = Col("grade");
                int colHS            = Col("h s", "hs", "health score");
                int colComment       = Col("comments", "comment");
                int colACN           = Col("animal control number");
                int colOfficeUse2    = Col("office use 2");
                int colState         = Col("state");
                int colBuyer         = Col("buyer");
                int colAnimalType2   = Col("animal type 2");
                int colVetName       = Col("vet name");

                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
                var billRef = $"IMPORT_{DateTime.Now:yyyyMMdd_HHmmss}";

                for (int row = 2; row <= lastRow; row++)
                {
                    // Get tag — skip blank rows
                    var tag1 = colTag1 > 0 ? ws.Cell(row, colTag1).GetString().Trim() : "";
                    if (string.IsNullOrEmpty(tag1)) continue;

                    vm.TotalRows++;

                    // ── Vendor ────────────────────────────────────────────
                    var vendorName = colVendor > 0
                        ? ws.Cell(row, colVendor).GetString().Trim()
                        : "";
                    if (string.IsNullOrEmpty(vendorName)) { vm.Skipped++; continue; }

                    // Find or auto-create vendor
                    var vendor = vendors.FirstOrDefault(v =>
                        v.VendorName.Equals(vendorName, StringComparison.OrdinalIgnoreCase));
                    int vendorId;
                    if (vendor == null)
                    {
                        vendorId = await _vendorService.GetOrCreateAsync(vendorName);
                        vendors  = (await _vendorService.GetAllActiveAsync()).ToList();
                    }
                    else vendorId = vendor.VendorID;

                    // ── Purchase type ─────────────────────────────────────
                    var purchType = colPurchType > 0
                        ? ws.Cell(row, colPurchType).GetString().Trim()
                        : "Sale Bill";
                    if (purchType.ToLower().Contains("consignment"))
                        purchType = "Consignment Bill";
                    else
                        purchType = "Sale Bill";

                    // ── Purchase date ─────────────────────────────────────
                    DateTime purchDate = DateTime.Today;
                    if (colPurchDate > 0)
                    {
                        var cell = ws.Cell(row, colPurchDate);
                        if (cell.DataType == XLDataType.DateTime)
                            purchDate = cell.GetDateTime();
                        else
                            DateTime.TryParse(cell.GetString(), out purchDate);
                    }

                    // ── Kill date — 2000-01-02 means pending (not killed yet) ──
                    DateTime? killDate = null;
                    if (colKillDate > 0)
                    {
                        var cell = ws.Cell(row, colKillDate);
                        DateTime kd = DateTime.MinValue;
                        if (cell.DataType == XLDataType.DateTime)
                            kd = cell.GetDateTime();
                        else
                            DateTime.TryParse(cell.GetString(), out kd);

                        // Only set kill date if it's a real date (not the 2000-01-02 placeholder)
                        if (kd > PENDING_DATE && kd.Year > 2000)
                            killDate = kd;
                    }

                    // ── Numeric fields ────────────────────────────────────
                    decimal GetDecimal(int col)
                    {
                        if (col < 0) return 0;
                        var v = ws.Cell(row, col).GetString()
                                   .Replace("$", "").Replace(",", "").Trim();
                        return decimal.TryParse(v, out var d) ? d : 0;
                    }

                    decimal liveWeight = GetDecimal(colLiveWeight);
                    decimal liveRate   = GetDecimal(colLiveRate);
                    decimal hotWeight  = GetDecimal(colHotWeight);

                    // ── Hot weight — 0 means not yet measured ─────────────
                    decimal? hotWt = hotWeight > 0 ? hotWeight : null;

                    // ── Grade — trim whitespace ───────────────────────────
                    var grade = colGrade > 0
                        ? ws.Cell(row, colGrade).GetString().Trim()
                        : null;
                    if (string.IsNullOrEmpty(grade)) grade = null;

                    // ── Health score ──────────────────────────────────────
                    int? hs = null;
                    if (colHS > 0)
                    {
                        var hsStr = ws.Cell(row, colHS).GetString().Trim();
                        if (int.TryParse(hsStr, out var hsVal) && hsVal > 0)
                            hs = hsVal;
                    }

                    // ── Animal type ───────────────────────────────────────
                    var animalType = colAnimalType > 0
                        ? ws.Cell(row, colAnimalType).GetString().Trim()
                        : "Cow";
                    if (string.IsNullOrEmpty(animalType)) animalType = "Cow";
                    // Normalize Str/str → Steer
                    if (animalType.StartsWith("Str", StringComparison.OrdinalIgnoreCase))
                        animalType = "Steer";

                    // ── Comments — check for condemned ────────────────────
                    var comment = colComment > 0
                        ? ws.Cell(row, colComment).GetString().Trim()
                        : null;
                    bool isCond = !string.IsNullOrEmpty(comment) &&
                                  comment.ToLower().Contains("cond");
                    if (string.IsNullOrEmpty(comment)) comment = null;

                    // ── Program code from vendor name ─────────────────────
                    var progCode = vendorName.ToUpper().Contains("ABF") ? "ABF" : "REG";

                    // ── Kill status ───────────────────────────────────────
                    var killStatus = killDate.HasValue ? "Killed" : "Pending";

                    // ── Build animal ──────────────────────────────────────
                    var animal = new Animal
                    {
                        VendorID             = vendorId,
                        TagNumber1           = tag1,
                        TagNumber2           = colTag2 > 0
                            ? NullIfEmpty(ws.Cell(row, colTag2).GetString().Trim())
                            : null,
                        Tag3                 = colTag3 > 0
                            ? NullIfEmpty(ws.Cell(row, colTag3).GetString().Trim())
                            : null,
                        AnimalType           = animalType,
                        AnimalType2          = colAnimalType2 > 0
                            ? NullIfEmpty(ws.Cell(row, colAnimalType2).GetString().Trim())
                            : null,
                        ProgramCode          = progCode,
                        PurchaseDate         = purchDate,
                        PurchaseType         = purchType,
                        LiveWeight           = liveWeight,
                        LiveRate             = liveRate,
                        KillDate             = killDate,
                        HotWeight            = hotWt,
                        Grade                = grade,
                        HealthScore          = hs,
                        Comment              = comment,
                        AnimalControlNumber  = colACN > 0
                            ? NullIfEmpty(ws.Cell(row, colACN).GetString().Trim())
                            : null,
                        OfficeUse2           = colOfficeUse2 > 0
                            ? NullIfEmpty(ws.Cell(row, colOfficeUse2).GetString().Trim())
                            : null,
                        State                = colState > 0
                            ? NullIfEmpty(ws.Cell(row, colState).GetString().Trim())
                            : null,
                        BuyerName            = colBuyer > 0
                            ? NullIfEmpty(ws.Cell(row, colBuyer).GetString().Trim())
                            : null,
                        VetName              = colVetName > 0
                            ? NullIfEmpty(ws.Cell(row, colVetName).GetString().Trim())
                            : null,
                        IsCondemned          = isCond,
                        KillStatus           = killStatus,
                        SaleBillRef          = billRef,
                    };

                    // Preview row
                    vm.Preview.Add(new SaleBillPreviewRow
                    {
                        VendorName   = vendorName,
                        Tag1         = tag1,
                        Tag2         = animal.TagNumber2,
                        AnimalType   = animalType,
                        LiveWeight   = liveWeight,
                        LiveRate     = liveRate,
                        PurchaseType = purchType,
                        Comment      = comment,
                        IsCondemned  = isCond,
                        Status       = "OK"
                    });

                    toImport.Add(animal);
                }

                // ── Bulk import ───────────────────────────────────────────
                var (imported, skipped, errors) = await _animalService.BulkImportAsync(toImport);
                vm.Imported  = imported;
                vm.Skipped  += skipped;
                vm.Errors    = errors;

                TempData["SuccessMessage"] =
                    $"Import complete: {imported} animals imported, {vm.Skipped} skipped.";
            }
            catch (Exception ex)
            {
                vm.Errors.Add($"Import failed: {ex.Message}");
                ModelState.AddModelError("", $"Import error: {ex.Message}");
            }

            return View("SaleBillResult", vm);
        }

        // ── MARK AS KILLED ────────────────────────────────────────────────
        public async Task<IActionResult> MarkKilled(int? vendorId)
        {
            var vendors = await _vendorService.GetAllActiveAsync();
            var pending = await _animalService.GetPendingAsync(vendorId);

            var vm = new MarkKilledViewModel
            {
                KillDate   = DateTime.Today,
                VendorId   = vendorId,
                VendorList = vendors.Select(v =>
                    new SelectListItem(v.VendorName, v.VendorID.ToString())),
                Animals = pending.Select(a => new PendingAnimalRow
                {
                    ControlNo           = a.ControlNo,
                    VendorName          = a.Vendor?.VendorName ?? "",
                    Tag1                = a.TagNumber1,
                    Tag2                = a.TagNumber2,
                    Tag3                = a.Tag3,
                    AnimalType          = a.AnimalType,
                    AnimalType2         = a.AnimalType2,
                    LiveWeight          = a.LiveWeight,
                    LiveRate            = a.LiveRate,
                    PurchaseType        = a.PurchaseType,
                    PurchaseDate        = a.PurchaseDate,
                    AnimalControlNumber = a.AnimalControlNumber,
                    Comment             = a.Comment,
                    State               = a.State,
                    BuyerName           = a.BuyerName,
                    VetName             = a.VetName,
                    OfficeUse2          = a.OfficeUse2,
                    ProgramCode         = a.ProgramCode,
                    Selected            = false,
                }).ToList()
            };

            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> MarkKilled(IFormCollection form)
        {
            if (!DateTime.TryParse(form["killDate"], out var killDate))
                killDate = DateTime.Today;

            // Collect selected control numbers
            var selectedIds = form["selectedIds"]
                .Where(v => !string.IsNullOrEmpty(v))
                .Select(v => int.TryParse(v, out var id) ? id : 0)
                .Where(id => id > 0)
                .Distinct()
                .ToList();

            if (!selectedIds.Any())
            {
                TempData["ErrorMessage"] = "No animals selected. Please check at least one animal.";
                return RedirectToAction(nameof(MarkKilled));
            }

            // Build per-animal kill data from form fields
            var animalData = selectedIds.Select(id => new KillAnimalData
            {
                ControlNo   = id,
                HotWeight   = decimal.TryParse(form[$"hotWeight_{id}"], out var hw) && hw > 0 ? hw : null,
                Grade       = form[$"grade_{id}"].FirstOrDefault(),
                HealthScore = int.TryParse(form[$"healthScore_{id}"], out var hs) && hs > 0 ? hs : null,
                IsCondemned = form[$"condemned_{id}"].Any(v => v == "true" || v == "on"),
            }).ToList();

            int count = await _animalService.MarkKilledWithDataAsync(animalData, killDate);

            TempData["SuccessMessage"] =
                $"{count} animals marked as killed on {killDate:MM/dd/yyyy}. Tally report is ready.";

            return RedirectToAction("Tally", "Report",
                new { killDate = killDate.ToString("yyyy-MM-dd") });
        }

        // ── Helper ────────────────────────────────────────────────────────
        private static string? NullIfEmpty(string? s)
            => string.IsNullOrWhiteSpace(s) ? null : s;

        // ── EXCEL IMPORT — GET ────────────────────────────────────────────
        public IActionResult Excel()
        {
            return View();
        }

        // ── EXCEL IMPORT — POST (parse & preview) ─────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Excel(IFormFile? file)
        {
            if (file == null || file.Length == 0)
            {
                ModelState.AddModelError("", "Please select an Excel file.");
                return View();
            }

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".xlsx")
            {
                ModelState.AddModelError("", "Only .xlsx files are supported.");
                return View();
            }

            var vm      = new ExcelImportViewModel { FileName = file.FileName };
            var vendors = (await _vendorService.GetAllActiveAsync()).ToList();

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using var wb = new XLWorkbook(stream);
                var ws = wb.Worksheets.First();

                // Map headers by name
                var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 30;
                for (int c = 1; c <= lastCol; c++)
                {
                    var h = ws.Cell(1, c).GetString().Trim().Replace(":", "").ToLowerInvariant();
                    if (!string.IsNullOrEmpty(h)) colMap[h] = c;
                }

                int Col(params string[] names)
                {
                    foreach (var n in names)
                        if (colMap.TryGetValue(n.ToLowerInvariant(), out int c)) return c;
                    return -1;
                }

                int colAnimalType  = Col("animal type");
                int colTag1        = Col("tag number one", "tag one", "tag 1");
                int colTag2        = Col("tag number two", "tag two", "tag 2");
                int colTag3        = Col("tag 3", "tag3");
                int colPurchDate   = Col("purchase date");
                int colPurchType   = Col("purchase type");
                int colVendor      = Col("vendor");
                int colLiveWeight  = Col("live weight");
                int colLiveRate    = Col("live rate");
                int colKillDate    = Col("kill date");
                int colHotWeight   = Col("hot weight");
                int colGrade       = Col("grade");
                int colHS          = Col("h s", "hs", "health score");
                int colComment     = Col("comments", "comment");
                int colACN         = Col("animal control number");
                int colOfficeUse2  = Col("office use 2");
                int colState       = Col("state");
                int colBuyer       = Col("buyer");
                int colAnimalType2 = Col("animal type 2");
                int colVetName     = Col("vet name");

                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

                decimal GetDecimal(int col, int row)
                {
                    if (col < 0) return 0;
                    var v = ws.Cell(row, col).GetString().Replace("$", "").Replace(",", "").Trim();
                    return decimal.TryParse(v, out var d) ? d : 0;
                }

                for (int row = 2; row <= lastRow; row++)
                {
                    var tag1 = colTag1 > 0 ? ws.Cell(row, colTag1).GetString().Trim() : "";
                    if (string.IsNullOrEmpty(tag1)) continue;

                    vm.TotalRows++;

                    var vendorName = colVendor > 0 ? ws.Cell(row, colVendor).GetString().Trim() : "";
                    if (string.IsNullOrEmpty(vendorName))
                    {
                        vm.Rows.Add(new ExcelPreviewRow { RowNum = row, TagNumber1 = tag1, Status = "Error", StatusNote = "Missing vendor" });
                        continue;
                    }

                    // Purchase type
                    var purchTypeRaw = colPurchType > 0 ? ws.Cell(row, colPurchType).GetString().Trim() : "Sale Bill";
                    var purchType    = purchTypeRaw.ToLower().Contains("consignment") ? "Consignment Bill" : "Sale Bill";

                    // Purchase date
                    DateTime purchDate = DateTime.Today;
                    if (colPurchDate > 0)
                    {
                        var cell = ws.Cell(row, colPurchDate);
                        if (cell.DataType == XLDataType.DateTime) purchDate = cell.GetDateTime();
                        else DateTime.TryParse(cell.GetString(), out purchDate);
                    }

                    // Kill date
                    DateTime? killDate = null;
                    if (colKillDate > 0)
                    {
                        var cell = ws.Cell(row, colKillDate);
                        DateTime kd = DateTime.MinValue;
                        if (cell.DataType == XLDataType.DateTime) kd = cell.GetDateTime();
                        else DateTime.TryParse(cell.GetString(), out kd);
                        if (kd > PENDING_DATE && kd.Year > 2000) killDate = kd;
                    }

                    // Animal type
                    var animalType = colAnimalType > 0 ? ws.Cell(row, colAnimalType).GetString().Trim() : "Cow";
                    if (string.IsNullOrEmpty(animalType)) animalType = "Cow";
                    if (animalType.StartsWith("Str", StringComparison.OrdinalIgnoreCase)) animalType = "Steer";

                    // Comment / condemned
                    var comment = colComment > 0 ? ws.Cell(row, colComment).GetString().Trim() : null;
                    bool isCond = !string.IsNullOrEmpty(comment) && comment.ToLower().Contains("cond");
                    if (string.IsNullOrEmpty(comment)) comment = null;

                    // Health score
                    int? hs = null;
                    if (colHS > 0)
                    {
                        var hsStr = ws.Cell(row, colHS).GetString().Trim();
                        if (int.TryParse(hsStr, out var hsVal) && hsVal > 0) hs = hsVal;
                    }

                    decimal hotWtRaw = GetDecimal(colHotWeight, row);
                    decimal liveWeight = GetDecimal(colLiveWeight, row);
                    decimal liveRate   = GetDecimal(colLiveRate, row);

                    // Check duplicate in existing DB (preview only — don't save yet)
                    var vendor    = vendors.FirstOrDefault(v => v.VendorName.Equals(vendorName, StringComparison.OrdinalIgnoreCase));
                    int vendorId  = vendor?.VendorID ?? 0;
                    string status = "OK";
                    string? note  = null;

                    if (vendorId > 0)
                    {
                        bool isDup = await _animalService.IsTagDuplicateAsync(tag1, vendorId);
                        if (isDup) { status = "Duplicate"; note = "Tag already exists for this vendor"; }
                    }

                    vm.Rows.Add(new ExcelPreviewRow
                    {
                        RowNum              = row,
                        VendorName          = vendorName,
                        TagNumber1          = tag1,
                        TagNumber2          = colTag2 > 0 ? NullIfEmpty(ws.Cell(row, colTag2).GetString().Trim()) : null,
                        Tag3                = colTag3 > 0 ? NullIfEmpty(ws.Cell(row, colTag3).GetString().Trim()) : null,
                        AnimalType          = animalType,
                        AnimalType2         = colAnimalType2 > 0 ? NullIfEmpty(ws.Cell(row, colAnimalType2).GetString().Trim()) : null,
                        PurchaseType        = purchType,
                        PurchaseDate        = purchDate,
                        LiveWeight          = liveWeight,
                        LiveRate            = liveRate,
                        KillDate            = killDate,
                        HotWeight           = hotWtRaw > 0 ? hotWtRaw : null,
                        Grade               = colGrade > 0 ? NullIfEmpty(ws.Cell(row, colGrade).GetString().Trim()) : null,
                        HealthScore         = hs,
                        Comment             = comment,
                        AnimalControlNumber = colACN > 0 ? NullIfEmpty(ws.Cell(row, colACN).GetString().Trim()) : null,
                        OfficeUse2          = colOfficeUse2 > 0 ? NullIfEmpty(ws.Cell(row, colOfficeUse2).GetString().Trim()) : null,
                        State               = colState > 0 ? NullIfEmpty(ws.Cell(row, colState).GetString().Trim()) : null,
                        BuyerName           = colBuyer > 0 ? NullIfEmpty(ws.Cell(row, colBuyer).GetString().Trim()) : null,
                        VetName             = colVetName > 0 ? NullIfEmpty(ws.Cell(row, colVetName).GetString().Trim()) : null,
                        IsCondemned         = isCond,
                        Status              = status,
                        StatusNote          = note,
                    });
                }
            }
            catch (Exception ex)
            {
                vm.Errors.Add($"Parse error: {ex.Message}");
                return View(vm);
            }

            // Store preview in TempData as JSON for confirm step
            TempData["ExcelPreview"] = System.Text.Json.JsonSerializer.Serialize(vm);
            return View("ExcelPreview", vm);
        }

        // ── EXCEL IMPORT — CONFIRM & SAVE ─────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ExcelConfirm()
        {
            var json = TempData["ExcelPreview"] as string;
            if (string.IsNullOrEmpty(json))
            {
                TempData["ErrorMessage"] = "Preview session expired. Please re-upload the file.";
                return RedirectToAction(nameof(Excel));
            }

            var vm = System.Text.Json.JsonSerializer.Deserialize<ExcelImportViewModel>(json);
            if (vm == null || !vm.Rows.Any())
            {
                TempData["ErrorMessage"] = "No data to import.";
                return RedirectToAction(nameof(Excel));
            }

            var vendors    = (await _vendorService.GetAllActiveAsync()).ToList();
            var toImport   = new List<Animal>();
            var billRef    = $"EXCEL_{DateTime.Now:yyyyMMdd_HHmmss}";

            foreach (var r in vm.Rows.Where(r => r.Status == "OK"))
            {
                var vendor = vendors.FirstOrDefault(v =>
                    v.VendorName.Equals(r.VendorName, StringComparison.OrdinalIgnoreCase));
                int vendorId;
                if (vendor == null)
                {
                    vendorId = await _vendorService.GetOrCreateAsync(r.VendorName);
                    vendors  = (await _vendorService.GetAllActiveAsync()).ToList();
                }
                else vendorId = vendor.VendorID;

                toImport.Add(new Animal
                {
                    VendorID            = vendorId,
                    TagNumber1          = r.TagNumber1,
                    TagNumber2          = r.TagNumber2,
                    Tag3                = r.Tag3,
                    AnimalType          = r.AnimalType,
                    AnimalType2         = r.AnimalType2,
                    ProgramCode         = r.VendorName.ToUpper().Contains("ABF") ? "ABF" : "REG",
                    PurchaseDate        = r.PurchaseDate,
                    PurchaseType        = r.PurchaseType,
                    LiveWeight          = r.LiveWeight,
                    LiveRate            = r.LiveRate,
                    KillDate            = r.KillDate,
                    HotWeight           = r.HotWeight,
                    Grade               = r.Grade,
                    HealthScore         = r.HealthScore,
                    Comment             = r.Comment,
                    AnimalControlNumber = r.AnimalControlNumber,
                    OfficeUse2          = r.OfficeUse2,
                    State               = r.State,
                    BuyerName           = r.BuyerName,
                    VetName             = r.VetName,
                    IsCondemned         = r.IsCondemned,
                    KillStatus          = r.KillDate.HasValue ? "Killed" : "Pending",
                    SaleBillRef         = billRef,
                });
            }

            var (imported, skipped, errors) = await _animalService.BulkImportAsync(toImport);
            int dupeCount = vm.Rows.Count(r => r.Status == "Duplicate");

            TempData["SuccessMessage"] =
                $"Import complete: {imported} animals saved, {skipped + dupeCount} skipped.";

            return RedirectToAction("Index", "Animal", new { status = "pending" });
        }
    }
}
