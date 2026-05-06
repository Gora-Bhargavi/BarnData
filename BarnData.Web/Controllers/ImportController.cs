using BarnData.Core.Services;
using BarnData.Core.Validation;
using BarnData.Data.Entities;
using BarnData.Web.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;
using System.Globalization;
using System.Text.Json;
using Microsoft.Extensions.Logging;
namespace BarnData.Web.Controllers
{
    public class ImportController : Controller
    {
        private readonly IAnimalService _animalService;
        private readonly IVendorService _vendorService;
        private readonly IAnimalQueryService _animalQueryService;
        private readonly IImportStagingService _stagingService;
        private readonly ILogger<ImportController> _logger;
        private readonly IConfiguration _configuration;

        public ImportController(IAnimalService animalService, IVendorService vendorService,
                                 IAnimalQueryService animalQueryService,
                                 IImportStagingService stagingService,
                                 ILogger<ImportController> logger,
                                 IConfiguration configuration)
        {
            _animalService = animalService;
            _vendorService = vendorService;
            _animalQueryService = animalQueryService;
            _stagingService = stagingService;
            _logger = logger;
            _configuration = configuration;
        }

        //  SALE BILL IMPORT — GET 
        public IActionResult SaleBill()
        {
            return View();
        }

        private async Task<HotWeightImportViewModel?> ReadHotWeightPreviewVmAsync()
        {
            // HW staging is shared across all users (see StagingBridge.SharedHotWeightKey).
            // Per-user Session/TempData are NOT consulted because they would shadow the
            // collaborative state any teammate may have just updated.
            var json = await StagingBridge.ReadSharedAsync(
                _stagingService,
                StagingBridge.Types.HotWeight,
                StagingBridge.SharedHotWeightKey);

        if (string.IsNullOrWhiteSpace(json)) return null;

        try
        {
            return System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(json);
        }
        catch
        {
            return null;
        }
    }

private async Task PersistHotWeightPreviewVmAsync(HotWeightImportViewModel vm)
{
    var updatedJson = System.Text.Json.JsonSerializer.Serialize(vm);

    await StagingBridge.WriteSharedAsync(
        _stagingService,
        StagingBridge.Types.HotWeight,
        StagingBridge.SharedHotWeightKey,
        updatedJson,
        sourceFileName: null);
}

private static void ApplySavedValuesToPreviewRow(HotWeightPreviewRow row, AnimalRowDto dto)
{
    if (dto.HotWeight > 0) row.NewHotWeight = dto.HotWeight;
    if (!string.IsNullOrWhiteSpace(dto.Grade)) row.NewGrade = dto.Grade.Trim().ToUpperInvariant();
    if (dto.HealthScore > 0) row.NewHealthScore = dto.HealthScore;

    if (!string.IsNullOrWhiteSpace(dto.AnimalControlNumber))
    {
        var normalizedAcn = dto.AnimalControlNumber.Trim().TrimStart('0');
        row.NewAnimalControlNumber = normalizedAcn;
        row.AnimalControlNumber = normalizedAcn;
    }

    row.Status = "Loaded";
    row.FlagReason = "";
}

    private async Task SyncHotWeightPreviewAfterSaveAsync(IEnumerable<AnimalRowDto> savedRows)
    {
        var vm = await ReadHotWeightPreviewVmAsync();
        if (vm == null) return;

        var dataList = savedRows.Where(r => r.ControlNo > 0).ToList();
        if (dataList.Count == 0) return;

        bool changed = false;

        foreach (var dto in dataList)
        {
            var controlNo = dto.ControlNo;
            var origNo    = dto.OriginalControlNo > 0 ? dto.OriginalControlNo : controlNo;

            // Try to find the flagged row by either its original ControlNo OR the saved/picked ControlNo
            var flagged = vm.FlaggedRows.FirstOrDefault(r => r.ControlNo == origNo)
                    ?? vm.FlaggedRows.FirstOrDefault(r => r.ControlNo == controlNo);

            if (flagged != null)
            {
                ApplySavedValuesToPreviewRow(flagged, dto);
                vm.FlaggedRows.Remove(flagged);

                var autoExisting = vm.AutoRows.FirstOrDefault(r => r.ControlNo == controlNo);
                if (autoExisting == null)
                    vm.AutoRows.Add(flagged);
                else
                    ApplySavedValuesToPreviewRow(autoExisting, dto);

                changed = true;
                continue;
            }

            var auto = vm.AutoRows.FirstOrDefault(r => r.ControlNo == controlNo);
            if (auto != null)
            {
                ApplySavedValuesToPreviewRow(auto, dto);
                changed = true;
            }
        }

        if (changed)
            await PersistHotWeightPreviewVmAsync(vm);
    }

        //  SALE BILL IMPORT — POST 
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
            if (!string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase))
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

                //  Map headers by name 
                var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 30;

                for (int c = 1; c <= lastCol; c++)
                {
                    var h = ws.Cell(1, c).GetString().Trim()
                               .Replace(":", "").ToLowerInvariant();
                    if (!string.IsNullOrEmpty(h))
                        colMap[h] = c;
                }

                // Helper: read any cell as string regardless of data type
                /*string GetCellString(IXLCell cell)
                {
                    if (cell == null) return "";
                    try
                    {
                        // Numeric stored as number — convert to string without decimal
                        if (cell.DataType == XLDataType.Number)
                        {
                            var d = cell.GetDouble();
                            return d == Math.Floor(d)
                                ? ((long)d).ToString()
                                : d.ToString();
                        }
                        if (cell.DataType == XLDataType.Text)   return cell.GetString().Trim();
                        if (cell.DataType == XLDataType.Boolean) return cell.GetBoolean().ToString();
                        // Formula or other — try cached value
                        var v = cell.CachedValue;
                        //return v != null ? v.ToString()?.Trim() ?? "" : cell.GetString().Trim();
                        var value = v.ToString()?.Trim();
                        return string.IsNullOrEmpty(value) ? cell.GetString().Trim() : value;
                    }
                    catch { return cell.GetString().Trim(); }
                }*/

                // Helper: read date from cell handling formulas and text
                /*DateTime? GetCellDate(IXLCell cell)
                {
                    if (cell == null) return null;
                    try
                    {
                        if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();
                        // Try cached value for formula cells
                        //var raw = cell.CachedValue?.ToString() ?? cell.GetString();
                        var raw = cell.CachedValue.ToString();
                        if(string.IsNullOrEmpty(raw)) raw = cell.GetString();
                        if (DateTime.TryParse(raw, out var dt)) return dt;
                    }
                    catch { }
                    return null;
                }*/

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
                    var tag1 = colTag1 > 0 ? GetCellString(ws.Cell(row, colTag1)) : "";
                    if (string.IsNullOrEmpty(tag1)) continue;

                    vm.TotalRows++;

                    // Vendor 
                    var vendorName = colVendor > 0
                        ? GetCellString(ws.Cell(row, colVendor))
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

                    //  Purchase type 
                    var purchType = colPurchType > 0
                        ? GetCellString(ws.Cell(row, colPurchType))
                        : "Sale Bill";
                    if (purchType.ToLower().Contains("consignment"))
                        purchType = "Consignment Bill";
                    else
                        purchType = "Sale Bill";

                    //  Purchase date 
                    DateTime purchDate = DateTime.Today;
                    if (colPurchDate > 0)
                    {
                        var pd = GetCellDate(ws.Cell(row, colPurchDate));
                        if (pd.HasValue) purchDate = pd.Value;
                    }

                    // Kill date is always empty at import time.

                    //  Numeric fields 
                    decimal GetDecimal(int col)
                    {
                        if (col < 0) return 0;
                        var cell = ws.Cell(row, col);
                        if (cell.DataType == XLDataType.Number)
                            return (decimal)cell.GetDouble();
                        var v = GetCellString(cell).Replace("$", "").Replace(",", "").Trim();
                        return decimal.TryParse(v, out var d) ? d : 0;
                    }

                    decimal liveWeight = GetDecimal(colLiveWeight);
                    decimal liveRate   = GetDecimal(colLiveRate);
                    decimal hotWeight  = GetDecimal(colHotWeight);

                    //  Hot weight — 0 means not yet measured 
                    decimal? hotWt = hotWeight > 0 ? hotWeight : null;

                    //  Grade - trim whitespace 
                    var grade = colGrade > 0 ? GetCellString(ws.Cell(row, colGrade)) : null;
                    if (string.IsNullOrEmpty(grade)) grade = null;

                    //  Health score 
                    int? hs = null;
                    if (colHS > 0)
                    {
                        var hsStr = GetCellString(ws.Cell(row, colHS));
                        if (int.TryParse(hsStr, out var hsVal) && hsVal > 0) hs = hsVal;
                    }

                    //  Animal type 
                    var animalType = colAnimalType > 0 ? GetCellString(ws.Cell(row, colAnimalType)) : "Cow";
                    if (string.IsNullOrEmpty(animalType)) animalType = "Cow";
                    if (animalType.StartsWith("Str", StringComparison.OrdinalIgnoreCase)) animalType = "Steer";

                    //  Comments — check for condemned 
                    var comment = colComment > 0 ? GetCellString(ws.Cell(row, colComment)) : null;
                    bool isCond = !string.IsNullOrEmpty(comment) && comment.ToLower().Contains("cond");
                    if (string.IsNullOrEmpty(comment)) comment = null;

                    //  Program code from vendor name 
                    var progCode = vendorName.ToUpper().Contains("ABF") ? "ABF" : "REG";

                    // Kill status is always pending at import time.
                    var killStatus = "Pending";

                    //  Build animal 
                    var animal = new Animal
                    {
                        VendorID             = vendorId,
                        TagNumber1           = tag1,
                        TagNumber2           = colTag2 > 0
                            ? NullIfEmpty(GetCellString(ws.Cell(row, colTag2)))
                            : null,
                        Tag3                 = colTag3 > 0
                            ? NullIfEmpty(GetCellString(ws.Cell(row, colTag3)))
                            : null,
                        AnimalType           = animalType,
                        AnimalType2          = colAnimalType2 > 0
                            ? NullIfEmpty(GetCellString(ws.Cell(row, colAnimalType2)))
                            : null,
                        ProgramCode          = progCode,
                        PurchaseDate         = purchDate,
                        PurchaseType         = purchType,
                        LiveWeight           = liveWeight,
                        LiveRate             = liveRate,
                        KillDate             = null,
                        HotWeight            = hotWt,
                        Grade                = grade,
                        HealthScore          = hs,
                        Comment              = comment,
                        AnimalControlNumber  = colACN > 0
                            ? NullIfEmpty(GetCellString(ws.Cell(row, colACN)))
                            : null,
                        OfficeUse2           = colOfficeUse2 > 0
                            ? NullIfEmpty(GetCellString(ws.Cell(row, colOfficeUse2)))
                            : null,
                        State                = colState > 0
                            ? NullIfEmpty(GetCellString(ws.Cell(row, colState)))
                            : null,
                        BuyerName            = colBuyer > 0
                            ? NullIfEmpty(GetCellString(ws.Cell(row, colBuyer)))
                            : null,
                        VetName              = colVetName > 0
                            ? NullIfEmpty(GetCellString(ws.Cell(row, colVetName)))
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

                //  Bulk import 
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

        
        // AJAX: Save edits (only edited rows - no form size limit) 
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SaveEditsApi([FromBody] SaveEditsRequest req)
        {
            _logger.LogInformation("[SAVE-API] Called. req={ReqNull}, Rows={Count}",
                req == null ? "NULL" : "OK", req?.Rows?.Count ?? 0);
            if (req?.Rows == null || !req.Rows.Any())
            {
                _logger.LogWarning("[SAVE-API] No rows to save.");
                return Json(new { success = false, message = "No edits to save." });
            }

            // Validate each row individually and collect error details
            var rowValidationResults = req.Rows
                .Select(r => new
                {
                    Row = r,
                    Error = ValidateMarkKilledRow(r)
                })
                .ToList();

            var validRows = rowValidationResults.Where(x => string.IsNullOrWhiteSpace(x.Error)).Select(x => x.Row).ToList();
            var invalidRows = rowValidationResults
                .Where(x => !string.IsNullOrWhiteSpace(x.Error))
                .Select(x => new
                {
                    id = x.Row.ControlNo,
                    controlNo = x.Row.ControlNo,
                    error = x.Error
                })
                .ToList();

            // Save valid rows
            int savedCount = 0;
            if (validRows.Any())
            {
                var animalData = validRows.Select(r => new KillAnimalData
                {
                    ControlNo           = r.ControlNo,
                    AnimalControlNumber = NormalizeAcn(r.AnimalControlNumber),
                    KillDate            = DateTime.TryParse(r.KillDate, out var kd) ? kd : (DateTime?)null,
                    LiveWeight          = r.LiveWeight > 0
                                            ? r.LiveWeight
                                            : (r.PurchaseType.Contains("consignment", StringComparison.OrdinalIgnoreCase) && r.HotWeight > 0
                                            ? r.HotWeight
                                            : (decimal?)null),
                    LiveRate = r.LiveRate,
                    HotWeight           = r.HotWeight > 0 ? r.HotWeight : (decimal?)null,
                    Grade               = string.IsNullOrWhiteSpace(r.Grade) ? null : r.Grade,
                    HealthScore         = r.HealthScore > 0 ? r.HealthScore : (int?)null,
                    IsCondemned         = r.IsCondemned,
                    State               = string.IsNullOrWhiteSpace(r.State) ? null : r.State,
                    VetName             = string.IsNullOrWhiteSpace(r.VetName) ? null : r.VetName,
                    OfficeUse2          = string.IsNullOrWhiteSpace(r.OfficeUse2) ? null : r.OfficeUse2,
                    Comment             = string.IsNullOrWhiteSpace(r.Comment) ? null : r.Comment,
                }).ToList();

                savedCount = await _animalService.SaveKillDataAsync(animalData);

                if (savedCount > 0)
                {
                    try
                    {
                        await SyncHotWeightPreviewAfterSaveAsync(validRows);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "[SAVE-API] Saved DB rows but failed to sync HWPreview state.");
                    }
                }
            }

            // Build response message
            string message = savedCount > 0 
                ? $"{savedCount} record{(savedCount != 1 ? "s" : "")} saved."
                : "";

            if (invalidRows.Any())
            {
                message += (message.Length > 0 ? " " : "") +
                        $"⚠️ {invalidRows.Count} row{(invalidRows.Count != 1 ? "s" : "")} have errors and were NOT saved.";
            }

            return Json(new
            {
                success = savedCount > 0, // True if at least some rows saved
                saved = savedCount,
                count = savedCount,
                failed = invalidRows.Count,
                invalidRows = invalidRows,
                message = message
            });
}

        // AJAX: Mark as killed (only selected rows - no form size limit) 
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> MarkKilledApi([FromBody] MarkKilledRequest req)
        {
            if (req?.Rows == null || !req.Rows.Any())
                return Json(new { success = false, message = "No animals selected to mark as killed." });

            if (!DateTime.TryParse(req.KillDate, out var killDate))
                killDate = DateTime.Today;

            bool IsCompleteForKill(AnimalRowDto r)
            {
                // ACN is always required.
                if (NormalizeAcn(r.AnimalControlNumber) == null) return false;

                // Condemned animals don't have carcass weight, grade, or HS.
                // The condemnation grade itself (X-code) is sufficient.
                if (r.IsCondemned) return true;

                // Production-killed animals need full weight, grade, and HS.
                return r.HotWeight > 0
                    && !string.IsNullOrWhiteSpace(r.Grade)
                    && r.HealthScore >= 1 && r.HealthScore <= 5;
            }

            // ---- CHANGED: validate per-row; split into valid / invalid ----
            var rowResults = req.Rows.Select(r =>
            {
                if (!IsCompleteForKill(r))
                {
                    var msg = r.IsCondemned
                        ? $"Ctrl No {r.ControlNo}: ACN is required even for condemned animals."
                        : $"Ctrl No {r.ControlNo}: missing required fields (ACN, Hot Wt, Grade, HS).";
                    return (Row: r, Error: msg);
                }
                var err = ValidateMarkKilledRow(r);
                return (Row: r, Error: err);
            }).ToList();

            var validRows   = rowResults.Where(x => string.IsNullOrWhiteSpace(x.Error)).Select(x => x.Row).ToList();
            var invalidRows = rowResults.Where(x => !string.IsNullOrWhiteSpace(x.Error))
                                        .Select(x => new { row = x.Row, x.Row.ControlNo, error = x.Error })
                                        .ToList();

            if (!validRows.Any())
            {
                // All rows failed — return the first few errors
                return Json(new
                {
                    success = false,
                    message = string.Join(" ", invalidRows.Take(3).Select(x => x.error)),
                    failed  = invalidRows
                });
            }
            // ---- END CHANGED ----

            var animalData = validRows.Select(r => new KillAnimalData
            {
                ControlNo           = r.ControlNo,
                AnimalControlNumber = NormalizeAcn(r.AnimalControlNumber),
                KillDate            = killDate,
                LiveWeight          = r.LiveWeight > 0
                                        ? r.LiveWeight
                                        : (r.HwImported
                                            && (r.PurchaseType ?? "").Contains("consignment", StringComparison.OrdinalIgnoreCase)
                                            && r.HotWeight > 0
                                                ? r.HotWeight
                                                : (decimal?)null),
                HotWeight           = r.HotWeight > 0 ? r.HotWeight : (decimal?)null,
                Grade               = string.IsNullOrWhiteSpace(r.Grade) ? null : r.Grade,
                HealthScore         = r.HealthScore > 0 ? r.HealthScore : (int?)null,
                IsCondemned         = r.IsCondemned,
                State               = string.IsNullOrWhiteSpace(r.State) ? null : r.State,
                VetName             = string.IsNullOrWhiteSpace(r.VetName) ? null : r.VetName,
                OfficeUse2          = string.IsNullOrWhiteSpace(r.OfficeUse2) ? null : r.OfficeUse2,
                Comment             = string.IsNullOrWhiteSpace(r.Comment) ? null : r.Comment,
            }).ToList();

            int count = await _animalService.MarkKilledWithDataAsync(animalData, killDate);

            // ---- CHANGED: partial success response ----
            bool hasInvalid = invalidRows.Any();
            string msg = $"{count} animal{(count != 1 ? "s" : "")} marked as killed on {killDate:MM/dd/yyyy}.";
            if (hasInvalid)
                msg += $" ⚠️ {invalidRows.Count} row{(invalidRows.Count != 1 ? "s" : "")} had errors and were skipped.";

            return Json(new
            {
                success  = true,
                partial  = hasInvalid,
                message  = msg,
                failed   = invalidRows,
                redirect = hasInvalid ? (string?)null : Url.Action("Tally", "Report", new { killDate = killDate.ToString("yyyy-MM-dd") })
            });
            // ---- END CHANGED ----
        }

        // AJAX: Mark ALL complete-for-kill animals as killed across every page.
        // ----------------------------------------------------------------------
        //
        // Operator clicks one button. Server enumerates the entire pending set
        // matching the current vendor + search filter (no pagination), applies
        // Hot Weight pre-fill from session, finds the rows that are
        // complete-for-kill, and saves them in one transaction.
        //
        // This replaces the per-page workflow that required operators to
        // navigate to each page and click "Mark selected as killed" on each.
        // After save, saved ControlNos are marked Loaded in session HW data so
        // they don't re-pre-fill on subsequent visits.
        //
        // Request body shape:
        //   { killDate: "2026-05-02", vendorIds: "12,33", q: "search term" }
        // Response shape:
        //   { success, savedCount, skippedCount, errors: [...], message }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> MarkAllCompleteAsKilled([FromBody] MarkAllCompleteRequest req)
        {
            if (!DateTime.TryParse(req?.KillDate, out var killDate))
                killDate = DateTime.Today;

            // 1. Parse vendor filter (same as MarkKilledFast)
            List<int> vendorIdList = new();
            if (!string.IsNullOrEmpty(req?.VendorIds))
            {
                vendorIdList = req.VendorIds
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(v => int.TryParse(v.Trim(), out var id) ? id : 0)
                    .Where(id => id > 0).ToList();
            }

            // 2. Pull ALL pending animals matching filter (cap 5000 for safety)
            var (pending, totalCount) = await _animalQueryService.GetPendingPagedAsync(
                vendorIdList.Count > 0 ? vendorIdList : null,
                page: 1,
                pageSize: 5000,
                searchTerm: string.IsNullOrWhiteSpace(req?.Q) ? null : req!.Q);

            _logger.LogInformation(
                "[MarkAllComplete] Scanning {Count} pending animals (totalMatching={Total}) for kill date {Date:yyyy-MM-dd}",
                pending.Count, totalCount, killDate);

            // 3. Build HW lookup from session (TempData is single-request only)
            HotWeightImportViewModel? hwVm = null;
            var hwLookup = new Dictionary<int, HotWeightPreviewRow>();
            var hwJson = HttpContext.Session.GetString("HWPreview");
            if (!string.IsNullOrEmpty(hwJson))
            {
                try
                {
                    hwVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(hwJson);
                    if (hwVm != null)
                    {
                        foreach (var r in hwVm.AutoRows.Where(r => r.ControlNo > 0 && r.NewHotWeight.HasValue))
                            hwLookup[r.ControlNo] = r;
                        foreach (var r in hwVm.FlaggedRows.Where(r => r.ControlNo > 0 && r.NewHotWeight.HasValue))
                            hwLookup[r.ControlNo] = r;
                        foreach (var r in hwVm.AutoRows.Where(r => r.ControlNo > 0 && !string.IsNullOrEmpty(r.NewAnimalControlNumber)))
                            if (!hwLookup.ContainsKey(r.ControlNo)) hwLookup[r.ControlNo] = r;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "[MarkAllComplete] Could not deserialize HWPreview JSON");
                }
            }

            // 4. For each pending animal, build the merged row and check completeness
            bool IsCompleteForKill(string? acn, decimal? hotWeight, string? grade, int? hs, bool isCondemned)
            {
                if (NormalizeAcn(acn) == null) return false;
                if (isCondemned) return true;
                return hotWeight.HasValue && hotWeight.Value > 0
                    && !string.IsNullOrWhiteSpace(grade)
                    && hs.HasValue && hs.Value >= 1 && hs.Value <= 5;
            }

            var killData = new List<KillAnimalData>();
            var skipped  = new List<object>();
            int hwSourcedRows = 0;
            int billOnlyRows  = 0;

            foreach (var a in pending)
            {
                hwLookup.TryGetValue(a.ControlNo, out var hw);

                // Apply HW pre-fill ONLY if the bill itself doesn't have the value.
                // This matches the MarkKilledFast view behavior where HW values
                // override bill values when present.
                var acn = (hw != null && !string.IsNullOrEmpty(hw.NewAnimalControlNumber))
                            ? hw.NewAnimalControlNumber
                            : a.AnimalControlNumber;
                var hotWeight = hw != null ? hw.NewHotWeight : a.HotWeight;
                var grade     = hw != null ? hw.NewGrade : a.Grade;
                var hs        = hw != null ? hw.NewHealthScore : a.HealthScore;
                var condemned = (hw != null && hw.IsCondemned) ? true : a.IsCondemned;

                if (!IsCompleteForKill(acn, hotWeight, grade, hs, condemned))
                {
                    // Don't add to skipped list — we only report rows where the operator
                    // would expect a save. A pending row with no HW data and incomplete
                    // bill data isn't an "error", it's just not ready yet.
                    continue;
                }

                if (hw != null) hwSourcedRows++; else billOnlyRows++;

                // Trim comment append for trim-variance rows
                string? mergedComment = a.Comment;
                if (hw != null && !string.IsNullOrEmpty(hw.TrimComment))
                {
                    mergedComment = string.IsNullOrEmpty(a.Comment)
                        ? hw.TrimComment
                        : a.Comment + " " + hw.TrimComment;
                }

                killData.Add(new KillAnimalData
                {
                    ControlNo           = a.ControlNo,
                    AnimalControlNumber = NormalizeAcn(acn),
                    KillDate            = killDate,
                    LiveWeight          = a.LiveWeight > 0
                                            ? a.LiveWeight
                                            : ((a.PurchaseType ?? "").Contains("consignment", StringComparison.OrdinalIgnoreCase)
                                                && hotWeight.HasValue && hotWeight.Value > 0
                                                    ? hotWeight
                                                    : (decimal?)null),
                    HotWeight           = hotWeight.HasValue && hotWeight.Value > 0 ? hotWeight : (decimal?)null,
                    Grade               = string.IsNullOrWhiteSpace(grade) ? null : grade,
                    HealthScore         = hs.HasValue && hs.Value > 0 ? hs : (int?)null,
                    IsCondemned         = condemned,
                    State               = string.IsNullOrWhiteSpace(a.State) ? null : a.State,
                    VetName             = string.IsNullOrWhiteSpace(a.VetName) ? null : a.VetName,
                    OfficeUse2          = string.IsNullOrWhiteSpace(a.OfficeUse2) ? null : a.OfficeUse2,
                    Comment             = string.IsNullOrWhiteSpace(mergedComment) ? null : mergedComment,
                });
            }

            if (killData.Count == 0)
            {
                return Json(new
                {
                    success = false,
                    savedCount = 0,
                    message = $"No complete-for-kill rows found across {totalCount} pending animals. Make sure Hot Weight data is loaded and rows have ACN, HotWeight, Grade, and HS."
                });
            }

            // 5. Save in one transaction
            int savedCount = 0;
            try
            {
                savedCount = await _animalService.MarkKilledWithDataAsync(killData, killDate);
                _logger.LogInformation(
                    "[MarkAllComplete] Saved {Saved} of {Attempted} rows (HW-sourced={HW}, bill-only={Bill}) for {Date:yyyy-MM-dd}",
                    savedCount, killData.Count, hwSourcedRows, billOnlyRows, killDate);
            }
            catch (Exception saveEx)
            {
                _logger.LogError(saveEx, "[MarkAllComplete] Save transaction failed");
                return Json(new
                {
                    success = false,
                    savedCount = 0,
                    message = "Save failed: " + saveEx.Message
                });
            }

            // 6. Update HW session: mark these ControlNos as "Loaded" so they don't
            //    re-pre-fill on subsequent visits.
            if (hwVm != null && savedCount > 0)
            {
                var savedSet = killData.Select(d => d.ControlNo).ToHashSet();
                int markedLoaded = 0;
                foreach (var r in hwVm.AutoRows.Concat(hwVm.FlaggedRows))
                {
                    if (savedSet.Contains(r.ControlNo) && r.Status != "Loaded")
                    {
                        r.Status = "Loaded";
                        markedLoaded++;
                    }
                }
                if (markedLoaded > 0)
                {
                    try
                    {
                        var updatedJson = System.Text.Json.JsonSerializer.Serialize(hwVm);
                        HttpContext.Session.SetString("HWPreview", updatedJson);
                        // Persist to shared staging too so other teammates see the update.
                        await StagingBridge.WriteSharedAsync(
                            _stagingService,
                            StagingBridge.Types.HotWeight,
                            StagingBridge.SharedHotWeightKey,
                            updatedJson,
                            sourceFileName: null);
                        _logger.LogInformation("[MarkAllComplete] Marked {Count} HW preview rows as Loaded.", markedLoaded);
                    }
                    catch (Exception updEx)
                    {
                        _logger.LogWarning(updEx, "[MarkAllComplete] Could not persist updated HW preview after save.");
                    }
                }
            }

            return Json(new
            {
                success    = true,
                savedCount,
                skippedCount = totalCount - savedCount,
                hwSourcedRows,
                billOnlyRows,
                message = $"{savedCount} animals marked as killed on {killDate:MM/dd/yyyy} ({hwSourcedRows} from Hot Weight, {billOnlyRows} from existing bill data). Reload to see updated list.",
                redirect = Url.Action("Tally", "Report", new { killDate = killDate.ToString("yyyy-MM-dd") })
            });
        }

        public class MarkAllCompleteRequest
        {
            public string? KillDate  { get; set; }
            public string? VendorIds { get; set; }
            public string? Q         { get; set; }
        }

        
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SaveAllHwData()
        {
            // 1. Read HW preview from session (TempData is single-request only)
            var hwJson = HttpContext.Session.GetString("HWPreview");
            if (string.IsNullOrEmpty(hwJson))
            {
                // Fall back to shared staging if session is empty
                hwJson = await StagingBridge.ReadSharedAsync(
                    _stagingService,
                    StagingBridge.Types.HotWeight,
                    StagingBridge.SharedHotWeightKey);
            }
            if (string.IsNullOrEmpty(hwJson))
            {
                return Json(new { success = false, message = "No Hot Weight session found. Refresh from Hot Scale first." });
            }

            HotWeightImportViewModel? hwVm = null;
            try
            {
                hwVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(hwJson);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[SaveAllHw] Could not deserialize HW preview JSON");
                return Json(new { success = false, message = "Could not parse Hot Weight session data." });
            }
            if (hwVm == null)
            {
                return Json(new { success = false, message = "Hot Weight session was empty." });
            }

            // 2. Collect all rows that have HW data to save and aren't already Loaded
            var rowsToPersist = hwVm.AutoRows.Concat(hwVm.FlaggedRows)
                .Where(r => r.ControlNo > 0
                            && r.Status != "Loaded"
                            && (r.NewHotWeight.HasValue
                                || !string.IsNullOrWhiteSpace(r.NewGrade)
                                || r.NewHealthScore.HasValue
                                || !string.IsNullOrWhiteSpace(r.NewAnimalControlNumber)
                                || r.IsCondemned))
                .ToList();

            if (rowsToPersist.Count == 0)
            {
                return Json(new
                {
                    success = false,
                    message = "No Hot Weight data to save. All HW rows are either already Loaded or have no data."
                });
            }

            // 3. Build KillAnimalData and persist via SaveKillDataAsync
            //    (writes HotWeight, Grade, HS, ACN, IsCondemned to bills — does NOT touch KillStatus)
            int billsUpdated = 0;
            var savedControlNos = new HashSet<int>();
            try
            {
                var killData = rowsToPersist.Select(r => new KillAnimalData
                {
                    ControlNo           = r.ControlNo,
                    AnimalControlNumber = !string.IsNullOrWhiteSpace(r.NewAnimalControlNumber)
                                            ? r.NewAnimalControlNumber
                                            : r.AnimalControlNumber,
                    HotWeight           = r.NewHotWeight,
                    Grade               = r.NewGrade,
                    HealthScore         = r.NewHealthScore,
                    IsCondemned         = r.IsCondemned,
                    Comment             = string.IsNullOrWhiteSpace(r.TrimComment) ? null : r.TrimComment
                }).ToList();

                billsUpdated = await _animalService.SaveKillDataAsync(killData);
                foreach (var d in killData) savedControlNos.Add(d.ControlNo);
                _logger.LogInformation(
                    "[SaveAllHw] Persisted HW data to {Count} bills across all pages. KillStatus unchanged.",
                    billsUpdated);
            }
            catch (Exception saveEx)
            {
                _logger.LogError(saveEx, "[SaveAllHw] Failed to persist HW data");
                return Json(new
                {
                    success = false,
                    message = "Save failed: " + saveEx.Message
                });
            }

            // 4. Mark saved rows as Loaded in staging so they don't reappear as Ready
            int markedLoaded = 0;
            foreach (var r in hwVm.AutoRows.Concat(hwVm.FlaggedRows))
            {
                if (savedControlNos.Contains(r.ControlNo) && r.Status != "Loaded")
                {
                    r.Status = "Loaded";
                    markedLoaded++;
                }
            }

            // Move newly-Loaded rows out of FlaggedRows into AutoRows so the breakdown stays sane
            var loadedFromFlagged = hwVm.FlaggedRows.Where(r => r.Status == "Loaded").ToList();
            foreach (var r in loadedFromFlagged)
            {
                hwVm.FlaggedRows.Remove(r);
                if (!hwVm.AutoRows.Any(a => a.ControlNo > 0 && a.ControlNo == r.ControlNo))
                    hwVm.AutoRows.Add(r);
            }

            // 5. Persist updated staging back to session AND shared staging
            try
            {
                var updatedJson = System.Text.Json.JsonSerializer.Serialize(hwVm);
                HttpContext.Session.SetString("HWPreview", updatedJson);
                await StagingBridge.WriteSharedAsync(
                    _stagingService,
                    StagingBridge.Types.HotWeight,
                    StagingBridge.SharedHotWeightKey,
                    updatedJson,
                    sourceFileName: null);
            }
            catch (Exception updEx)
            {
                _logger.LogWarning(updEx, "[SaveAllHw] Could not persist updated HW preview after save.");
            }

            return Json(new
            {
                success = true,
                billsUpdated,
                markedLoaded,
                message = $"Saved Hot Weight data for {billsUpdated} bills (HotWeight, Grade, HS, ACN written). Bills stay Pending — click 'Mark all complete (all pages)' when ready to finalize."
            });
        }


        public async Task<IActionResult> MarkKilled(int? vendorId, string? vendorIds)
        {
            // Forward query string so vendor filters survive the redirect.
            var routeValues = new Microsoft.AspNetCore.Routing.RouteValueDictionary();
            if (!string.IsNullOrEmpty(vendorIds))
                routeValues["vendorIds"] = vendorIds;
            else if (vendorId.HasValue)
                routeValues["vendorIds"] = vendorId.Value.ToString();

            await Task.CompletedTask; // method must remain async to match controller convention
            return RedirectToAction(nameof(MarkKilledFast), routeValues);

            // ----- Legacy implementation kept below for reference; never reached -----
            #pragma warning disable CS0162 // Unreachable code
            var vendors = await _vendorService.GetAllActiveAsync();

            // Multi-vendor support
            List<int> multiIds = new();
            if (!string.IsNullOrEmpty(vendorIds))
                multiIds = vendorIds.Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(v => int.TryParse(v.Trim(), out int id) ? id : 0)
                    .Where(id => id > 0).ToList();
            bool useMulti = multiIds.Any();

            var pending = useMulti
                ? await _animalService.GetPendingByVendorsAsync(multiIds)
                : await _animalService.GetPendingAsync(vendorId);
            ViewBag.VendorIds  = vendorIds ?? "";
            ViewBag.VendorList2 = vendors.Select(v => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem(v.VendorName, v.VendorID.ToString(), multiIds.Contains(v.VendorID) || v.VendorID == vendorId));
            ViewBag.SelectedVendorNames = useMulti
                ? vendors.Where(v => multiIds.Contains(v.VendorID)).Select(v => v.VendorName).ToList()
                : vendorId.HasValue ? vendors.Where(v => v.VendorID == vendorId).Select(v => v.VendorName).ToList() : new List<string>();

            // Pre-fill from Hot Weight import if just loaded
            // Read from TempData first, fall back to Session if TempData was cleared
            var hwLoaded = (TempData["HWLoaded"] as string == "1")
                        || (HttpContext.Session.GetString("HWLoaded") == "1");
            var hwJson   = TempData.Peek("HWPreview") as string
                        ?? HttpContext.Session.GetString("HWPreview");
            // Clear the session flag so it only pre-fills once
            HttpContext.Session.Remove("HWLoaded");
            var hwLookup = new Dictionary<int, HotWeightPreviewRow>();

            _logger.LogInformation("[MK-GET] hwLoaded={Loaded}, TempDataHWLoaded={TD}, SessionHWLoaded={Sess}, hwJsonLength={Len}",
                hwLoaded,
                TempData.Peek("HWLoaded") as string ?? "null",
                HttpContext.Session.GetString("HWLoaded") ?? "null",
                hwJson?.Length ?? 0);

            if (hwLoaded && !string.IsNullOrEmpty(hwJson))
            {
                try
                {
                    var hwVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(hwJson);
                    if (hwVm != null)
                    {
                        // Include AutoRows (OK + Overwrite) AND any FlaggedRows that have NewHotWeight
                        // (FlaggedRows can have NewHotWeight if staff fixed them and selected for load)
                        foreach (var r in hwVm.AutoRows.Where(r => r.ControlNo > 0 && r.NewHotWeight.HasValue))
                            hwLookup[r.ControlNo] = r;
                        foreach (var r in hwVm.FlaggedRows.Where(r => r.ControlNo > 0 && r.NewHotWeight.HasValue))
                            hwLookup[r.ControlNo] = r;
                        // Also include tag-matched rows that need ACN written
                        foreach (var r in hwVm.AutoRows.Where(r => r.ControlNo > 0 && !string.IsNullOrEmpty(r.NewAnimalControlNumber)))
                            if (!hwLookup.ContainsKey(r.ControlNo)) hwLookup[r.ControlNo] = r;
                        TempData["HWLoadedCount"]  = hwLookup.Count.ToString();
                        TempData["HWFlaggedCount"] = hwVm.FlaggedRows.Count(r => !hwLookup.ContainsKey(r.ControlNo)).ToString();
                        _logger.LogInformation("[MK-GET] hwLookup built: {Count} entries. Sample: {Sample}",
                            hwLookup.Count,
                            string.Join(", ", hwLookup.Take(3).Select(kv => $"CtrlNo={kv.Key} HW={kv.Value.NewHotWeight}")));
                    }
                    else
                    {
                        _logger.LogWarning("[MK-GET] Deserialize returned null. hwJson preview: {Preview}", hwJson?[..Math.Min(200, hwJson.Length)]);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "[MK-GET] Exception deserializing HWPreview JSON");
                }
            }
            else
            {
                _logger.LogWarning("[MK-GET] Skipping HW pre-fill: hwLoaded={Loaded}, hwJsonNull={Null}", hwLoaded, hwJson == null);
            }

            var vm = new MarkKilledViewModel
            {
                KillDate   = DateTime.Today,
                VendorId   = vendorId,
                VendorList = vendors.Select(v =>
                    new SelectListItem(v.VendorName, v.VendorID.ToString())),
                Animals = pending.Select(a =>
                {
                    hwLookup.TryGetValue(a.ControlNo, out var hw);
                    return new PendingAnimalRow
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
                        AnimalControlNumber = (hw != null && !string.IsNullOrEmpty(hw.NewAnimalControlNumber))
                                    ? hw.NewAnimalControlNumber
                                    : a.AnimalControlNumber,
                        // Append LTrim/RTrim to Comment when saving trim-variance rows
                        Comment      = hw != null && !string.IsNullOrEmpty(hw.TrimComment)
                                        ? (string.IsNullOrEmpty(a.Comment)
                                            ? hw.TrimComment
                                            : a.Comment + " " + hw.TrimComment)
                                        : a.Comment,
                        State               = a.State,
                        BuyerName           = a.BuyerName,
                        VetName             = a.VetName,
                        OfficeUse2          = a.OfficeUse2,
                        ProgramCode         = a.ProgramCode,
                        Selected            = false,
                        // When HW Load brings a condemned row in, pre-check the Condemned
                        // box. Otherwise preserve whatever's already on the bill.
                        IsCondemned         = (hw != null && hw.IsCondemned) ? true : a.IsCondemned,
                        HotWeight    = hw != null ? hw.NewHotWeight  : a.HotWeight,
                        Grade        = hw != null ? hw.NewGrade       : a.Grade,
                        HealthScore  = hw != null ? hw.NewHealthScore : a.HealthScore,
                        HwImported   = hw != null,
                    };
                }).ToList()
            };

            return View(vm);
#pragma warning restore CS0162
        }

        //  MARK AS KILLED — FAST, PAGINATED  (Phase 2b)
        public async Task<IActionResult> MarkKilledFast(
            string? vendorIds,
            int page = 1,
            int pageSize = 500,
            string? q = null)
        {
            if (page < 1) page = 1;
            if (pageSize < 100) pageSize = 100;
            if (pageSize > 1000) pageSize = 1000;

            // Parse optional vendor filter
            List<int> vendorIdList = new();
            if (!string.IsNullOrEmpty(vendorIds))
            {
                vendorIdList = vendorIds
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(v => int.TryParse(v.Trim(), out var id) ? id : 0)
                    .Where(id => id > 0)
                    .ToList();
            }

            // Single paged, JOINed query — uses IX_Animal_KillStatus_VendorID
            var (pending, totalCount) = await _animalQueryService.GetPendingPagedAsync(
                vendorIdList.Count > 0 ? vendorIdList : null,
                page,
                pageSize,
                string.IsNullOrWhiteSpace(q) ? null : q);

            // Vendor list for the picker
            var vendors = await _vendorService.GetAllActiveAsync();
            ViewBag.VendorIds = vendorIds ?? "";
            ViewBag.VendorList2 = vendors.Select(v =>
                new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem(
                    v.VendorName,
                    v.VendorID.ToString(),
                    vendorIdList.Contains(v.VendorID)));
            ViewBag.SelectedVendorNames = vendorIdList.Count > 0
                ? vendors.Where(v => vendorIdList.Contains(v.VendorID))
                        .Select(v => v.VendorName).ToList()
                : new List<string>();

            
            string? tempHwLoaded   = TempData.Peek("HWLoaded")  as string;
            string? tempHwPreview  = TempData.Peek("HWPreview") as string;
            string? sessHwLoaded   = HttpContext.Session.GetString("HWLoaded");
            string? sessHwPreview  = HttpContext.Session.GetString("HWPreview");

            
            var hwLoaded = (tempHwLoaded == "1") || (sessHwLoaded == "1");
            var hwJson   = !string.IsNullOrEmpty(tempHwPreview) ? tempHwPreview : sessHwPreview;

            
            if (hwLoaded && !string.IsNullOrEmpty(hwJson))
            {
                if (sessHwLoaded != "1")
                    HttpContext.Session.SetString("HWLoaded", "1");
                if (string.IsNullOrEmpty(sessHwPreview) || sessHwPreview != hwJson)
                    HttpContext.Session.SetString("HWPreview", hwJson);
            }

            _logger.LogInformation(
                "[MKFast-GET] page={Page} hwLoaded={HwLoaded} hwJsonLen={Len} src={Src}",
                page,
                hwLoaded,
                hwJson?.Length ?? 0,
                tempHwPreview != null ? "TempData" : (sessHwPreview != null ? "Session" : "none"));

            var hwLookup = new Dictionary<int, HotWeightPreviewRow>();

            if (hwLoaded && !string.IsNullOrEmpty(hwJson))
            {
                try
                {
                    var hwVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(hwJson);
                    if (hwVm != null)
                    {
                        
                        foreach (var r in hwVm.AutoRows.Where(r => r.ControlNo > 0 && r.NewHotWeight.HasValue && r.Status != "Loaded"))
                            hwLookup[r.ControlNo] = r;
                        foreach (var r in hwVm.FlaggedRows.Where(r => r.ControlNo > 0 && r.NewHotWeight.HasValue && r.Status != "Loaded"))
                            hwLookup[r.ControlNo] = r;
                        foreach (var r in hwVm.AutoRows.Where(r => r.ControlNo > 0 && !string.IsNullOrEmpty(r.NewAnimalControlNumber) && r.Status != "Loaded"))
                            if (!hwLookup.ContainsKey(r.ControlNo)) hwLookup[r.ControlNo] = r;

                        
                        var hwLoadedCount       = hwLookup.Count;
                        var hwFlaggedCount      = hwVm.FlaggedRows.Count(r => r.Status != "Loaded" && !hwLookup.ContainsKey(r.ControlNo));
                        var hwAlreadyKilledCount = hwVm.AutoRows.Count(r => r.Status == "Loaded");

                        TempData["HWLoadedCount"]  = hwLoadedCount.ToString();
                        TempData["HWFlaggedCount"] = hwFlaggedCount.ToString();

                        ViewBag.HwLoadedCount        = hwLoadedCount;
                        ViewBag.HwFlaggedCount       = hwFlaggedCount;
                        ViewBag.HwAlreadyKilledCount = hwAlreadyKilledCount;
                        ViewBag.HwTotalInExcel       = hwVm.TotalInExcel;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "[MKFast-GET] Exception deserializing HWPreview JSON");
                }
            }

            
            var vm = new MarkKilledViewModel
            {
                KillDate = DateTime.Today,
                VendorId = vendorIdList.Count == 1 ? vendorIdList[0] : (int?)null,
                VendorList = vendors.Select(v =>
                    new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem(v.VendorName, v.VendorID.ToString())),
                Animals = pending.Select(a =>
                {
                    hwLookup.TryGetValue(a.ControlNo, out var hw);
                    return new PendingAnimalRow
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
                        AnimalControlNumber = (hw != null && !string.IsNullOrEmpty(hw.NewAnimalControlNumber))
                                                ? hw.NewAnimalControlNumber
                                                : a.AnimalControlNumber,
                        Comment             = hw != null && !string.IsNullOrEmpty(hw.TrimComment)
                                                ? (string.IsNullOrEmpty(a.Comment)
                                                    ? hw.TrimComment
                                                    : a.Comment + " " + hw.TrimComment)
                                                : a.Comment,
                        State               = a.State,
                        BuyerName           = a.BuyerName,
                        VetName             = a.VetName,
                        OfficeUse2          = a.OfficeUse2,
                        ProgramCode         = a.ProgramCode,
                        Selected            = false,
                        // Pre-check Condemned when HW row is condemned. Otherwise preserve bill state.
                        IsCondemned         = (hw != null && hw.IsCondemned) ? true : a.IsCondemned,
                        HotWeight           = hw != null ? hw.NewHotWeight : a.HotWeight,
                        Grade               = hw != null ? hw.NewGrade : a.Grade,
                        HealthScore         = hw != null ? hw.NewHealthScore : a.HealthScore,
                        HwImported          = hw != null,
                    };
                }).ToList()
            };

            // Paging info for the view
            ViewBag.CurrentPage = page;
            ViewBag.PageSize    = pageSize;
            ViewBag.TotalCount  = totalCount;
            ViewBag.TotalPages  = (int)Math.Ceiling((double)totalCount / pageSize);
            ViewBag.SearchTerm  = q ?? "";

            return View("MarkKilledFast", vm);
        }

        
        [HttpGet]
        public async Task<IActionResult> MarkKilledFastCompleteCandidates(
            string? vendorIds,
            string? q = null)
        {
            // Same filter parsing as MarkKilledFast
            List<int> vendorIdList = new();
            if (!string.IsNullOrEmpty(vendorIds))
            {
                vendorIdList = vendorIds
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(v => int.TryParse(v.Trim(), out var id) ? id : 0)
                    .Where(id => id > 0).ToList();
            }

            
            var (allPending, totalCount) = await _animalQueryService.GetPendingPagedAsync(
                vendorIdList.Count > 0 ? vendorIdList : null,
                page: 1,
                pageSize: 5000,
                searchTerm: string.IsNullOrWhiteSpace(q) ? null : q);

            // Build HW lookup from session (NOT TempData — that's per-request).
            var hwLookup = new Dictionary<int, HotWeightPreviewRow>();
            var hwJson = HttpContext.Session.GetString("HWPreview");
            if (!string.IsNullOrEmpty(hwJson))
            {
                try
                {
                    var hwVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(hwJson);
                    if (hwVm != null)
                    {
                        foreach (var r in hwVm.AutoRows.Where(r => r.ControlNo > 0 && r.NewHotWeight.HasValue))
                            hwLookup[r.ControlNo] = r;
                        foreach (var r in hwVm.FlaggedRows.Where(r => r.ControlNo > 0 && r.NewHotWeight.HasValue))
                            hwLookup[r.ControlNo] = r;
                        foreach (var r in hwVm.AutoRows.Where(r => r.ControlNo > 0 && !string.IsNullOrEmpty(r.NewAnimalControlNumber)))
                            if (!hwLookup.ContainsKey(r.ControlNo)) hwLookup[r.ControlNo] = r;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "[MKFast-Candidates] Could not deserialize HWPreview JSON");
                }
            }

            
            static bool IsCompleteForKill(string? acn, decimal? hotWeight, string? grade, int? hs, bool isCondemned)
            {
                var trimmed = (acn ?? "").Trim().TrimStart('0');
                if (string.IsNullOrEmpty(trimmed)) return false;
                if (isCondemned) return true;
                return hotWeight.HasValue && hotWeight.Value > 0
                    && !string.IsNullOrWhiteSpace(grade)
                    && hs.HasValue && hs.Value >= 1 && hs.Value <= 5;
            }

            var completeIds = new List<int>();
            int hwLoadedComplete = 0;
            int billOnlyComplete = 0;
            foreach (var a in allPending)
            {
                hwLookup.TryGetValue(a.ControlNo, out var hw);

                var acn = (hw != null && !string.IsNullOrEmpty(hw.NewAnimalControlNumber))
                            ? hw.NewAnimalControlNumber
                            : a.AnimalControlNumber;
                var hotWeight = hw != null ? hw.NewHotWeight : a.HotWeight;
                var grade     = hw != null ? hw.NewGrade : a.Grade;
                var hs        = hw != null ? hw.NewHealthScore : a.HealthScore;
                var condemned = (hw != null && hw.IsCondemned) ? true : a.IsCondemned;

                if (IsCompleteForKill(acn, hotWeight, grade, hs, condemned))
                {
                    completeIds.Add(a.ControlNo);
                    if (hw != null) hwLoadedComplete++; else billOnlyComplete++;
                }
            }

            return Json(new
            {
                success = true,
                totalScanned = totalCount,
                completeCount = completeIds.Count,
                hwLoadedComplete,
                billOnlyComplete,
                controlNos = completeIds
            });
        }




        //Adding Post method to handle SaveMarkedEdits
        [HttpPost]
        [ValidateAntiForgeryToken]
        [RequestFormLimits(ValueCountLimit = 32768)]
        public async Task<IActionResult> SaveMarkKilledEdits(IFormCollection form, int? vendorId)
{
    var allIds = form["allIds"]
        .Where(v => !string.IsNullOrWhiteSpace(v))
        .Select(v => int.TryParse(v, out var id) ? id : 0)
        .Where(id => id > 0)
        .Distinct()
        .ToList();

    bool IsEditedForSave(int id)
    {
        var animalCtrl = NullIfEmpty(form[$"animalCtrlNo_{id}"].FirstOrDefault());
        var origAnimalCtrl = NullIfEmpty(form[$"origAnimalCtrlNo_{id}"].FirstOrDefault());
        bool animalCtrlChanged = !string.Equals(animalCtrl ?? "", origAnimalCtrl ?? "", StringComparison.Ordinal);

        bool liveWeightEntered = decimal.TryParse(form[$"liveWeight_{id}"], out var lw) && lw > 0;
        bool hotWeightEntered = decimal.TryParse(form[$"hotWeight_{id}"], out var hw) && hw > 0;
        bool gradeEntered = !string.IsNullOrWhiteSpace(form[$"grade_{id}"].FirstOrDefault());
        bool hsEntered = int.TryParse(form[$"healthScore_{id}"], out var hs) && hs > 0;

        var stateNow = NullIfEmpty(form[$"state_{id}"].FirstOrDefault());
        var stateOrig = NullIfEmpty(form[$"origState_{id}"].FirstOrDefault());
        bool stateChanged = !string.Equals(stateNow ?? "", stateOrig ?? "", StringComparison.Ordinal);

        var vetNow = NullIfEmpty(form[$"vetName_{id}"].FirstOrDefault());
        var vetOrig = NullIfEmpty(form[$"origVetName_{id}"].FirstOrDefault());
        bool vetChanged = !string.Equals(vetNow ?? "", vetOrig ?? "", StringComparison.Ordinal);

        var officeUse2Now = NullIfEmpty(form[$"officeUse2_{id}"].FirstOrDefault());
        var officeUse2Orig = NullIfEmpty(form[$"origOfficeUse2_{id}"].FirstOrDefault());
        bool office2Changed = !string.Equals(officeUse2Now ?? "", officeUse2Orig ?? "", StringComparison.Ordinal);

        var commentNow = NullIfEmpty(form[$"comment_{id}"].FirstOrDefault());
        var commentOrig = NullIfEmpty(form[$"origComment_{id}"].FirstOrDefault());
        bool commentChanged = !string.Equals(commentNow ?? "", commentOrig ?? "", StringComparison.Ordinal);

        bool condemnedNow = form[$"condemned_{id}"].Any(v => v == "true" || v == "on");
        bool condemnedOrig = form[$"origCondemned_{id}"].Any(v => v == "true" || v == "on");
        bool condemnedChanged = condemnedNow != condemnedOrig;

        decimal.TryParse(form[$"origLiveWeight_{id}"].FirstOrDefault(), out var origLw);
        bool liveWeightChanged = liveWeightEntered && lw != origLw;

        // Kill-date-only should not be treated as save-edit trigger.
        // Kill-date-only should not be treated as save-edit trigger.
        // If this row was pre-filled from HW import, treat any HW/Grade/HS value as an edit
        bool hwImported = form[$"hwImported_{id}"].FirstOrDefault() == "1";
        var origHw    = form[$"origHotWeight_{id}"].FirstOrDefault() ?? "";
        var origGrade = form[$"origGrade_{id}"].FirstOrDefault() ?? "";
        var origHs    = form[$"origHealthScore_{id}"].FirstOrDefault() ?? "";
        bool hwChanged    = hotWeightEntered && form[$"hotWeight_{id}"].FirstOrDefault()?.Replace(".0","") != origHw.Replace(".0","");
        bool gradeChanged = gradeEntered && !string.Equals(form[$"grade_{id}"].FirstOrDefault() ?? "", origGrade, StringComparison.OrdinalIgnoreCase);
        bool hsChanged    = hsEntered && form[$"healthScore_{id}"].FirstOrDefault() != origHs;
        bool hwImportedWithData = hwImported && (hotWeightEntered || gradeEntered || hsEntered);

        return animalCtrlChanged || liveWeightChanged || hwChanged || gradeChanged || hsChanged
            || hwImportedWithData || condemnedChanged || stateChanged || vetChanged || office2Changed || commentChanged;
    }

    var editedIds = allIds.Where(IsEditedForSave).ToList();

    if (!editedIds.Any())
    {
        TempData["ErrorMessage"] = "No editable field changes found to save.";
        return RedirectToAction(nameof(MarkKilled), new { vendorId });
    }

    //Consignment validation: If hot weight entered, live weight is required and must be >= Hot weight
        var validationErrors = editedIds
            .Select(id => ValidateLegacyMarkKilledRow(id, form))
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        if (validationErrors.Any())
        {
            TempData["ErrorMessage"] = string.Join(" ", validationErrors.Take(3)) +
                (validationErrors.Count > 3 ? " More rows have the same issue." : "");
            return RedirectToAction(nameof(MarkKilled), new { vendorId });
        }

    var animalData = editedIds.Select(id =>
    {
        DateTime? rowKillDate = null;
        var rowKillRaw = form[$"killDate_{id}"].FirstOrDefault();
        if (DateTime.TryParse(rowKillRaw, out var parsedRowDate))
            rowKillDate = parsedRowDate;
            var purchaseType = form[$"purchaseType_{id}"].FirstOrDefault() ?? "";
            var isConsignment = purchaseType.Contains("consignment", StringComparison.OrdinalIgnoreCase);
            var hwImported = form[$"hwImported_{id}"].FirstOrDefault() == "1";

            return new KillAnimalData
            {
                ControlNo = id,
                AnimalControlNumber = NormalizeAcn(form[$"animalCtrlNo_{id}"].FirstOrDefault()),
                KillDate = rowKillDate,
                LiveWeight = decimal.TryParse(form[$"liveWeight_{id}"], out var lw) && lw > 0
                    ? lw
                    : (hwImported
                        && isConsignment
                        && decimal.TryParse(form[$"hotWeight_{id}"], out var hwForLive) && hwForLive > 0
                            ? hwForLive
                            : (decimal?)null),
                HotWeight = decimal.TryParse(form[$"hotWeight_{id}"], out var hw) && hw > 0 ? hw
            : decimal.TryParse(form[$"origHotWeight_{id}"].FirstOrDefault(), out var origHwFallback) && origHwFallback > 0 ? origHwFallback
            : null,
            Grade = NullIfEmpty(form[$"grade_{id}"].FirstOrDefault())
                ?? NullIfEmpty(form[$"origGrade_{id}"].FirstOrDefault()),
            HealthScore = int.TryParse(form[$"healthScore_{id}"], out var hs) && hs > 0 ? hs
                        : int.TryParse(form[$"origHealthScore_{id}"].FirstOrDefault(), out var origHsFallback) && origHsFallback > 0 ? origHsFallback
            : null,
            IsCondemned = form[$"condemned_{id}"].Any(v => v == "true" || v == "on"),
            State = NullIfEmpty(form[$"state_{id}"].FirstOrDefault()),
            VetName = NullIfEmpty(form[$"vetName_{id}"].FirstOrDefault()),
            OfficeUse2 = NullIfEmpty(form[$"officeUse2_{id}"].FirstOrDefault()),
            Comment = NullIfEmpty(form[$"comment_{id}"].FirstOrDefault()),
        };
    }).ToList();

    int count = await _animalService.SaveKillDataAsync(animalData);

    TempData["SuccessMessage"] = $"{count} animal records updated. They remain pending until marked as killed.";
    TempData.Keep("HWPreview");
    return RedirectToAction(nameof(MarkKilled), new { vendorId });
}
        [HttpPost]
        [ValidateAntiForgeryToken]
        [RequestFormLimits(ValueCountLimit = 32768)]
        public async Task<IActionResult> MarkKilled(IFormCollection form)
    {
        if (!DateTime.TryParse(form["killDate"], out var defaultKillDate))
            defaultKillDate = DateTime.Today;

        var selectedIds = form["selectedIds"]
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Select(v => int.TryParse(v, out var id) ? id : 0)
            .Where(id => id > 0)
            .Distinct()
            .ToList();

            if (!selectedIds.Any())
            {
                TempData["ErrorMessage"] = "Select at least one row to mark as killed.";
                return RedirectToAction(nameof(MarkKilled));
            }

            var validationErrors = selectedIds
                .Select(id => ValidateLegacyMarkKilledRow(id, form))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            if (validationErrors.Any())
            {
                TempData["ErrorMessage"] = string.Join(" ", validationErrors.Take(3)) +
                    (validationErrors.Count > 3 ? " More rows have the same issue." : "");
                return RedirectToAction(nameof(MarkKilled));
            }

            var animalData = selectedIds.Select(id =>
            {
                DateTime? rowKillDate = null;
                var rowKillRaw = form[$"killDate_{id}"].FirstOrDefault();
                if (DateTime.TryParse(rowKillRaw, out var parsedRowDate))
                    rowKillDate = parsedRowDate;

                var purchaseType = form[$"purchaseType_{id}"].FirstOrDefault() ?? "";
                var isConsignment = purchaseType.Contains("consignment", StringComparison.OrdinalIgnoreCase);
                var hwImported = form[$"hwImported_{id}"].FirstOrDefault() == "1";

                return new KillAnimalData
                {
                    ControlNo = id,
                    AnimalControlNumber = NormalizeAcn(form[$"animalCtrlNo_{id}"].FirstOrDefault()),
                    KillDate = rowKillDate ?? defaultKillDate,
                    LiveWeight = decimal.TryParse(form[$"liveWeight_{id}"], out var lw) && lw > 0
                        ? lw
                        : (hwImported
                            && isConsignment
                            && decimal.TryParse(form[$"hotWeight_{id}"], out var hwForLive) && hwForLive > 0
                                ? hwForLive
                                : (decimal?)null),
                    HotWeight = decimal.TryParse(form[$"hotWeight_{id}"], out var hw) && hw > 0 ? hw : null,
                    Grade = NullIfEmpty(form[$"grade_{id}"].FirstOrDefault()),
                    HealthScore = int.TryParse(form[$"healthScore_{id}"], out var hs) && hs > 0 ? hs : null,
                    IsCondemned = form[$"condemned_{id}"].Any(v => v == "true" || v == "on"),
                };
            }).ToList();

            int count = await _animalService.MarkKilledWithDataAsync(animalData, defaultKillDate);

            TempData["SuccessMessage"] = 
                $"{count} animals marked as killed on {defaultKillDate:MM/dd/yyyy}.";

            return RedirectToAction("Tally", "Report", new { killDate = defaultKillDate.ToString("yyyy-MM-dd") });

        
    }


        // HOT WEIGHT IMPORT — GET
        // ---------------------------------------------------------------------
        // New flow (auto-pull from Hot Scale DB on every visit):
        //   1. Read existing SHARED staging (everyone's collaborative session).
        //   2. Try to pull today's rows from the Hot Scale source DB.
        //      - On success: parse + match the rows through the existing pipeline,
        //        then SMART-MERGE the fresh VM with any manual fixes already in
        //        existing staging (preserves user work across auto-refreshes).
        //      - On failure: keep existing staging intact, show inline warning.
        //   3. Render the preview.
        //
        // The Excel upload form remains available on the preview page as a
        // fallback for cases when the Hot Scale DB is unreachable.
        // ---------------------------------------------------------------------
        public async Task<IActionResult> HotWeightImport(int? tab = null)
        {
            // Manual-refresh design: this GET just rehydrates whatever is in
            // shared staging and renders. It does NOT pull from Hot Scale on
            // every visit — that turned a routine page-click into a 50+ second
            // wait while the SQL query, matching pipeline and staging write all
            // ran. Operators trigger fresh pulls explicitly via the
            // "Refresh from Hot Scale" button on the preview page.
            HotWeightImportViewModel? existingVm = null;
            DateTime? lastRefreshedUtc = null;
            try
            {
                var existingJson = await StagingBridge.ReadSharedAsync(
                    _stagingService,
                    StagingBridge.Types.HotWeight,
                    StagingBridge.SharedHotWeightKey);

                if (!string.IsNullOrEmpty(existingJson))
                {
                    try { existingVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(existingJson); }
                    catch (Exception dex) { _logger.LogWarning(dex, "[HW] Existing shared staging payload could not be deserialized; treating as empty."); }
                }

                // Read the staging batch metadata so we can show "last refreshed" to the operator.
                lastRefreshedUtc = await GetSharedHotWeightLastRefreshedAsync();
            }
            catch (Exception rex)
            {
                _logger.LogWarning(rex, "[HW] Could not read shared staging; rendering empty.");
            }

            var finalVm = existingVm ?? new HotWeightImportViewModel();
            if (existingVm != null) ViewBag.RestoredFromStaging = true;
            if (lastRefreshedUtc.HasValue)
            {
                ViewBag.LastRefreshedUtc = lastRefreshedUtc.Value;
                ViewBag.LastRefreshedAgeMinutes = (DateTime.UtcNow - lastRefreshedUtc.Value).TotalMinutes;
            }

            _logger.LogInformation(
                "[HW-GET] Rendering HotWeightPreview. TotalInExcel={Total}, AutoRows={Auto} (OK={Ok}, Loaded={Loaded}), FlaggedRows={Flagged}, DupRows={Dups}. LastRefreshed={Last}.",
                finalVm.TotalInExcel,
                finalVm.AutoRows.Count,
                finalVm.AutoRows.Count(r => r.Status == "OK"),
                finalVm.AutoRows.Count(r => r.Status == "Loaded"),
                finalVm.FlaggedRows.Count,
                finalVm.DupRows.Count,
                lastRefreshedUtc?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") ?? "never");

            return View("HotWeightPreview", finalVm);
        }

        // Looks up the most recent shared HotWeight batch's UpdatedAt (UTC).
        // Returns null if no shared batch exists yet.
        private async Task<DateTime?> GetSharedHotWeightLastRefreshedAsync()
        {
            try
            {
                var batch = await _stagingService.GetActiveBatchAsync(
                    StagingBridge.Types.HotWeight,
                    StagingBridge.SharedHotWeightKey);
                if (batch == null) return null;
                // Prefer LoadedAt if set; else CreatedAt.
                return batch.LoadedAt ?? batch.CreatedAt;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[HW] Could not read shared HotWeight batch metadata.");
                return null;
            }
        }


        // HOT WEIGHT IMPORT — PREVIEW POST (Excel upload fallback)
        // -----------------------------------------------------------------
        // Behaviour:
        //   1. Validate file.
        //   2. Parse and match using the shared helper (same code path as
        //      the auto-pull from Hot Scale DB).
        //   3. Smart-merge with any existing shared staging so manual
        //      fixes already in progress are preserved.
        //   4. Persist merged VM to shared staging.
        //
        // The Excel upload remains valid as a backup path for cases where
        // the Hot Scale DB is unreachable.
        // -----------------------------------------------------------------
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> HotWeightImport(IFormFile? file)
        {
            var vm = new HotWeightImportViewModel
            {
                FileName = file?.FileName ?? ""
            };

            if (file == null || file.Length == 0)
            {
                vm.Errors.Add("Please select an Excel file.");
                return View("HotWeightPreview", vm);
            }

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".xlsx")
            {
                vm.Errors.Add("Only .xlsx files are supported.");
                return View("HotWeightPreview", vm);
            }

            // Parse and match through the shared helper.
            HotWeightImportViewModel freshVm;
            try
            {
                freshVm = await ParseAndMatchHotWeightFileAsync(file, treatAsAutoPull: false);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[HW-IMPORT] Parse/match failed for upload {File}", file.FileName);
                vm.Errors.Add($"Parse error: {ex.Message}");
                return View("HotWeightPreview", vm);
            }

            // Smart-merge with existing shared staging (preserves manual fixes
            // from any in-progress collaborative review).
            HotWeightImportViewModel? existingVm = null;
            try
            {
                var existingJson = await StagingBridge.ReadSharedAsync(
                    _stagingService,
                    StagingBridge.Types.HotWeight,
                    StagingBridge.SharedHotWeightKey);
                if (!string.IsNullOrEmpty(existingJson))
                    existingVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(existingJson);
            }
            catch (Exception rex)
            {
                _logger.LogWarning(rex, "[HW-IMPORT] Could not read existing shared staging; treating as empty.");
            }

            // Pre-compute currently-killed bills so the merge can drop stale Loaded markers.
            HashSet<int>? killedControlNos = null;
            try
            {
                var allKilled = await _animalService.GetAllAsync();
                killedControlNos = allKilled
                    .Where(a => string.Equals(a.KillStatus, "Killed", StringComparison.OrdinalIgnoreCase))
                    .Select(a => a.ControlNo)
                    .ToHashSet();
            }
            catch (Exception kex)
            {
                _logger.LogWarning(kex, "[HW-IMPORT] Could not load currently-killed bills; merge will preserve all Loaded markers conservatively.");
            }

            var merged = MergeHotWeightStaging(freshVm, existingVm, killedControlNos);

            // Persist to shared staging
            try
            {
                var mergedJson = System.Text.Json.JsonSerializer.Serialize(merged);
                await StagingBridge.WriteSharedAsync(
                    _stagingService,
                    StagingBridge.Types.HotWeight,
                    StagingBridge.SharedHotWeightKey,
                    mergedJson,
                    sourceFileName: file.FileName);
            }
            catch (Exception wex)
            {
                _logger.LogWarning(wex, "[HW-IMPORT] Could not write merged VM to shared staging; preview will still render.");
            }

            return View("HotWeightPreview", merged);
        }

        // ---------------------------------------------------------------------
        // ParseAndMatchHotWeightFileAsync
        //
        // Shared helper used by both the Excel-upload POST and the auto-pull
        // GET (via a synthesized in-memory workbook). Runs the existing
        // 8-step matching pipeline against the workbook and returns a
        // populated HotWeightImportViewModel (no staging side-effects).
        // ---------------------------------------------------------------------
        private async Task<HotWeightImportViewModel> ParseAndMatchHotWeightFileAsync(
            IFormFile file,
            bool treatAsAutoPull)
        {
            var vm = new HotWeightImportViewModel
            {
                FileName = file?.FileName ?? ""
            };

            // Use shared rules so this matches MarkKilledApi exactly.
            // Includes NP on both lists per operations spec.
            var bullGrades = GradeRules.BullGrades;
            var cowGrades  = GradeRules.CowGrades;


            string? NormalizeSexCode(string? sex)
            {
                var s = (sex ?? "").Trim().ToUpperInvariant();
                return s switch
                {
                    "B" => "B",
                    "F" => "F",
                    _ => null
                };
            }

            string? NormalizeTypeGroup(string? type)
            {
                var t = (type ?? "").Trim().ToUpperInvariant();
                return t switch
                {
                    "BULL" => "BULL",
                    "DCOW" => "COW",
                    "BCOW" => "COW",
                    _ => null
                };
            }

            (HashSet<string>? AllowedGrades, bool HasMismatch, string SexCode, string TypeCode) ResolveGradeRules(string? type, string? sex)
            {
                var sexCode = NormalizeSexCode(sex);
                var typeGroup = NormalizeTypeGroup(type);

                bool mismatch = sexCode != null
                                && typeGroup != null
                                && ((sexCode == "B" && typeGroup != "BULL")
                                    || (sexCode == "F" && typeGroup != "COW"));

                if (sexCode == "B") return (bullGrades, mismatch, "B", typeGroup ?? "-");
                if (sexCode == "F") return (cowGrades, mismatch, "F", typeGroup ?? "-");

                if (typeGroup == "BULL") return (bullGrades, mismatch, "-", "BULL");
                if (typeGroup == "COW") return (cowGrades, mismatch, "-", "COW");

                return (null, mismatch, sexCode ?? "-", typeGroup ?? "-");
            }

            try
            {
        // existing code...
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using var wb = new XLWorkbook(stream);
                var ws = wb.Worksheets.First();

                // Build column map
                var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 30;
                for (int c = 1; c <= lastCol; c++)
                {
                    var h = ws.Cell(1, c).GetString().Trim().ToLowerInvariant();
                    if (!string.IsNullOrEmpty(h)) colMap[h] = c;
                }

                int Col(params string[] names)
                {
                    foreach (var n in names)
                        if (colMap.TryGetValue(n.ToLowerInvariant(), out int c)) return c;
                    return -1;
                }

                // Detect required columns — block import if wrong file
                // Aliases cover both the actual Hot Scale export headers ("Number", "Side 1 Hot", "Side 2 Hot", "Grade\n", "HealthScore")
                // and any future renamed variants
                int colACN     = Col("number", "animal control number", "acn", "animal control no", "controlno", "ctrl no", "control no");
                int colSide1   = Col("side 1 hot", "hot weight side 1", "hotweightside1", "side 1", "side1", "hw side 1", "hwside1", "s1", "side1hot");
                int colSide2   = Col("side 2 hot", "hot weight side 2", "hotweightside2", "side 2", "side2", "hw side 2", "hwside2", "s2", "side2hot");
                int colGrade2  = Col("grade 2", "grade2", "grade\n2");
                int colOrigin  = Col("origin");
                int colLot     = Col("lot");
                int colSex     = Col("sex");
                int colType    = Col("type");
                int colProgram = Col("program");
                int colGrade   = Col("grade");
                int colHS      = Col("healthscore", "health score", "hs", "h s", "health_score");
                // Extended tag columns for ACN auto-match
                int colBackTag    = Col("backtag", "back tag", "back_tag", "btag");
                int colLiveWeight = Col("liveweight", "live weight", "live wt", "liveWt");
                int colTag1    = Col("tag1", "tag 1", "tag number one", "tag one");
                int colTag2    = Col("tag2", "tag 2", "tag number two", "tag two");
                bool hasTagCols = colBackTag > 0 || colTag1 > 0 || colTag2 > 0;

                if (colACN < 0 && !hasTagCols && (colSide1 < 0 && colSide2 < 0))
                {
                    vm.Errors.Add("Wrong file: could not find required columns. Need Side1+Side2 columns AND either ACN ('Number') or tag columns ('BackTag','Tag1','Tag2'). Please upload the Hot Scale Parsed Data report.");
                    return vm;
                }

                

                // Parse Excel rows
                var excelRows = new List<(string acn, string backTag, string tag1, string tag2, decimal? s1, decimal? s2, string? grade, int? hs, decimal? liveWt, string? grade2, string? origin, string? lot, string? sex, string? type, string? program)>();
    
                var seenAcns  = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int lastRow   = ws.LastRowUsed()?.RowNumber() ?? 1;

                for (int row = 2; row <= lastRow; row++)
                {
                    var acnRaw    = colACN     > 0 ? GetCellString(ws.Cell(row, colACN))     : "";
                    var backTagRaw = colBackTag > 0 ? GetCellString(ws.Cell(row, colBackTag)) : "";
                    var tag1Raw    = colTag1    > 0 ? GetCellString(ws.Cell(row, colTag1))    : "";
                    var tag2Raw    = colTag2    > 0 ? GetCellString(ws.Cell(row, colTag2))    : "";

                    // Skip blank rows — need at least ACN or one tag
                    bool hasIdentifier = !string.IsNullOrWhiteSpace(acnRaw)
                                      || !string.IsNullOrWhiteSpace(backTagRaw)
                                      || !string.IsNullOrWhiteSpace(tag1Raw)
                                      || !string.IsNullOrWhiteSpace(tag2Raw);
                    if (!hasIdentifier) continue;

                    vm.TotalInExcel++;
                    var acn      = acnRaw.Trim().TrimStart('0');
                    var backTag  = backTagRaw.Trim();
                    var tag1     = tag1Raw.Trim();
                    var tag2     = tag2Raw.Trim();

                    decimal? s1 = null, s2 = null;
                    if (colSide1 > 0) { var v = GetDecimalCell(ws.Cell(row, colSide1)); if (v > 0) s1 = v; }
                    if (colSide2 > 0) { var v = GetDecimalCell(ws.Cell(row, colSide2)); if (v > 0) s2 = v; }

                    decimal? liveWt = null;
                    if (colLiveWeight > 0) { var lv = GetDecimalCell(ws.Cell(row, colLiveWeight)); if (lv > 0) liveWt = lv; }
                    var grade2Raw = colGrade2 > 0 ? GetCellString(ws.Cell(row, colGrade2)).Trim().TrimEnd() : null;
                    string? grade2 = string.IsNullOrWhiteSpace(grade2Raw) ? null : grade2Raw.ToUpper().Trim();
                    var originRaw = colOrigin > 0 ? GetCellString(ws.Cell(row, colOrigin)).Trim() : null;
                    string? fileOrigin = string.IsNullOrWhiteSpace(originRaw) ? null : originRaw.Trim();
                    var fileLot  = colLot  > 0 ? GetCellString(ws.Cell(row, colLot)).Trim()  : null;
                    var fileSex  = colSex  > 0 ? GetCellString(ws.Cell(row, colSex)).Trim()  : null;
                    var fileType = colType > 0 ? GetCellString(ws.Cell(row, colType)).Trim()  : null;
                    var fileProgram = colProgram > 0 ? GetCellString(ws.Cell(row, colProgram)).Trim() : null;
                    if (string.IsNullOrWhiteSpace(fileLot))  fileLot  = null;
                    if (string.IsNullOrWhiteSpace(fileSex))  fileSex  = null;
                    if (string.IsNullOrWhiteSpace(fileType)) fileType = null;
                    if (string.IsNullOrWhiteSpace(fileProgram)) fileProgram = null;

                    var gradeRaw = colGrade > 0 ? GetCellString(ws.Cell(row, colGrade)) : null;
                    string? grade = string.IsNullOrWhiteSpace(gradeRaw) ? null : gradeRaw.Trim().ToUpper();

                    int? hs = null;
                    if (colHS > 0) { var hsStr = GetCellString(ws.Cell(row, colHS)); if (int.TryParse(hsStr, out int hv)) hs = hv; }
                    // "slow" in Tag1 means the animal was slow at weigh-in → HealthScore = 5
                    if (tag1Raw.Trim().Equals("slow", StringComparison.OrdinalIgnoreCase) && !hs.HasValue)
                        hs = 5;

                    // Dedup key: prefer ACN, fallback to BackTag then Tag1
                    var dedupKey = !string.IsNullOrEmpty(acn) ? acn
                                 : !string.IsNullOrEmpty(backTag) ? "bt:" + backTag
                                 : "t1:" + tag1;
                    if (string.IsNullOrEmpty(dedupKey)) dedupKey = "row:" + row;

                    if (seenAcns.ContainsKey(dedupKey))
                        vm.FlaggedRows.Add(new HotWeightPreviewRow
                        {
                            RowKey = Guid.NewGuid().ToString("N"),
                            AnimalControlNumber = acn, Side1 = s1, Side2 = s2,
                            NewGrade = grade, NewHealthScore = hs,
                            Status = "Flag", FlagReason = "Duplicate row in the Excel file — both held for manual review"
                        });
                    else
                    {
                        seenAcns[dedupKey] = row;
                        excelRows.Add((acn, backTag, tag1, tag2, s1, s2, grade, hs, liveWt, grade2, fileOrigin, fileLot, fileSex, fileType, fileProgram));
                    }
                }

                if (excelRows.Count == 0)
                {
                    vm.Errors.Add("No valid rows found in Excel file. Ensure you have at least one row with an Animal Control Number (ACN), BackTag, or Tag 1/2 value.");
                    _logger.LogWarning("[HW-IMPORT] No valid rows parsed from file");
                    return vm;
                }

                //  Match against system — 8-step pipeline 
                //  Match against system — 8-step pipeline 
                var allAcns    = excelRows.Select(r => r.acn).Where(a => !string.IsNullOrEmpty(a)).ToList();
                _logger.LogInformation("[HW-IMPORT] ACN lookup: {Count} distinct ACNs from {Total} rows", allAcns.Count, excelRows.Count);

                Dictionary<string, BarnData.Data.Entities.Animal> acnAnimals = new();
                try
                {
                    var acnResults = await _animalService.GetByAnimalControlNumbersAsync(allAcns);
                    // Safe against duplicate ACNs in DB — group first, then pick first per key
var acnGrouped = acnResults
    .GroupBy(a => (a.AnimalControlNumber ?? "").TrimStart('0'),
             StringComparer.OrdinalIgnoreCase)
    .ToList();

                acnAnimals = acnGrouped.ToDictionary(
                    g => g.Key,
                    g => g.First(),
                    StringComparer.OrdinalIgnoreCase);

                // Add DB-level duplicates to the Duplicates tab
                foreach (var grp in acnGrouped.Where(g => g.Count() > 1))
                {
                    foreach (var dupAnimal in grp)
                    {
                        vm.DupRows.Add(new HotWeightPreviewRow
                        {
                            RowKey              = $"DB-DUP-{dupAnimal.ControlNo}",
                            ControlNo           = dupAnimal.ControlNo,
                            AnimalControlNumber = dupAnimal.AnimalControlNumber ?? "",
                            Status              = "Dup",
                            FlagReason          = $"DB has multiple records for ACN {grp.Key}"
                        });
                    }
                }
                _logger.LogInformation("[HW-IMPORT] ACN lookup returned {Count} animals ({Dups} DB duplicates)",
                    acnAnimals.Count, vm.DupRows.Count);
                }
                catch (Exception acnEx)
                {
                    _logger.LogError(acnEx, "[HW-IMPORT] ACN lookup failed. AllAcns: {Acns}", 
                        string.Join(", ", allAcns.Take(10)));
                    vm.Errors.Add($"Database error during ACN lookup: {acnEx.Message}");
                    return vm;
                }
                static bool IsEidShaped(string t) =>
                    !string.IsNullOrEmpty(t) && t.Length >= 10
                    && t.All(char.IsDigit)
                    && (t.StartsWith("840") || t.StartsWith("124"));
                static IEnumerable<string> EidSuffixLadder(string eid)
                {
                    // Longest → shortest. Minimum length 3.
                    for (int n = 7; n >= 3; n--)
                        if (eid.Length >= n) yield return eid[^n..];
                }
                var allExactTags = excelRows
                .SelectMany(r => new[] { r.backTag, r.tag1, r.tag2 })
                .Where(t => !string.IsNullOrEmpty(t))
                .Concat(excelRows
                    .SelectMany(r => new[] { r.backTag, r.tag1, r.tag2 })
                    .Where(t => !string.IsNullOrEmpty(t) && IsEidShaped(t))
                    .SelectMany(EidSuffixLadder))
                .Concat(excelRows
                    .Select(r => r.backTag)
                    .Where(t => !string.IsNullOrEmpty(t))
                    .Select(t => System.Text.RegularExpressions.Regex.Match(t, @"^(\d+)([A-Z]+)(\d{4})$"))
                    .Where(m => m.Success)
                    .Select(m => m.Groups[3].Value))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

                _logger.LogInformation("[HW-IMPORT] Tag lookup: {Count} distinct tags", allExactTags.Count);

                List<BarnData.Data.Entities.Animal> tagAnimals = new();
                try
                {
                    tagAnimals = allExactTags.Any()
                        ? (await _animalService.GetByTagsAsync(allExactTags)).ToList()
                        : new();
                    _logger.LogInformation("[HW-IMPORT] Tag lookup returned {Count} animals", tagAnimals.Count);
                }
                catch (Exception tagEx)
                {
                    _logger.LogError(tagEx, "[HW-IMPORT] Tag lookup failed. SampleTags: {Tags}", 
                        string.Join(", ", allExactTags.Take(10)));
                    vm.Errors.Add($"Database error during tag lookup: {tagEx.Message}");
                    return vm;
                }
                var tagIndex = new Dictionary<string, List<BarnData.Data.Entities.Animal>>(StringComparer.OrdinalIgnoreCase);
                foreach (var a in tagAnimals)
                    foreach (var t in new[] { a.TagNumber1, a.TagNumber2, a.Tag3 }.Where(t => !string.IsNullOrEmpty(t)))
                    { if (!tagIndex.ContainsKey(t!)) tagIndex[t!] = new(); tagIndex[t!].Add(a); }

                // Pending animals for weight fallback (loaded once)
                List<BarnData.Data.Entities.Animal>? pendingAnimals = null;

                var matchedControlNos = new HashSet<int>();

                // Helper: try exact tag lookup with leading-zero variants
                // Tries: exact → stripped (0893→893) → padded (893→0893)
                List<BarnData.Data.Entities.Animal> ExactTagLookup(string tag)
                {
                    if (tagIndex.TryGetValue(tag, out var m)) return m;
                    var stripped = tag.TrimStart('0');
                    if (stripped != tag && tagIndex.TryGetValue(stripped, out var m2)) return m2;
                    var padded = "0" + tag;   // FIX 2: reverse — try adding leading zero
                    if (tagIndex.TryGetValue(padded, out var m3)) return m3;
                    return new();
                }
                List<BarnData.Data.Entities.Animal> PrepareTagCandidates(
                List<BarnData.Data.Entities.Animal> hits,
                string context,
                ref string reason)
                {
                    var pendingHits = hits
                        .Where(a => string.Equals(a.KillStatus ?? "", "Pending", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (!pendingHits.Any() && hits.Any())
                    {
                        reason = context + " only matches non-pending records";
                        return new List<BarnData.Data.Entities.Animal>();
                    }

                    var missingAcnPendingHits = pendingHits
                        .Where(a => IsAcnMissing(a.AnimalControlNumber))
                        .ToList();

                    if (!missingAcnPendingHits.Any() && pendingHits.Any())
                    {
                        reason = context + " only matches pending records that already have ACN assigned";
                    }

                    return missingAcnPendingHits;
                }

                (BarnData.Data.Entities.Animal? match,
                 List<BarnData.Data.Entities.Animal>? candidates,
                 string usedSuffix,
                 string reason)
                FindByEidSuffix(string eid, decimal? fileLiveWt, decimal? fileSide1, decimal? fileSide2,
                                string sourceContext /* "BackTag" or "Tag1" etc. */)
                {
                    if (string.IsNullOrEmpty(eid)) return (null, null, "", "");

                    var fileHotTotal = (fileSide1 ?? 0) + (fileSide2 ?? 0);

                    for (int n = 7; n >= 3; n--)
                    {
                        if (eid.Length < n) continue;
                        var suffix = eid[^n..];
                        var hitsAll = ExactTagLookup(suffix);
                        if (hitsAll.Count == 0) continue;

                        string stageReason = "";
                        var unassigned = PrepareTagCandidates(hitsAll,
                            $"{sourceContext} EID '{eid}' suffix '{suffix}'", ref stageReason);

                        // All hits already have ACN assigned → shrink further.
                        if (unassigned.Count == 0) continue;

                        if (unassigned.Count == 1)
                        {
                            return (unassigned[0], null, suffix,
                                $"{sourceContext}-EID-Last{n}");
                        }

                        // 2+ candidates: closest-weight wins; tie → flag with picker.
                        decimal DiffFor(BarnData.Data.Entities.Animal a)
                        {
                            if (a.LiveWeight > 0 && fileLiveWt.HasValue && fileLiveWt.Value > 0)
                                return Math.Abs(a.LiveWeight - fileLiveWt.Value);
                            if (a.LiveWeight > 0 && fileHotTotal > 0)
                                return Math.Abs(a.LiveWeight - fileHotTotal);
                            return decimal.MaxValue; // cannot compare - won't win, won't uniquely lose either
                        }

                        var scored = unassigned
                            .Select(a => new { Animal = a, Diff = DiffFor(a) })
                            .OrderBy(x => x.Diff)
                            .ToList();

                        var bestDiff = scored[0].Diff;
                        var winners = scored.Where(x => x.Diff == bestDiff).ToList();

                        if (bestDiff == decimal.MaxValue)
                        {
                            // No weights at all to compare → staff picks.
                            return (null, unassigned, suffix,
                                $"{sourceContext} EID '{eid}' suffix '{suffix}' — {unassigned.Count} candidates (no weight data to tiebreak)");
                        }

                        if (winners.Count == 1)
                        {
                            return (winners[0].Animal, null, suffix,
                                $"{sourceContext}-EID-Last{n}+ClosestWeight(diff {bestDiff:N0} lbs)");
                        }

                        // Exact tie on distance → staff picks.
                        return (null, unassigned, suffix,
                            $"{sourceContext} EID '{eid}' suffix '{suffix}' — {unassigned.Count} candidates, weight tied at {bestDiff:N0} lbs");
                    }

                    return (null, null, "", ""); // no hits at any length
                }

                
                (BarnData.Data.Entities.Animal? animal, string reason) WeightPickFromCandidates(
                    List<BarnData.Data.Entities.Animal> candidates, decimal? fileLiveWt, string context,
                    string vendorCode = "", string filePurchaseType = "")
                {
                    if (candidates.Count == 0)
                        return (null, $"{context} — no candidates");

                    // Score each candidate using explicit variables (avoids positional tuple naming issues)
                    var scoredList = new List<(BarnData.Data.Entities.Animal Animal, int Score, string Signals)>();
                    foreach (var ca in candidates)
                    {
                        int sc = 0;
                        string sg = "";

                        if (!string.IsNullOrEmpty(vendorCode) && ca.Vendor != null)
                        {
                            var vName = (ca.Vendor.VendorName ?? "").ToUpper();
                            var words = vName.Split(new[]{' ','-','.'}, StringSplitOptions.RemoveEmptyEntries);
                            var initials  = string.Concat(words.Select(w => w.Length > 0 ? w[0].ToString() : ""));
                            var firstWord = words.Length > 0 ? words[0] : "";
                            bool inInitials  = initials.Contains(vendorCode);
                            bool inFirstWord = firstWord.StartsWith(vendorCode, StringComparison.OrdinalIgnoreCase);
                            bool inName      = vName.Contains(vendorCode);
                            if (inInitials || inFirstWord || inName)
                            { sc += 30; sg += $"VendorCode({vendorCode}→{ca.Vendor.VendorName}) "; }
                        }

                        if (fileLiveWt.HasValue && fileLiveWt.Value > 0 && ca.LiveWeight > 0)
                        {
                            var diff = Math.Abs(ca.LiveWeight - fileLiveWt.Value);
                            if      (diff <= 50)  { sc += 20; sg += $"LiveWt(±{diff:N0}) "; }
                            else if (diff <= 100) { sc += 15; sg += $"LiveWt(±{diff:N0}) "; }
                            else if (diff <= 150) { sc += 10; sg += $"LiveWt(±{diff:N0}) "; }
                            else if (diff <= 200) { sc += 5;  sg += $"LiveWt(±{diff:N0}) "; }
                        }

                        if (!string.IsNullOrEmpty(filePurchaseType) && !string.IsNullOrEmpty(ca.PurchaseType))
                        {
                            bool saleFile = filePurchaseType.Contains("Sale", StringComparison.OrdinalIgnoreCase);
                            bool saleDb   = ca.PurchaseType.Contains("Sale", StringComparison.OrdinalIgnoreCase);
                            if (saleFile == saleDb) { sc += 10; sg += "PurchType "; }
                        }

                        scoredList.Add((ca, sc, sg.Trim()));
                    }
                    scoredList = scoredList.OrderByDescending(x => x.Score).ToList();

                    if (scoredList.Count == 0)
                        return (null, $"{context} — no scoreable candidates");

                    var bestAnimal   = scoredList[0].Animal;
                    var bestScore    = scoredList[0].Score;
                    var bestSignals  = scoredList[0].Signals;
                    var secondScore  = scoredList.Count > 1 ? scoredList[1].Score : -1;

                    int gap = bestScore - secondScore;
                    if (bestScore >= 10 && gap >= 10)
                    {
                        return (bestAnimal,
                            $"best-match ({context}: score={bestScore} [{bestSignals}] vs next={secondScore}, flagged for confirmation)");
                    }

                    // Fallback: if only live weight available and clear gap (≥50 lbs)
                    if (fileLiveWt.HasValue && fileLiveWt.Value > 0)
                    {
                        var byWt = candidates
                            .Where(a => a.LiveWeight > 0)
                            .OrderBy(a => Math.Abs(a.LiveWeight - fileLiveWt.Value))
                            .ToList();
                        if (byWt.Count >= 1)
                        {
                            var bwBest    = byWt[0];
                            var bwBestD   = Math.Abs(bwBest.LiveWeight - fileLiveWt.Value);
                            var bwSecondD = byWt.Count > 1 ? Math.Abs(byWt[1].LiveWeight - fileLiveWt.Value) : (decimal)9999;
                            if (bwBestD <= 150 && (bwSecondD - bwBestD) >= 50)
                                return (bwBest, $"weight-only fallback ({context}, LiveWt {fileLiveWt:N0}→DB {bwBest.LiveWeight:N0}, diff {bwBestD:N0} lbs, flagged for confirmation)");
                        }
                    }

                    return (null, $"{context} — {candidates.Count} candidates, scores too close (best={bestScore} [{bestSignals}])");
                }

                // Helper: build HwCandidate list from DB animals for picker UI (all 24 fields)
                List<HwCandidate> BuildCandidates(List<BarnData.Data.Entities.Animal> animals, decimal? fw)
                {
                    return animals.Select(a => new HwCandidate
                    {
                        ControlNo           = a.ControlNo,
                        AnimalControlNumber = a.AnimalControlNumber ?? "",
                        Tag1                = a.TagNumber1 ?? "",
                        Tag2                = a.TagNumber2 ?? "",
                        Tag3                = a.Tag3 ?? "",
                        VendorName          = a.Vendor?.VendorName ?? "",
                        LiveWeight          = a.LiveWeight,
                        WeightDiff          = fw.HasValue && a.LiveWeight > 0
                                              ? Math.Abs(a.LiveWeight - fw.Value) : 0,
                        AnimalType          = a.AnimalType ?? "",
                        AnimalType2         = a.AnimalType2 ?? "",
                        ProgramCode         = a.ProgramCode ?? "",
                        PurchaseDate        = a.PurchaseDate.ToString("MM/dd/yyyy"),
                        PurchaseType        = a.PurchaseType ?? "",
                        LiveRate            = a.LiveRate,
                        ConsignmentRate     = a.ConsignmentRate,
                        Grade               = a.Grade ?? "",
                        Grade2              = a.Grade2 ?? "",
                        HealthScore         = a.HealthScore,
                        Comment             = a.Comment ?? "",
                        State               = a.State ?? "",
                        BuyerName           = a.BuyerName ?? "",
                        VetName             = a.VetName ?? "",
                        Origin              = a.Origin ?? "",
                        KillStatus          = a.KillStatus ?? "",
                        CreatedAt           = a.CreatedAt.ToString("MM/dd/yyyy HH:mm"),
                    }).OrderBy(c => c.WeightDiff).ToList();
                }

                foreach (var (acn, backTag, tag1, tag2, s1, s2, grade, hs, liveWt, grade2, fileOrigin, fileLot, fileSex, fileType, fileProgram) in excelRows)
                {
                    BarnData.Data.Entities.Animal? animal = null;
                    string matchMethod = "ACN";
                    string flagReason  = "";
                    List<HwCandidate>? _candidateBuffer = null;

                    if (animal == null && (!string.IsNullOrWhiteSpace(backTag) || !string.IsNullOrWhiteSpace(tag1) || !string.IsNullOrWhiteSpace(tag2)))
                    {
                        var completeByBackTagMatches = tagAnimals
                                                        .Where(a =>
                                                            !IsAcnMissing(a.AnimalControlNumber) &&
                                                            a.HotWeight.HasValue && a.HotWeight.Value > 0 &&
                                                            AnimalMatchesAnyFileTag(a, backTag, tag1, tag2))
                                                        .ToList();

                        if (completeByBackTagMatches.Count == 1)
                        {
                            var completeByBackTag = completeByBackTagMatches[0];
                            var dbAcn = (completeByBackTag.AnimalControlNumber ?? "").TrimStart('0');
                            var dbHw = completeByBackTag.HotWeight!.Value;

                            vm.DupRows.Add(new HotWeightPreviewRow
                            {
                                ControlNo = completeByBackTag.ControlNo,
                                AnimalControlNumber = dbAcn,
                                CurrentHotWeight = dbHw.ToString("N1"),
                                CurrentGrade = completeByBackTag.Grade,
                                CurrentHealthScore = completeByBackTag.HealthScore?.ToString(),
                                Side1 = s1,
                                Side2 = s2,
                                NewGrade = grade,
                                NewGrade2 = grade2,
                                NewHealthScore = hs,
                                FileLiveWeight = liveWt,
                                FileLot = fileLot,
                                FileSex = fileSex,
                                FileType = fileType,
                                FileOrigin = fileOrigin,
                                FileBackTag = backTag,
                                FileTag1 = tag1,
                                FileTag2 = tag2,
                                FileProgram = fileProgram,
                                Status = "Dup",
                                FlagReason = $"Duplicate by ACN/Tag match — ACN {dbAcn}, HotWeight {dbHw:N1} already in DB"
                            });

                            vm.AlreadyHasData++;
                            matchedControlNos.Add(completeByBackTag.ControlNo);
                            continue;
                        }
                        else if (completeByBackTagMatches.Count > 1)
                        {
                            flagReason = $"Tag identifiers (BackTag/Tag1/Tag2) match {completeByBackTagMatches.Count} complete animals (ACN + HotWeight already set) — manual review required";
                            _candidateBuffer = BuildCandidates(completeByBackTagMatches, liveWt);
                            goto AddFlag;
                        }
                    }
                    
                    //Direct ACN match 
                    if (!string.IsNullOrEmpty(acn))
                        acnAnimals.TryGetValue(acn, out animal);
                    if (animal != null
                        && animal.HotWeight.HasValue && animal.HotWeight.Value > 0
                        && AnimalMatchesAnyFileTag(animal, backTag, tag1, tag2))
                    {
                        var dbAcn = (animal.AnimalControlNumber ?? "").TrimStart('0');
                        var dbHw  = animal.HotWeight!.Value;
                        vm.DupRows.Add(new HotWeightPreviewRow
                        {
                            ControlNo            = animal.ControlNo,
                            AnimalControlNumber  = dbAcn,
                            CurrentHotWeight     = dbHw.ToString("N1"),
                            CurrentGrade         = animal.Grade,
                            CurrentHealthScore   = animal.HealthScore?.ToString(),
                            Side1                = s1,
                            Side2                = s2,
                            NewGrade             = grade,
                            NewGrade2            = grade2,
                            NewHealthScore       = hs,
                            FileLiveWeight       = liveWt,
                            FileLot              = fileLot,
                            FileSex              = fileSex,
                            FileType             = fileType,
                            FileOrigin           = fileOrigin,
                            FileBackTag          = backTag,
                            FileTag1             = tag1,
                            FileTag2             = tag2,
                            FileProgram          = fileProgram,
                            MatchMethod          = "ACN+Tag (already complete)",
                            Status               = "Dup",
                            FlagReason           = $"Already complete — DB has HotWeight {dbHw:N1} lbs, Grade {animal.Grade ?? "-"}, HS {animal.HealthScore?.ToString() ?? "-"}, ACN {dbAcn}",
                        });
                        vm.AlreadyHasData++;
                        matchedControlNos.Add(animal.ControlNo);
                        continue;
                    }

                    if (animal != null)
                    {
                        bool fileHasAnyTag = !string.IsNullOrWhiteSpace(backTag)
                                             || !string.IsNullOrWhiteSpace(tag1)
                                             || !string.IsNullOrWhiteSpace(tag2);
                        bool billHasAnyTag = !string.IsNullOrWhiteSpace(animal.TagNumber1)
                                             || !string.IsNullOrWhiteSpace(animal.TagNumber2)
                                             || !string.IsNullOrWhiteSpace(animal.Tag3);
                        if (fileHasAnyTag && billHasAnyTag
                            && !AnimalMatchesAnyFileTag(animal, backTag, tag1, tag2))
                        {
                            flagReason = $"ACN {acn} matches Ctrl No {animal.ControlNo}, but tags differ — bill has {animal.TagNumber1}/{animal.TagNumber2}/{animal.Tag3}, file has {backTag}/{tag1}/{tag2}. Possible wrong bill assignment — verify before saving.";
                            animal = null;
                            goto AddFlag;
                        }
                    }

                    
                    
                    //EID-shaped BackTag  suffix ladder 7→3 with weight tiebreaker 
                    if (animal == null && IsEidShaped(backTag))
                    {
                        var (m, cands, sfx, rsn) = FindByEidSuffix(backTag, liveWt, s1, s2, "BackTag");
                        if (m != null)
                        {
                            animal = m;
                            matchMethod = rsn; // e.g. "BackTag-EID-Last7" or "...+ClosestWeight(...)"
                        }
                        else if (cands != null && cands.Count > 0)
                        {
                            flagReason = rsn;
                            _candidateBuffer = BuildCandidates(cands, liveWt);
                            goto AddFlag;
                        }
                        // 0 hits at every length → fall through to later match steps
                    }

                     
                    // HotScale often puts the full 15-digit EID in Tag1 while DB stores
                    // only the hand-entered last 3–7 digits (e.g. '8622').
                    if (animal == null && IsEidShaped(tag1))
                    {
                        var (m, cands, sfx, rsn) = FindByEidSuffix(tag1, liveWt, s1, s2, "Tag1");
                        if (m != null)
                        {
                            animal = m;
                            matchMethod = rsn;
                        }
                        else if (cands != null && cands.Count > 0)
                        {
                            flagReason = rsn;
                            _candidateBuffer = BuildCandidates(cands, liveWt);
                            goto AddFlag;
                        }
                    }

                    //EID-shaped Tag2 — same suffix ladder 
                    if (animal == null && IsEidShaped(tag2))
                    {
                        var (m, cands, sfx, rsn) = FindByEidSuffix(tag2, liveWt, s1, s2, "Tag2");
                        if (m != null)
                        {
                            animal = m;
                            matchMethod = rsn;
                        }
                        else if (cands != null && cands.Count > 0)
                        {
                            flagReason = rsn;
                            _candidateBuffer = BuildCandidates(cands, liveWt);
                            goto AddFlag;
                        }
                    }

                    //Lot prefix - [digits][LETTERS][4digits] → strip, use suffix
                    if (animal == null && !string.IsNullOrEmpty(backTag))
                    {
                        var lotMatch = System.Text.RegularExpressions.Regex.Match(backTag, @"^(\d+)([A-Z]+)(\d{4})$");
                        if (lotMatch.Success)
                        {
                            var suffix4 = lotMatch.Groups[3].Value;
                            var vendorCode = lotMatch.Groups[2].Value;

                            var lotHitsAll = ExactTagLookup(suffix4);
                            var lotHits = PrepareTagCandidates(lotHitsAll, $"Lot prefix '{backTag}' suffix '{suffix4}'", ref flagReason);

                            if (lotHits.Count == 1)
                            {
                                
                                animal = lotHits[0];
                                matchMethod = $"LotPrefix({suffix4})";
                                _logger.LogInformation("[HW-TRACE] Lot prefix matched: backTag={Bt}, suffix={Sx}, bill={Cn}",
                                    backTag, suffix4, animal.ControlNo);
                            }
                            else if (lotHits.Count > 1)
                            {
                                var (wpick, wreason) = WeightPickFromCandidates(
                                    lotHits, liveWt, $"Lot prefix '{backTag}' suffix '{suffix4}'", vendorCode);
                                if (wpick != null)
                                {
                                    animal = wpick;
                                    matchMethod = $"LotPrefix+Score({suffix4})";
                                    flagReason = wreason;
                                }
                                else
                                {
                                    flagReason = wreason;
                                    _candidateBuffer = BuildCandidates(lotHits, liveWt);
                                    goto AddFlag;
                                }
                            }
                            else if (lotHitsAll.Count > 0)
                            {
                                goto AddFlag;
                            }
                            else
                            {
                                var sfxHitsAll = (await _animalService.GetByTagSuffixAsync(suffix4)).ToList();
                                var sfxHits = PrepareTagCandidates(
                                    sfxHitsAll, $"Lot suffix '{suffix4}' from '{backTag}'", ref flagReason);

                                if (sfxHits.Count == 1)
                                {
                                    animal = sfxHits[0];
                                    matchMethod = $"LotSuffix({suffix4})";
                                }
                                else if (sfxHits.Count > 1)
                                {
                                    var (wpick, wreason) = WeightPickFromCandidates(
                                        sfxHits, liveWt, $"Lot suffix '{suffix4}' from '{backTag}'", vendorCode);
                                    if (wpick != null)
                                    {
                                        animal = wpick;
                                        matchMethod = $"LotSuffix+Score({suffix4})";
                                        flagReason = wreason;
                                    }
                                    else
                                    {
                                        flagReason = wreason;
                                        _candidateBuffer = BuildCandidates(sfxHits, liveWt);
                                        goto AddFlag;
                                    }
                                }
                                else if (sfxHitsAll.Count > 0)
                                {
                                    goto AddFlag;
                                }
                            }
                        }
                        else
                        {
                            var nonStdLot = System.Text.RegularExpressions.Regex.Match(backTag, @"^(\d+)([A-Z]+)(\d{1,3}|\d{5,})$");
                            if (nonStdLot.Success)
                            {
                                flagReason = $"Lot prefix '{backTag}' has non-4-digit suffix — manual review required";
                                goto AddFlag;
                            }
                        }
                    }

                    //Exact BackTag match (with leading-zero variant) 
                    if (animal == null && !string.IsNullOrEmpty(backTag) && !backTag.Contains('?'))
                    {
                        var btHitsAll = ExactTagLookup(backTag);
                        var btHits = PrepareTagCandidates(btHitsAll, $"BackTag '{backTag}'", ref flagReason);

                        if (btHits.Count == 1)
                        {
                            animal = btHits[0];
                            matchMethod = "BackTag";
                        }
                        else if (btHits.Count > 1)
                        {
                            var (wpick, wreason) = WeightPickFromCandidates(btHits, liveWt, $"BackTag '{backTag}'");
                            if (wpick != null)
                            {
                                animal = wpick;
                                matchMethod = $"BackTag+Weight({backTag})";
                                flagReason = wreason;
                            }
                            else
                            {
                                flagReason = wreason;
                                _candidateBuffer = BuildCandidates(btHits, liveWt);
                                goto AddFlag;
                            }
                        }
                        else if (btHitsAll.Count > 0)
                        {
                            goto AddFlag;
                        }
                    }

                    // Alpha-prefix tag try exact then numeric suffix 
                    if (animal == null && !string.IsNullOrEmpty(backTag))
                    {
                        var alphaMatch = System.Text.RegularExpressions.Regex.Match(backTag, @"^([A-Z]+)(\d+)$");
                        if (alphaMatch.Success)
                        {
                            var numPart = alphaMatch.Groups[2].Value;
                            var alphaHitsAll = ExactTagLookup(numPart);
                            var alphaHits = PrepareTagCandidates(alphaHitsAll, $"Alpha '{backTag}'", ref flagReason);

                            if (alphaHits.Count == 1)
                            {
                                animal = alphaHits[0];
                                matchMethod = $"AlphaNum({numPart})";
                            }
                            else if (alphaHits.Count > 1)
                            {
                                var (wpick, wreason) = WeightPickFromCandidates(alphaHits, liveWt, $"Alpha '{backTag}'");
                                if (wpick != null)
                                {
                                    animal = wpick;
                                    matchMethod = $"Alpha+Weight({numPart})";
                                    flagReason = wreason;
                                }
                                else
                                {
                                    flagReason = wreason;
                                    _candidateBuffer = BuildCandidates(alphaHits, liveWt);
                                    goto AddFlag;
                                }
                            }
                            else if (alphaHitsAll.Count > 0)
                            {
                                goto AddFlag;
                            }
                        }
                    }

                    // Tag1/Tag2 from file (ear/rt/slow = tag note, still search) 
                    if (animal == null && !string.IsNullOrEmpty(tag1))
                    {
                        var skipWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ear", "rt", "slow", "nt", "none", "no tag" };
                        if (!skipWords.Contains(tag1))
                        {
                            var t1HitsAll = ExactTagLookup(tag1);
                            var t1Hits = PrepareTagCandidates(t1HitsAll, $"Tag1 '{tag1}'", ref flagReason);

                            if (t1Hits.Count == 1)
                            {
                                animal = t1Hits[0];
                                matchMethod = "Tag1";
                            }
                            else if (t1Hits.Count > 1)
                            {
                                var (wpick, wreason) = WeightPickFromCandidates(t1Hits, liveWt, $"Tag1 '{tag1}'");
                                if (wpick != null)
                                {
                                    animal = wpick;
                                    matchMethod = "Tag1+Weight";
                                    flagReason = wreason;
                                }
                                else
                                {
                                    flagReason = wreason;
                                    _candidateBuffer = BuildCandidates(t1Hits, liveWt);
                                    goto AddFlag;
                                }
                            }
                            else if (t1HitsAll.Count > 0)
                            {
                                goto AddFlag;
                            }
                        }
                    }

                    if (animal == null && !string.IsNullOrEmpty(tag2))
                    {
                        var t2HitsAll = ExactTagLookup(tag2);
                        var t2Hits = PrepareTagCandidates(t2HitsAll, $"Tag2 '{tag2}'", ref flagReason);

                        if (t2Hits.Count == 1)
                        {
                            animal = t2Hits[0];
                            matchMethod = "Tag2";
                        }
                        else if (t2Hits.Count > 1)
                        {
                            var (wpick, wreason) = WeightPickFromCandidates(t2Hits, liveWt, $"Tag2 '{tag2}'");
                            if (wpick != null)
                            {
                                animal = wpick;
                                matchMethod = "Tag2+Weight";
                                flagReason = wreason;
                            }
                            else
                            {
                                flagReason = wreason;
                                _candidateBuffer = BuildCandidates(t2Hits, liveWt);
                                goto AddFlag;
                            }
                        }
                        else if (t2HitsAll.Count > 0)
                        {
                            goto AddFlag;
                        }
                    }

                    //Wildcard 
                    if (animal == null && !string.IsNullOrEmpty(backTag) && backTag.Contains('?'))
                    {
                    var zeroTag = backTag.Replace("?", "0");
                    if (!string.IsNullOrWhiteSpace(zeroTag))
                    {
                    var zHitsAll = ExactTagLookup(zeroTag);
                    var zHits = PrepareTagCandidates(zHitsAll, "ZeroSub '" + backTag + "' -> '" + zeroTag + "'", ref flagReason);
                        if (zHits.Count == 1)
                        {
                            animal = zHits[0];
                            matchMethod = "ZeroSub(" + backTag + "->" + zeroTag + ")";
                        }
                        else if (zHits.Count > 1)
                        {
                            var picked = WeightPickFromCandidates(zHits, liveWt, "ZeroSub '" + backTag + "' -> '" + zeroTag + "'");
                            if (picked.animal != null)
                            {
                                animal = picked.animal;
                                matchMethod = "ZeroSub+Weight(" + backTag + "->" + zeroTag + ")";
                                flagReason = picked.reason;
                            }
                            else
                            {
                                flagReason = picked.reason;
                                _candidateBuffer = BuildCandidates(zHits, liveWt);
                                goto AddFlag;
                            }
                        }
                        else if (zHitsAll.Count > 0)
                        {
                            goto AddFlag;
                        }
                    }
                    }
                    // Wildcard ? match 
                    if (animal == null && !string.IsNullOrEmpty(backTag) && backTag.Contains('?'))
                    {
                        var wcHitsAll = (await _animalService.GetByTagPatternAsync(backTag)).ToList();
                        var wcHits = PrepareTagCandidates(wcHitsAll, $"Wildcard '{backTag}'", ref flagReason);

                        if (wcHits.Count == 1)
                        {
                            animal = wcHits[0];
                            matchMethod = $"Wildcard({backTag})";
                        }
                        else if (wcHits.Count == 2)
                        {
                            var (wpick, wreason) = WeightPickFromCandidates(wcHits, liveWt, $"Wildcard '{backTag}'");
                            if (wpick != null)
                            {
                                animal = wpick;
                                matchMethod = $"Wildcard+Weight({backTag})";
                                flagReason = wreason;
                            }
                            else
                            {
                                flagReason = wreason;
                                _candidateBuffer = BuildCandidates(wcHits, liveWt);
                                goto AddFlag;
                            }
                        }
                        else if (wcHits.Count > 2)
                        {
                            flagReason = $"Wildcard '{backTag}' matches {wcHits.Count} unassigned animals";
                            _candidateBuffer = BuildCandidates(wcHits, liveWt);
                            goto AddFlag;
                        }
                        else if (wcHitsAll.Count > 0)
                        {
                            goto AddFlag;
                        }
                    }

                    //  Live Weight proximity fallback 
                    if (animal == null && liveWt.HasValue && liveWt.Value > 0)
                    {
                        if (pendingAnimals == null)
                            pendingAnimals = (await _animalService.GetAllPendingAsync()).ToList();

                        var candidates = pendingAnimals
                        .Where(a =>
                        !matchedControlNos.Contains(a.ControlNo) &&
                        a.LiveWeight > 0 &&
                        IsAcnMissing(a.AnimalControlNumber))
                        .Select(a => ValueTuple.Create(a, Math.Abs(a.LiveWeight - liveWt.Value)))
                        .OrderBy(x => x.Item2)
                        .ToList();

                        if (candidates.Count > 0)
                        {
                            var bestAnimal = candidates[0].Item1;
                            var bestDiff   = candidates[0].Item2;
                            var secondDiff = candidates.Count > 1 ? candidates[1].Item2 : (decimal)9999;

                            if (bestDiff <= 200 && (secondDiff - bestDiff) > 50)
                            {
                                matchMethod = $"LiveWeight({liveWt:N0}→DB:{bestAnimal.LiveWeight:N0},diff:{bestDiff:N0})";
                                flagReason  = $"Weight-matched: file LiveWt {liveWt:N0} lbs → DB unassigned animal LiveWt {bestAnimal.LiveWeight:N0} lbs (diff {bestDiff:N0} lbs) — please confirm";

                                vm.FlaggedRows.Add(new HotWeightPreviewRow
                                {
                                    AnimalControlNumber = acn,
                                    NewAnimalControlNumber = !string.IsNullOrWhiteSpace(acn) ? acn : null,
                                    Side1 = s1,
                                    Side2 = s2,
                                    NewGrade = grade,
                                    NewGrade2 = grade2,
                                    NewHealthScore = hs,
                                    FileLiveWeight = liveWt,
                                    FileLot = fileLot,
                                    FileSex = fileSex,
                                    FileType = fileType,
                                    FileOrigin = fileOrigin,
                                    FileBackTag = backTag,
                                    FileTag1 = tag1,
                                    FileTag2 = tag2,
                                    FileProgram = fileProgram,
                                    Status = "Flag",
                                    FlagReason = flagReason,
                                    Candidates = _candidateBuffer
                                });
                                //matchedControlNos.Add(bestAnimal.ControlNo);
                                continue;
                            }
                            else if (bestDiff <= 200)
                            {
                                flagReason = $"Weight match ambiguous among unassigned animals — best diff {bestDiff:N0} lbs, second {secondDiff:N0} lbs for LiveWt {liveWt:N0}";
                            }
                            else
                            {
                                flagReason = $"No unassigned weight match within 200 lbs — file LiveWt {liveWt:N0}, closest DB {bestAnimal.LiveWeight:N0} lbs";
                            }
                        }
                        else
                        {
                            var hasAssignedOnly = pendingAnimals.Any(a =>
                            !matchedControlNos.Contains(a.ControlNo) &&
                            a.LiveWeight > 0 &&
                            !IsAcnMissing(a.AnimalControlNumber));

                            flagReason = hasAssignedOnly
                                ? "Weight matching found only already-assigned ACN records — manual review required"
                                : "No pending animals available for weight matching";
                        }
                        goto AddFlag;
                    }

                    if (animal == null)
                    {
                        var tried = $"ACN='{acn}', BackTag='{backTag}', Tag1='{tag1}'";
                        if (string.IsNullOrEmpty(flagReason))
                            flagReason = $"No tag or weight match — tried {tried}";
                        goto AddFlag;
                    }

                    _logger.LogInformation("[HW-TRACE] About to goto DoneMatch: ControlNo={Cn}, backTag={Bt}, matchMethod={Mm}",
    animal?.ControlNo ?? -1, backTag, matchMethod);

                    goto DoneMatch;
                    _logger.LogInformation("[HW-TRACE] At DoneMatch: ControlNo={Cn}, backTag={Bt}, matchMethod={Mm}, animalNull={Null}",
    animal?.ControlNo ?? -1, backTag, matchMethod, animal == null);
                    AddFlag:
                    // Condemnation detection runs here too, not just for matched rows.
                    // Without this, a flagged condemned animal (Grade or Grade2 starts
                    // with X, sides = 0) lands in FlaggedRows with IsCondemned=false,
                    // and the JS validator falsely demands Side1/Side2/Grade/HS.
                    bool flaggedIsCondemned = GradeRules.IsCondemnationCode(grade)
                                              || GradeRules.IsCondemnationCode(grade2);

                    vm.FlaggedRows.Add(new HotWeightPreviewRow
                    {
                        AnimalControlNumber = acn, Side1 = s1, Side2 = s2,
                        NewGrade       = grade,
                        NewGrade2      = grade2,
                        NewHealthScore = hs,
                        FileLiveWeight = liveWt,
                        FileLot        = fileLot,
                        FileSex        = fileSex,
                        FileType       = fileType,
                        FileOrigin     = fileOrigin,
                        FileBackTag    = backTag,
                        FileTag1       = tag1,
                        FileTag2       = tag2,
                        FileProgram    = fileProgram,
                        IsCondemned    = flaggedIsCondemned,
                        Status         = "Flag",
                        FlagReason     = flagReason,
                        // Store candidates if this is a multi-match flag (for picker UI)
                        Candidates     = _candidateBuffer
                    });
                    _candidateBuffer = null;
                    continue;
                    DoneMatch: 
                     _logger.LogInformation("[HW-TRACE] At DoneMatch: ControlNo={Cn}, backTag={Bt}, matchMethod={Mm}, animalNull={Null}",
                animal?.ControlNo ?? -1, backTag, matchMethod, animal == null);
                    // Dedup - silently skip if we already matched this animal
                    // (hot scale machine sometimes writes 2 rows per animal - both have same BackTag)
                    if (matchedControlNos.Contains(animal.ControlNo))
                    {
                        vm.TotalInExcel--;   // don't count the silent duplicate in total
                        continue;
                    }
                    matchedControlNos.Add(animal.ControlNo);

                    // Already-complete check: ACN assigned + HotWeight saved + tag matches BackTag → Duplicate
                    var fileAcnNorm = NormalizeAcn(acn)?.TrimStart('0');
                    var dbAcnNorm   = NormalizeAcn(animal.AnimalControlNumber)?.TrimStart('0');

                    bool acnDirectMatch = !string.IsNullOrWhiteSpace(fileAcnNorm)
                                    && !string.IsNullOrWhiteSpace(dbAcnNorm)
                                    && string.Equals(fileAcnNorm, dbAcnNorm, StringComparison.OrdinalIgnoreCase);

                    bool anyTagMatch = AnimalMatchesAnyFileTag(animal, backTag, tag1, tag2);

                    bool alreadyComplete = !IsAcnMissing(animal.AnimalControlNumber)
                                        && animal.HotWeight.HasValue
                                        && animal.HotWeight.Value > 0
                                        && (acnDirectMatch || anyTagMatch);

                    if (alreadyComplete)
                    {
                        var dbAcn = (animal.AnimalControlNumber ?? "").TrimStart('0');
                        var dbHw  = animal.HotWeight!.Value;
                        vm.DupRows.Add(new HotWeightPreviewRow
                        {
                            ControlNo           = animal.ControlNo,
                            AnimalControlNumber = dbAcn,
                            CurrentHotWeight    = dbHw.ToString("N1"),
                            CurrentGrade        = animal.Grade,
                            CurrentHealthScore  = animal.HealthScore?.ToString(),
                            Side1 = s1, Side2 = s2,
                            NewGrade       = grade,
                            NewGrade2      = grade2,
                            NewHealthScore = hs,
                            FileLiveWeight = liveWt,
                            FileLot        = fileLot,
                            FileSex        = fileSex,
                            FileType       = fileType,
                            FileOrigin     = fileOrigin,
                            FileBackTag    = backTag,
                            FileTag1       = tag1,
                            FileTag2       = tag2,
                            FileProgram    = fileProgram,
                            MatchMethod    = matchMethod,
                            Status         = "Dup",
                            FlagReason     = $"Already complete — ACN {dbAcn}, HotWeight {dbHw:N1} lbs already in DB",
                        });
                        vm.AlreadyHasData++;
                        continue;
                    }


                    // Use the matched animal's ACN (or write from Excel if tag-matched)
                    // Bug fix: also fall back to acn when bill ACN is whitespace-only
                    // OR when it strips to empty after removing leading zeros (e.g. "0", "0000").
                    string? trimmedBillAcn = !string.IsNullOrWhiteSpace(animal.AnimalControlNumber)
                        ? animal.AnimalControlNumber.Trim().TrimStart('0')
                        : null;
                    var resolvedAcn = !string.IsNullOrEmpty(trimmedBillAcn)
                        ? trimmedBillAcn
                        : acn;

                    vm.Matched++;
                    _logger.LogInformation("[HW-TRACE] Matched++ fired: ControlNo={Cn}, vm.Matched now={M}, matchMethod={Mm}",
                        animal.ControlNo, vm.Matched, matchMethod);

                    var row = new HotWeightPreviewRow
                    {
                        ControlNo            = animal.ControlNo,
                        AnimalControlNumber  = resolvedAcn,
                        CurrentHotWeight     = animal.HotWeight.HasValue ? animal.HotWeight.Value.ToString("N1") : null,
                        CurrentGrade         = animal.Grade,
                        CurrentHealthScore   = animal.HealthScore.HasValue ? animal.HealthScore.Value.ToString() : null,
                        Side1 = s1, Side2 = s2,
                        NewGrade       = grade,
                        NewGrade2      = grade2,
                        NewHealthScore = hs,
                        FileLiveWeight = liveWt,
                        FileLot        = fileLot,
                        FileSex        = fileSex,
                        FileType       = fileType,
                        FileOrigin     = fileOrigin,
                        MatchMethod    = matchMethod,
                        NewAnimalControlNumber = (matchMethod != "ACN" && !string.IsNullOrEmpty(acn)) ? acn : null,
                        FileBackTag    = backTag,
                        FileTag1       = tag1,
                        FileTag2       = tag2,
                        FileProgram    = fileProgram,
                    };

                    var flags = new List<string>();

                    // Condemnation detection: any grade starting with X means the
                    // animal was condemned and didn't make it through processing.
                    // Condemned rows have no carcass weight and bypass side, grade
                    // allow-list, and HealthScore validation. The IsCondemned flag
                    // pre-checks the Condemned checkbox at Mark Killed load time.
                    bool isCondemned = GradeRules.IsCondemnationCode(grade)
                                       || GradeRules.IsCondemnationCode(grade2);
                    row.IsCondemned = isCondemned;

                    // Side checks — skipped for condemned rows.
                    bool side1Ok = s1.HasValue && s1.Value > 0;
                    bool side2Ok = s2.HasValue && s2.Value > 0;

                    if (!isCondemned)
                    {
                        if (!side1Ok && !side2Ok)
                            flags.Add("Both sides missing");
                        else if (!side1Ok)
                            flags.Add("Side1 missing");
                        else if (!side2Ok)
                            flags.Add("Side2 missing");
                        else
                        {
                            // Both sides present — check for zero and variance
                            if (s1!.Value == 0) flags.Add("Side1 is zero");
                            if (s2!.Value == 0) flags.Add("Side2 is zero");

                            if (s1!.Value > 0 && s2!.Value > 0)
                            {
                                var variance = Math.Abs(s1!.Value - s2!.Value) / ((s1!.Value + s2!.Value) / 2);
                                if (variance > 0.05m)
                                {
                                    // NOT flagged - set trim comment and auto-load
                                    // LTrim = left side (Side1) was trimmed more; RTrim = right side (Side2)
                                    row.TrimComment = s1!.Value < s2!.Value ? "LTrim" : "RTrim";
                                }
                            }
                        }
                    }

                    // Grade validation (Sex primary, Type secondary fallback)
                    var ruleInfo = ResolveGradeRules(fileType, fileSex);
                    var allowedGrades = ruleInfo.AllowedGrades;

                    if (ruleInfo.HasMismatch)
                    {
                        flags.Add($"Sex and Type mismatch: Sex={ruleInfo.SexCode}, Type={ruleInfo.TypeCode}");
                    }

                    var g1 = string.IsNullOrWhiteSpace(grade) ? null : grade.Trim().ToUpperInvariant();
                    var g2 = string.IsNullOrWhiteSpace(grade2) ? null : grade2.Trim().ToUpperInvariant();

                    bool g1Valid = allowedGrades != null && g1 != null && allowedGrades.Contains(g1);
                    bool g2Valid = allowedGrades != null && g2 != null && allowedGrades.Contains(g2);

                    // Use primary Grade when valid; otherwise accept Grade2 as correction.
                    if (g1Valid)
                        row.NewGrade = g1;
                    else if (g2Valid)
                        row.NewGrade = g2;
                    else if (!string.IsNullOrWhiteSpace(g1))
                        row.NewGrade = g1;
                    else
                        row.NewGrade = g2;

                    if (!isCondemned)
                    {
                        if (g1 == null && g2 == null)
                        {
                            flags.Add("Grade or Grade2 required");
                        }
                        else if (allowedGrades != null)
                        {
                            // Primary Grade invalid and Grade2 missing/invalid => flag (review only, never hard-block).
                            if (!g1Valid && !g2Valid)
                            {
                                flags.Add($"Invalid grade for Sex={ruleInfo.SexCode}, Type={ruleInfo.TypeCode}. Grade={g1 ?? "-"}, Grade2={g2 ?? "-"}");
                            }
                        }
                    }

                    // HealthScore validation — skipped for condemned rows.
                    if (!isCondemned)
                    {
                        if (!hs.HasValue)
                            flags.Add("HealthScore missing");
                        else if (hs.Value < 1 || hs.Value > 5)
                            flags.Add($"HealthScore {hs} out of range 1–5");
                    }

                    // Compute HotWeight only when both sides valid
                    bool sidesClean = side1Ok && side2Ok && s1!.Value > 0 && s2!.Value > 0;
                    if (sidesClean)
                        row.NewHotWeight = s1!.Value + s2!.Value;

                    // Duplicate detection 
                    // A row is a DUPLICATE if the animal already has data in DB AND
                    // HotWeight + Grade + HealthScore + ACN all match the incoming file exactly.
                    // If any field differs -> not a dup -> goes to Ready (allow overwrite).
                    bool hasExisting = animal.HotWeight.HasValue && animal.HotWeight.Value > 0;
                    if (hasExisting && !flags.Any())
                    {
                        var incomingHW    = row.NewHotWeight ?? 0;
                        var dbHW          = animal.HotWeight ?? 0;
                        var incomingGrade = (row.NewGrade ?? "").Trim().ToUpper();
                        var dbGrade       = (animal.Grade  ?? "").Trim().ToUpper();
                        var incomingHS    = hs ?? 0;
                        var dbHS          = animal.HealthScore ?? 0;
                        var incomingACN   = (acn ?? "").TrimStart('0');
                        var dbACN         = (animal.AnimalControlNumber ?? "").TrimStart('0');

                        bool hwMatch    = Math.Abs(incomingHW - dbHW) < 0.1m;
                        bool gradeMatch = string.Equals(incomingGrade, dbGrade, StringComparison.OrdinalIgnoreCase);
                        bool hsMatch    = incomingHS == dbHS;
                        bool acnMatch   = string.IsNullOrEmpty(incomingACN) || string.IsNullOrEmpty(dbACN)
                                          || string.Equals(incomingACN, dbACN, StringComparison.OrdinalIgnoreCase);

                        if (hwMatch && gradeMatch && hsMatch && acnMatch)
                        {
                            // Exact duplicate - route to Dup tab
                            row.Status    = "Dup";
                            row.FlagReason = $"Duplicate — DB already has HotWeight {dbHW:N1} lbs, Grade {dbGrade}, HS {dbHS}";
                            vm.DupRows.Add(row);
                            vm.AlreadyHasData++;
                            matchedControlNos.Add(animal.ControlNo);
                            continue;
                        }
                        else
                        {
                            // Data changed - allow overwrite - note what's different
                            var diffs = new List<string>();
                            if (!hwMatch)    diffs.Add($"HW {dbHW:N1}→{incomingHW:N1}");
                            if (!gradeMatch) diffs.Add($"Grade {dbGrade}→{incomingGrade}");
                            if (!hsMatch)    diffs.Add($"HS {dbHS}→{incomingHS}");
                            row.FlagReason = $"Updating existing record ({string.Join(", ", diffs)})";
                            vm.AlreadyHasData++;
                            // Falls through to AutoRows below
                        }
                    }

                    // Apply TrimComment
                    if (!string.IsNullOrEmpty(row.TrimComment) && !flags.Any() && string.IsNullOrEmpty(row.FlagReason))
                    {
                        row.FlagReason = row.TrimComment;
                    }

                    if (flags.Any())
                    {
                        row.Status = "Flag";
                        row.FlagReason = string.Join("; ", flags);
                        if (!string.IsNullOrEmpty(row.TrimComment))
                            row.FlagReason += $"; {row.TrimComment}";
                        vm.FlaggedRows.Add(row);
                    }
                    else
                    {
                        // Final completeness gate — mirrors IsCompleteForKill in the
                        // MarkKilledApi save endpoint. A row should only be marked
                        // Ready if it would actually pass save-time validation.
                        // Without this gate, "ready" rows with missing HW/Grade/HS
                        // get rejected at save time, producing the
                        // "N rows had errors and were NOT saved" warning.
                        var savabilityFlags = new List<string>();

                        // ACN must resolve to a non-empty value (matches NormalizeAcn).
                        if (string.IsNullOrWhiteSpace(row.AnimalControlNumber))
                            savabilityFlags.Add("ACN missing — required for save");

                        if (!isCondemned)
                        {
                            // Production-killed animals need full weight, grade, and HS.
                            if (!row.NewHotWeight.HasValue || row.NewHotWeight.Value <= 0)
                                savabilityFlags.Add("HotWeight missing or zero");
                            if (string.IsNullOrWhiteSpace(row.NewGrade))
                                savabilityFlags.Add("Grade missing");
                            if (!row.NewHealthScore.HasValue
                                || row.NewHealthScore.Value < 1
                                || row.NewHealthScore.Value > 5)
                                savabilityFlags.Add("HealthScore missing or out of 1–5");
                        }

                        if (savabilityFlags.Any())
                        {
                            row.Status     = "Flag";
                            row.FlagReason = string.Join("; ", savabilityFlags);
                            if (!string.IsNullOrEmpty(row.TrimComment))
                                row.FlagReason += $"; {row.TrimComment}";
                            vm.FlaggedRows.Add(row);
                            _logger.LogInformation("[HW-TRACE] Sent to FlaggedRows (savability): ControlNo={Cn}, ACN={Acn}, flags={Flags}",
                                row.ControlNo, row.AnimalControlNumber, string.Join("|", savabilityFlags));
                        }
                        else
                        {
                            row.Status = "OK";
                            vm.AutoRows.Add(row);
                            _logger.LogInformation("[HW-TRACE] Added to AutoRows: ControlNo={Cn}, ACN={Acn}, NewHotWeight={Hw}",
                            row.ControlNo, row.AnimalControlNumber, row.NewHotWeight);
                        }
                    }
                }
                var hwJson = System.Text.Json.JsonSerializer.Serialize(vm);
                _logger.LogInformation(
                    "[HW-IMPORT] Parsed: TotalInExcel={Total}, AutoRows={Auto} (OK={Ok}, Loaded={Loaded}), FlaggedRows={Flagged}, DupRows={Dups}, Matched={Matched}, JsonLength={Len}. Sum={Sum}.",
                    vm.TotalInExcel,
                    vm.AutoRows.Count,
                    vm.AutoRows.Count(r => r.Status == "OK"),
                    vm.AutoRows.Count(r => r.Status == "Loaded"),
                    vm.FlaggedRows.Count,
                    vm.DupRows.Count,
                    vm.Matched,
                    hwJson.Length,
                    vm.AutoRows.Count + vm.FlaggedRows.Count + vm.DupRows.Count);
                _logger.LogInformation("[HW-IMPORT] Sample AutoRow ControlNos: {Ids}",
                    string.Join(", ", vm.AutoRows.Take(5).Select(r => $"{r.ControlNo}={r.NewHotWeight}")));
            }
            catch (Exception ex)
            {
                vm.Errors.Add($"Parse error: {ex.Message}");
                return vm;
            }

            return vm;
        }

        // HOT WEIGHT - LOAD SELECTED ROWS INTO MARK AS KILLED (no DB write yet)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> HotWeightLoadToMarkKilled(string? selectedControlNos, string? fixedFlaggedJson)
        {
            // Read from SHARED staging — collaborative review across all teammates.
            var json = await StagingBridge.ReadSharedAsync(
                _stagingService,
                StagingBridge.Types.HotWeight,
                StagingBridge.SharedHotWeightKey);
            if (string.IsNullOrEmpty(json))
            { TempData["ErrorMessage"] = "Preview session expired. Please re-upload the file or refresh from Hot Scale."; return RedirectToAction(nameof(HotWeightImport)); }

            if (string.IsNullOrWhiteSpace(selectedControlNos))
            {
            TempData["ErrorMessage"] = "No rows were selected to load.";
            return RedirectToAction(nameof(HotWeightImport));
            }

            // If specific control numbers were posted, filter the session VM to only those rows
            if (!string.IsNullOrEmpty(selectedControlNos))
            {
                try
                {
                    var selectedIds = selectedControlNos.Split(',', StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => int.TryParse(s.Trim(), out int id) ? id : 0)
                        .Where(id => id > 0).ToHashSet();

                    if (selectedIds.Any())
                    {
                        var vm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(json);
                        if (vm != null)
                        {
                            vm.AutoRows    = vm.AutoRows.Where(r => selectedIds.Contains(r.ControlNo)).ToList();
                            vm.FlaggedRows = vm.FlaggedRows.Where(r => selectedIds.Contains(r.ControlNo)).ToList();
                            json = System.Text.Json.JsonSerializer.Serialize(vm);
                        }
                    }
                }
                catch { /* use full session on error */ }
            }

            var fixedRowsParsed = new List<FixedFlaggedRow>();

            // Merge any fixed flagged rows from the JS into the vm
            if (!string.IsNullOrEmpty(fixedFlaggedJson))
            {
                try
                {
                    var fixedRows = System.Text.Json.JsonSerializer.Deserialize<List<FixedFlaggedRow>>(fixedFlaggedJson);
                    fixedRowsParsed = fixedRows;
                    if (fixedRows != null && fixedRows.Any())
                    {
                        var vmForMerge = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(json);
                        if (vmForMerge != null)
                        {
                            foreach (var fix in fixedRows)
                            {
                                var existing = !string.IsNullOrWhiteSpace(fix.RowKey)
                                    ? vmForMerge.FlaggedRows.FirstOrDefault(r => r.RowKey == fix.RowKey)
                                    : vmForMerge.FlaggedRows.FirstOrDefault(r => r.ControlNo == fix.ControlNo);

                                if (existing != null)
                                {
                                    existing.ControlNo = fix.ControlNo;
                                    if (!string.IsNullOrWhiteSpace(fix.AnimalControlNumber))
                                    existing.AnimalControlNumber = fix.AnimalControlNumber.Trim();
                                    // Guard each overwrite — picker-only flow sends zeros.
                                    if (fix.Side1 > 0) existing.Side1 = fix.Side1;
                                    if (fix.Side2 > 0) existing.Side2 = fix.Side2;
                                    if (fix.Side1 + fix.Side2 > 0)
                                        existing.NewHotWeight = fix.Side1 + fix.Side2;
                                    if (!string.IsNullOrWhiteSpace(fix.Grade))
                                        existing.NewGrade = fix.Grade;
                                    if (fix.HealthScore > 0)
                                        existing.NewHealthScore = fix.HealthScore;
                                    existing.Status = "OK";
                                    existing.FlagReason = "";
                                    existing.IsManuallyEdited = true; // smart-merge: preserve this fix across auto-refreshes
                                    vmForMerge.FlaggedRows.Remove(existing);
                                    vmForMerge.AutoRows.Add(existing);
                                }
                            }
                            json = System.Text.Json.JsonSerializer.Serialize(vmForMerge);
                            _logger.LogInformation("[HW-LOAD] Merged {Count} fixed flagged rows into AutoRows", fixedRows.Count);
                        }
                    }
                }
                catch (Exception ex) { _logger.LogWarning(ex, "[HW-LOAD] Could not merge fixed flagged rows"); }
            }

            // Don't write session yet — we want session to hold the FULL post-Load
            // master (with all FlaggedRows, DupRows, and Loaded markers), not the
            // pre-Load filtered subset. We write below after the master is updated.
            TempData["HWLoaded"]  = "1";
            HttpContext.Session.SetString("HWLoaded", "1");

            
            //  merge fixed flagged rows, then mark loaded rows --
           var masterJson = await StagingBridge.ReadSharedAsync(
                  _stagingService,
                  StagingBridge.Types.HotWeight,
                  StagingBridge.SharedHotWeightKey);
            if (!string.IsNullOrEmpty(masterJson))
            {
                try
                {
                    var masterVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(masterJson);
                    if (masterVm != null)
                    {
                        // 1) Merge fixed flagged rows into master (same logic as vmForMerge)
                        if (fixedRowsParsed.Any())
                        {
                            foreach (var fix in fixedRowsParsed)
                            {
                                var existing = !string.IsNullOrWhiteSpace(fix.RowKey)
                                    ? masterVm.FlaggedRows.FirstOrDefault(r => r.RowKey == fix.RowKey)
                                    : masterVm.FlaggedRows.FirstOrDefault(r => r.ControlNo == fix.ControlNo);

                                if (existing == null) continue;

                                existing.ControlNo = fix.ControlNo;

                                if (!string.IsNullOrWhiteSpace(fix.AnimalControlNumber))
                                    existing.AnimalControlNumber = fix.AnimalControlNumber.Trim();

                                // Only overwrite if the fix has a meaningful value.
                                // The picker-only flow (operator picked a candidate but
                                // didn't edit Side/Grade/HS) sends zeros and empties —
                                // without these guards, the original Hot Scale values
                                // would be wiped, leaving a "ready" row with no HW data.
                                if (fix.Side1 > 0) existing.Side1 = fix.Side1;
                                if (fix.Side2 > 0) existing.Side2 = fix.Side2;
                                if (fix.Side1 + fix.Side2 > 0)
                                    existing.NewHotWeight = fix.Side1 + fix.Side2;
                                if (!string.IsNullOrWhiteSpace(fix.Grade))
                                    existing.NewGrade = fix.Grade;
                                if (fix.HealthScore > 0)
                                    existing.NewHealthScore = fix.HealthScore;

                                existing.Status = "OK";
                                existing.FlagReason = "";
                                existing.IsManuallyEdited = true; // smart-merge: preserve this fix across auto-refreshes

                                masterVm.FlaggedRows.Remove(existing);

                                // Avoid duplicate AutoRows entries
                                var alreadyInAuto = masterVm.AutoRows.Any(r =>
                                    (!string.IsNullOrWhiteSpace(existing.RowKey) && r.RowKey == existing.RowKey) ||
                                    (existing.ControlNo > 0 && r.ControlNo == existing.ControlNo));

                                if (!alreadyInAuto)
                                    masterVm.AutoRows.Add(existing);
                            }
                        }

                        // Final completeness re-check. Any AutoRow that fails the
                        // save-time validation (no HotWeight, no Grade, no HS, no ACN)
                        // goes back to FlaggedRows with a clear reason. Without this,
                        // picker-only fixes can land rows in Ready that look complete
                        // but actually fail at save time → "0 rows updated" surprise.
                        var demote = masterVm.AutoRows.Where(r =>
                        {
                            if (r.IsCondemned) return false; // condemned bypasses HW/Grade/HS
                            if (string.IsNullOrWhiteSpace(r.AnimalControlNumber)) return true;
                            if (!r.NewHotWeight.HasValue || r.NewHotWeight.Value <= 0) return true;
                            if (string.IsNullOrWhiteSpace(r.NewGrade)) return true;
                            if (!r.NewHealthScore.HasValue
                                || r.NewHealthScore.Value < 1
                                || r.NewHealthScore.Value > 5) return true;
                            return false;
                        }).ToList();

                        foreach (var d in demote)
                        {
                            masterVm.AutoRows.Remove(d);
                            d.Status = "Flag";
                            var reasons = new List<string>();
                            if (string.IsNullOrWhiteSpace(d.AnimalControlNumber)) reasons.Add("ACN missing");
                            if (!d.NewHotWeight.HasValue || d.NewHotWeight.Value <= 0) reasons.Add("HotWeight missing or zero");
                            if (string.IsNullOrWhiteSpace(d.NewGrade)) reasons.Add("Grade missing");
                            if (!d.NewHealthScore.HasValue || d.NewHealthScore.Value < 1 || d.NewHealthScore.Value > 5)
                                reasons.Add("HealthScore missing or out of 1–5");
                            d.FlagReason = string.Join("; ", reasons);
                            // Avoid duplicates in FlaggedRows
                            if (!masterVm.FlaggedRows.Any(f => f.ControlNo == d.ControlNo))
                                masterVm.FlaggedRows.Add(d);
                        }
                        if (demote.Count > 0)
                            _logger.LogInformation("[HW-LOAD] Demoted {Count} incomplete rows back to FlaggedRows post-merge.", demote.Count);

                        // 2) Determine which ControlNos to load
                        HashSet<int> loadedIds;
                        if (!string.IsNullOrEmpty(selectedControlNos))
                        {
                            loadedIds = selectedControlNos.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(s => int.TryParse(s.Trim(), out int id) ? id : 0)
                                .Where(id => id > 0)
                                .ToHashSet();
                        }
                        else
                        {
                            loadedIds = masterVm.AutoRows
                                .Where(r => r.Status == "OK")
                                .Select(r => r.ControlNo)
                                .ToHashSet();
                        }

                        // 3) Persist HW data (HotWeight, Grade, HS, ACN) onto the
                        //    matching bills via SaveKillDataAsync. KillStatus is NOT
                        //    touched — bills stay Pending until operator clicks
                        //    "Mark all complete (all pages)" on the Mark Killed page.
                        //
                        //    This step is the meaningful work of the Load button:
                        //    it makes the HW data visible on the bill record itself,
                        //    so it survives session loss / browser close, and the
                        //    Hot Weight page can mark these rows as "Loaded" (purple)
                        //    so they don't reappear as fresh Ready rows on the next refresh.
                        var rowsToPersist = masterVm.AutoRows
                            .Where(r => loadedIds.Contains(r.ControlNo)
                                        && r.ControlNo > 0
                                        && (r.NewHotWeight.HasValue
                                            || !string.IsNullOrWhiteSpace(r.NewGrade)
                                            || r.NewHealthScore.HasValue
                                            || !string.IsNullOrWhiteSpace(r.NewAnimalControlNumber)
                                            || r.IsCondemned))
                            .ToList();

                        int billsUpdated = 0;
                        if (rowsToPersist.Any())
                        {
                            try
                            {
                                var killData = rowsToPersist.Select(r => new KillAnimalData
                                {
                                    ControlNo           = r.ControlNo,
                                    AnimalControlNumber = !string.IsNullOrWhiteSpace(r.NewAnimalControlNumber)
                                                            ? r.NewAnimalControlNumber
                                                            : r.AnimalControlNumber,
                                    HotWeight           = r.NewHotWeight,
                                    Grade               = r.NewGrade,
                                    HealthScore         = r.NewHealthScore,
                                    IsCondemned         = r.IsCondemned,
                                    Comment             = string.IsNullOrWhiteSpace(r.TrimComment) ? null : r.TrimComment
                                }).ToList();

                                billsUpdated = await _animalService.SaveKillDataAsync(killData);
                                _logger.LogInformation(
                                    "[HW-LOAD] Persisted HW data to {Count} bills (HotWeight, Grade, HS, ACN) — KillStatus unchanged.",
                                    billsUpdated);
                            }
                            catch (Exception saveEx)
                            {
                                _logger.LogError(saveEx, "[HW-LOAD] Failed to persist HW data to bills");
                                TempData["ErrorMessage"] = "Could not save Hot Weight data to bills: " + saveEx.Message;
                                return RedirectToAction(nameof(HotWeightImport));
                            }
                        }

                        // 4) Mark loaded rows in HW staging as "Loaded" (purple badge)
                        //    so they don't show as Ready next time.
                        int marked = 0;
                        foreach (var r in masterVm.AutoRows.Where(r => loadedIds.Contains(r.ControlNo) && r.Status == "OK"))
                        {
                            r.Status = "Loaded";
                            marked++;
                        }

                        // Remove loaded flagged rows from FlaggedRows so they don't reappear after staging restore
                        var loadedFlaggedKeys = masterVm.FlaggedRows
                            .Where(r => r.ControlNo > 0 && loadedIds.Contains(r.ControlNo))
                            .ToList();
                        foreach (var r in loadedFlaggedKeys)
                            masterVm.FlaggedRows.Remove(r);
                        marked += loadedFlaggedKeys.Count;

                        var updatedMaster = System.Text.Json.JsonSerializer.Serialize(masterVm);
                        await StagingBridge.WriteSharedAsync(
                        _stagingService,
                        StagingBridge.Types.HotWeight,
                        StagingBridge.SharedHotWeightKey,
                        updatedMaster,
                        sourceFileName: null);

                        // Session now holds the FULL master post-Load:
                        //  - all original FlaggedRows + DupRows preserved
                        //  - Loaded ControlNos have Status="Loaded"
                        // MarkKilledFast hwLookup excludes Loaded → button hides,
                        // pre-fill goes away, bills show their saved values directly.
                        TempData["HWPreview"] = updatedMaster;
                        HttpContext.Session.SetString("HWPreview", updatedMaster);

                        _logger.LogInformation("[HW-LOAD] merged={Merged}, markedLoaded={Marked}",
                            fixedRowsParsed.Count, marked);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "[HW-LOAD] Could not update master session");
                }
            }
            else
            {
                // No master available — fall back to the filtered subset for session
                // so Mark Killed at least has SOMETHING to pre-fill from. This branch
                // should be rare since we read master at the top of this action.
                TempData["HWPreview"] = json;
                HttpContext.Session.SetString("HWPreview", json);
            }

            _logger.LogInformation("[HW-LOAD] Saved HW data. Redirecting back to Hot Weight Report (Ready tab). Json length={Len}", json.Length);

            // Operator-visible summary message. Tells them what happened and what's next.
            int idsLoadedCount = 0;
            try
            {
                idsLoadedCount = selectedControlNos?.Split(',', StringSplitOptions.RemoveEmptyEntries).Count() ?? 0;
            }
            catch { }
            if (idsLoadedCount > 0)
            {
                TempData["SuccessMessage"] = $"Saved Hot Weight data for {idsLoadedCount} bills — visible below in 'Ready & killed today' with a 'Loaded ✓' badge. When ready to finalize, click 'Go to Mark as Killed'.";
            }
            return RedirectToAction(nameof(HotWeightImport), new { tab = 1 });
        }

        // HOT WEIGHT - CLEAR SESSION
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> HotWeightClearSession()
        {
            await StagingBridge.ClearSharedAsync(
                _stagingService,
                StagingBridge.Types.HotWeight,
                StagingBridge.SharedHotWeightKey);

            // Clear all session/TempData keys related to Hot Weight pre-fill
            // so a fresh refresh starts truly empty.
            HttpContext.Session.Remove("HWLoaded");
            HttpContext.Session.Remove("HWPreview");
            TempData.Remove("HWPreview");
            TempData.Remove("HWLoaded");
            TempData["SuccessMessage"] = "Hot Weight session cleared. Click Refresh from Hot Scale to start fresh.";
            return RedirectToAction(nameof(HotWeightImport));
        }

        // EXCEL PREVIEW - AJAX-
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ExcelFixErrorRow([FromBody] ExcelFixRowRequest req)
    {
        var json = HttpContext.Session.GetString("ExcelPreview")
                ?? TempData.Peek("ExcelPreview") as string;
        if (string.IsNullOrEmpty(json))
            return Json(new { success = false, message = "Session expired. Please re-upload the file." });

        try
        {
            var vm = System.Text.Json.JsonSerializer.Deserialize<ExcelImportViewModel>(json);
            if (vm == null) return Json(new { success = false, message = "Invalid session data." });

            var row = vm.Rows.FirstOrDefault(r => r.RowNum == req.RowNum);
            if (row == null) return Json(new { success = false, message = $"Row {req.RowNum} not found." });

            // Validate required fields
            var missing = new List<string>();
            if (string.IsNullOrWhiteSpace(req.VendorName))    missing.Add("Vendor");
            if (string.IsNullOrWhiteSpace(req.TagNumber1))    missing.Add("Tag Number One");
            if (string.IsNullOrWhiteSpace(req.PurchaseType))  missing.Add("Purchase Type");
            if (!req.PurchaseDate.HasValue)                    missing.Add("Purchase Date");

            if (missing.Any())
                return Json(new { success = false, message = "Missing required: " + string.Join(", ", missing) });

            // Apply all editable fields
            row.VendorName           = req.VendorName!.Trim();
            row.TagNumber1           = req.TagNumber1!.Trim();
            row.PurchaseType         = req.PurchaseType!.Trim();
            row.PurchaseDate         = req.PurchaseDate!.Value;
            row.TagNumber2           = string.IsNullOrWhiteSpace(req.TagNumber2)    ? null : req.TagNumber2.Trim();
            row.Tag3                 = string.IsNullOrWhiteSpace(req.Tag3)          ? null : req.Tag3.Trim();
            row.AnimalType           = string.IsNullOrWhiteSpace(req.AnimalType)    ? "Cow" : req.AnimalType.Trim();
            row.AnimalType2          = string.IsNullOrWhiteSpace(req.AnimalType2)   ? null : req.AnimalType2.Trim();
            row.LiveWeight           = req.LiveWeight;
            row.LiveRate             = req.LiveRate;
            row.HotWeight            = req.HotWeight;
            row.Grade                = string.IsNullOrWhiteSpace(req.Grade)         ? null : req.Grade.Trim();
            row.HealthScore          = req.HealthScore;
            row.Comment              = string.IsNullOrWhiteSpace(req.Comment)       ? null : req.Comment.Trim();
            row.AnimalControlNumber  = string.IsNullOrWhiteSpace(req.AnimalControlNumber) ? null : req.AnimalControlNumber.Trim();
            row.OfficeUse2           = string.IsNullOrWhiteSpace(req.OfficeUse2)    ? null : req.OfficeUse2.Trim();
            row.State                = string.IsNullOrWhiteSpace(req.State)         ? null : req.State.Trim();
            row.BuyerName            = string.IsNullOrWhiteSpace(req.BuyerName)     ? null : req.BuyerName.Trim();
            row.VetName              = string.IsNullOrWhiteSpace(req.VetName)       ? null : req.VetName.Trim();
            row.Status               = "OK";
            row.StatusNote           = "Fixed manually in preview";

            // Save updated session
            var updated = System.Text.Json.JsonSerializer.Serialize(vm);
            //HttpContext.Session.SetString("ExcelPreview", updated);
            TempData["ExcelPreview"] = updated;
            await StagingBridge.WriteAsync(
                HttpContext.Session, _stagingService,
                StagingBridge.Types.Excel,
                StagingBridge.GetUserKey(HttpContext),
                updated);

            int okCount  = vm.Rows.Count(r => r.Status == "OK");
            int errCount = vm.Rows.Count(r => r.Status == "Error");
            int dupCount = vm.Rows.Count(r => r.Status == "Duplicate");

            return Json(new { success = true, okCount, errCount, dupCount,
                message = $"Row {req.RowNum} fixed and added to import." });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = "Error: " + ex.Message });
        }
    }

    // EXCEL PREVIEW — AJAX: Delete a row from the session
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ExcelDeleteRow([FromBody] ExcelDeleteRowRequest req)
    {
        var json = HttpContext.Session.GetString("ExcelPreview")
                ?? TempData.Peek("ExcelPreview") as string;
        if (string.IsNullOrEmpty(json))
            return Json(new { success = false, message = "Session expired." });

        try
        {
            var vm = System.Text.Json.JsonSerializer.Deserialize<ExcelImportViewModel>(json);
            if (vm == null) return Json(new { success = false, message = "Invalid session." });

            var removed = vm.Rows.RemoveAll(r => r.RowNum == req.RowNum);
            if (removed == 0) return Json(new { success = false, message = $"Row {req.RowNum} not found." });

            var updated = System.Text.Json.JsonSerializer.Serialize(vm);
            //HttpContext.Session.SetString("ExcelPreview", updated);
            TempData["ExcelPreview"] = updated;
            //also persist updated preview to staging so the fix survives close
            await StagingBridge.WriteAsync(
                HttpContext.Session, _stagingService,
                StagingBridge.Types.Excel,
                StagingBridge.GetUserKey(HttpContext),
                updated);

            int okCount  = vm.Rows.Count(r => r.Status == "OK");
            int errCount = vm.Rows.Count(r => r.Status == "Error");
            int dupCount = vm.Rows.Count(r => r.Status == "Duplicate");

            return Json(new { success = true, okCount, errCount, dupCount,
                message = $"Row {req.RowNum} removed." });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = "Error: " + ex.Message });
        }
    }

    private static decimal GetDecimalCell(IXLCell cell)
        {
            if (cell == null) return 0;
            try
            {
                if (cell.DataType == XLDataType.Number) return (decimal)cell.GetDouble();
                var s = GetCellString(cell).Replace(",", "").Replace("$", "").Trim();
                return decimal.TryParse(s, out var d) ? d : 0;
            }
            catch { return 0; }
        }

        //  Helper 
        private static string? NullIfEmpty(string? s)
            => string.IsNullOrWhiteSpace(s) ? null : s;
        
        private static string? NormalizeAcn(string? acn)
        {
        if (string.IsNullOrWhiteSpace(acn)) return null;
        var v = acn.Trim();
        return v.All(ch => ch == '0') ? null : v;
        }

        private static bool IsAcnMissing(string? acn) => NormalizeAcn(acn) == null;

        private static readonly HashSet<string> PlaceholderTagWords =
            new(StringComparer.OrdinalIgnoreCase) { "ear", "rt", "slow", "nt", "none", "no tag", "lt", "tag" };

        private static bool IsPlaceholderTag(string? tag)
        {
            if (string.IsNullOrWhiteSpace(tag)) return false;
            return PlaceholderTagWords.Contains(tag.Trim());
        }

        private static string NormalizeTagToken(string? tag)
        {
            if (string.IsNullOrWhiteSpace(tag)) return "";

            var cleaned = new string(tag
                .Trim()
                .ToUpperInvariant()
                .Where(char.IsLetterOrDigit)
                .ToArray());

            if (cleaned.Length == 0) return "";

            if (cleaned.All(char.IsDigit))
            {
                var noLeadingZeros = cleaned.TrimStart('0');
                return noLeadingZeros.Length == 0 ? "0" : noLeadingZeros;
            }

            return cleaned;
        }

        private static string TrailingFourDigits(string? tag)
        {
            if(string.IsNullOrWhiteSpace(tag)) return "";
            var m = System.Text.RegularExpressions.Regex.Match(tag.Trim().ToUpperInvariant(), @"(\d{1,4})$");
            return m.Success ? m.Groups[1].Value : "";
        }

        // Returns true if either side has wildcard '?' chars and the two
        // tokens are the same length and match position-for-position
        // (with each '?' covering any single character on the other side).
        // Used to recognise file tags like '299?' as equivalent to bill
        // tag '2998' — same animal, just one digit was unreadable on the
        // floor when the operator scanned/typed.
        private static bool WildcardTagMatch(string? rawLeft, string? rawRight)
        {
            if (string.IsNullOrWhiteSpace(rawLeft) || string.IsNullOrWhiteSpace(rawRight))
                return false;

            // Normalise: trim, uppercase, strip non-alphanumeric EXCEPT '?'
            // (the wildcard char must survive the normaliser).
            string Norm(string s) => new string(
                s.Trim().ToUpperInvariant()
                 .Where(c => char.IsLetterOrDigit(c) || c == '?')
                 .ToArray());

            var l = Norm(rawLeft);
            var r = Norm(rawRight);

            if (l.Length == 0 || r.Length == 0) return false;
            if (l.Length != r.Length) return false;

            // Bail early if neither side has wildcards — the direct compare
            // path in TagEquivalent already covers exact equality.
            if (!l.Contains('?') && !r.Contains('?')) return false;

            for (int i = 0; i < l.Length; i++)
            {
                if (l[i] == '?' || r[i] == '?') continue;
                if (l[i] != r[i]) return false;
            }
            return true;
        }

        private static bool TagEquivalent(string? left, string? right)
        {
            if(IsPlaceholderTag(left) || IsPlaceholderTag(right)) return false;
            var l = NormalizeTagToken(left);
            var r = NormalizeTagToken(right);
            if (l.Length == 0 || r.Length == 0) return false;

            if (string.Equals(l, r, StringComparison.OrdinalIgnoreCase))
                return true;

            var lTail = TrailingFourDigits(left);
            var rTail = TrailingFourDigits(right);
            if (lTail.Length == 4 && rTail.Length == 4
                && string.Equals(lTail, rTail, StringComparison.OrdinalIgnoreCase))
                return true;

            if (WildcardTagMatch(left, right))
                return true;

            return false;
        }

        private static bool AnimalMatchesAnyFileTag(BarnData.Data.Entities.Animal a, string? backTag, string? tag1, string? tag2)
        {
            return TagEquivalent(a.TagNumber1, backTag) ||
                TagEquivalent(a.TagNumber2, backTag) ||
                TagEquivalent(a.Tag3, backTag) ||
                TagEquivalent(a.TagNumber1, tag1) ||
                TagEquivalent(a.TagNumber2, tag1) ||
                TagEquivalent(a.Tag3, tag1) ||
                TagEquivalent(a.TagNumber1, tag2) ||
                TagEquivalent(a.TagNumber2, tag2) ||
                TagEquivalent(a.Tag3, tag2);
        }

    private static bool IsBullLikeAnimal(string? animalType)
    {
        var value = animalType ?? "";
        return value.Contains("Bull", StringComparison.OrdinalIgnoreCase)
            || value.Contains("Steer", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsCowLikeAnimal(string? animalType)
    {
        var value = animalType ?? "";
        return value.Contains("Cow", StringComparison.OrdinalIgnoreCase)
            || value.Contains("Heifer", StringComparison.OrdinalIgnoreCase);
    }

    private static string? ValidateMarkKilledRow(AnimalRowDto row)
{
    var isConsignment = (row.PurchaseType ?? "").Contains("consignment", StringComparison.OrdinalIgnoreCase);
    var isHwImported = row.HwImported;
    if (row.LiveRate < 0)
        return $"Ctrl No {row.ControlNo}: Live Rate cannot be negative.";

    // Imported HW rows can save without Live Wt for consignment.
    // Live Wt will be backfilled from Hot Wt in the save mapping.
    if (!isHwImported && isConsignment && row.HotWeight > 0 && row.LiveWeight <= 0)
        return $"Ctrl No {row.ControlNo}: Live Wt is required for consignment when Hot Wt is entered.";

    if (row.HotWeight > 0 && row.LiveWeight > 0 && row.HotWeight > row.LiveWeight)
        return $"Ctrl No {row.ControlNo}: Hot Wt ({row.HotWeight:N1}) cannot exceed Live Wt ({row.LiveWeight:N1}).";

    var grade = (row.Grade ?? "").Trim().ToUpperInvariant();
    if (string.IsNullOrEmpty(grade))
        return null;

    
    if (GradeRules.IsCondemnationCode(grade))
        return null;

    
    if (isHwImported)
        return null;

    
    var bullGrades = GradeRules.BullGrades;
    var cowGrades  = GradeRules.CowGrades;

    if (IsBullLikeAnimal(row.AnimalType) && !bullGrades.Contains(grade))
        return $"Ctrl No {row.ControlNo}: Grade {grade} is not valid for {row.AnimalType}. Allowed: {string.Join(", ", bullGrades)}.";

    if (IsCowLikeAnimal(row.AnimalType) && !cowGrades.Contains(grade))
        return $"Ctrl No {row.ControlNo}: Grade {grade} is not valid for {row.AnimalType}. Allowed: {string.Join(", ", cowGrades)}.";

    return null;
}

   private static string? ValidateLegacyMarkKilledRow(int id, IFormCollection form)
{
    var purchaseType = form[$"purchaseType_{id}"].FirstOrDefault() ?? "";
    var animalType = form[$"animalType_{id}"].FirstOrDefault() ?? "";
    var grade = (form[$"grade_{id}"].FirstOrDefault() ?? "").Trim().ToUpperInvariant();
    var hwImported = form[$"hwImported_{id}"].FirstOrDefault() == "1";

    var isConsignment = purchaseType.Contains("consignment", StringComparison.OrdinalIgnoreCase);

    bool hasHot = decimal.TryParse(form[$"hotWeight_{id}"], out var hot) && hot > 0;
    bool hasLive = decimal.TryParse(form[$"liveWeight_{id}"], out var live) && live > 0;

    if (!hwImported && isConsignment && hasHot && !hasLive)
        return $"Ctrl No {id}: Live Wt is required for consignment when Hot Wt is entered.";

    if (hasHot && hasLive && hot > live)
        return $"Ctrl No {id}: Hot Wt ({hot:N1}) cannot exceed Live Wt ({live:N1}).";

    if (string.IsNullOrEmpty(grade))
        return null;

    // Condemnation grades (anything starting with X) bypass allow-list checks.
    if (GradeRules.IsCondemnationCode(grade))
        return null;

    if (hwImported)
        return null;

    // Use shared rules so this matches the Hot Weight pipeline exactly.
    var bullGrades = GradeRules.BullGrades;
    var cowGrades  = GradeRules.CowGrades;

    if (IsBullLikeAnimal(animalType) && !bullGrades.Contains(grade))
        return $"Ctrl No {id}: Grade {grade} is not valid for {animalType}. Allowed: {string.Join(", ", bullGrades)}.";

    if (IsCowLikeAnimal(animalType) && !cowGrades.Contains(grade))
        return $"Ctrl No {id}: Grade {grade} is not valid for {animalType}. Allowed: {string.Join(", ", cowGrades)}.";

    return null;
}
        private static string GetCellString(IXLCell cell)
        {
            if (cell == null) return "";
            try
            {
                // Numeric stored as number — convert to string without decimal
                if (cell.DataType == XLDataType.Number)
                {
                    var d = cell.GetDouble();
                    return d == Math.Floor(d)
                        ? ((long)d).ToString()
                        : d.ToString();
                }
                if (cell.DataType == XLDataType.Text)   return cell.GetString().Trim();
                if(cell.DataType == XLDataType.Boolean) return cell.GetBoolean().ToString();

                var value = cell.CachedValue.ToString()?.Trim();
                return string.IsNullOrEmpty(value) ? cell.GetString().Trim() : value;
            }
            catch { return cell.GetString().Trim(); }
        }

        private static DateTime? GetCellDate(IXLCell cell)
        {
            if (cell == null) return null;
            try
            {
                if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();

                var raw = cell.CachedValue.ToString();
                if (string.IsNullOrEmpty(raw)) raw = cell.GetString();
                if (DateTime.TryParse(raw, out var dt)) return dt;
            }
            catch { }
            return null;
        }
        //  EXCEL IMPORT — GET 
        public IActionResult Excel()
        {
            var sessionJson = HttpContext.Session.GetString("ExcelPreview");
            ExcelImportViewModel? vm = null;
            if (!string.IsNullOrEmpty(sessionJson))
            {
                try { vm = System.Text.Json.JsonSerializer.Deserialize<ExcelImportViewModel>(sessionJson); }
                catch { /* corrupt session — show empty tabs */ }
            }
            return View("ExcelPreview", vm ?? new ExcelImportViewModel());
        }
        public async Task<IActionResult> ExcelPreview()
        {
            var json = await StagingBridge.ReadAsync(
                HttpContext.Session, _stagingService,
                StagingBridge.Types.Excel,
                StagingBridge.GetUserKey(HttpContext));

            if (string.IsNullOrEmpty(json))
            {
                TempData["InfoMessage"] = "No Excel preview available. Upload a file to begin.";
                return RedirectToAction(nameof(Excel));
            }

            ExcelImportViewModel? vm = null;
            try { vm = System.Text.Json.JsonSerializer.Deserialize<ExcelImportViewModel>(json); }
            catch { vm = null; }

            if (vm == null || vm.Rows == null || !vm.Rows.Any())
            {
                TempData["InfoMessage"] = "Staged Excel preview was empty. Please re-upload.";
                return RedirectToAction(nameof(Excel));
            }

            
            ViewBag.RestoredFromStaging = true;
            return View("ExcelPreview", vm);
        }

         
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
            if (ext != ".xlsx" && ext != ".xls")
            {
                ModelState.AddModelError("", "Only .xlsx and .xls files are supported.");
                return View();
            }

            var vm = new ExcelImportViewModel { FileName = file.FileName };
            var vendors = (await _vendorService.GetAllActiveAsync()).ToList();
            var inFileDupes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var existingTagVendorKeys = await _animalService.GetAllTagVendorKeysAsync();

            const int maxVendorName = 150;
            const int maxTag = 50;
            const int maxAnimalType = 50;
            const int maxGrade = 20;
            const int maxComment = 500;
            const int maxControlNo = 50;
            const int maxOfficeUse2 = 100;
            const int maxState = 2;
            const int maxBuyer = 100;
            const int maxVet = 100;

            string NormalizeHeader(string? s)
                => (s ?? "").Trim().Replace(":", "").ToLowerInvariant();

            string NormalizeRequiredField(string? s)
                => (s ?? "").Trim();

            string? TrimWithNote(string? value, int maxLen, string fieldName, List<string> notes)
            {
                if (string.IsNullOrWhiteSpace(value)) return null;
                var trimmed = value.Trim();
                if (trimmed.Length <= maxLen) return trimmed;

                notes.Add($"{fieldName} trimmed to {maxLen} chars");
                return trimmed[..maxLen];
            }

            decimal ParseDecimal(string raw)
            {
                if (string.IsNullOrWhiteSpace(raw)) return 0;
                var cleaned = raw.Replace("$", "").Replace(",", "").Trim();
                if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out var invariant)) return invariant;
                if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.CurrentCulture, out var current)) return current;
                return 0;
            }

            async Task AddPreviewRowAsync(
                int rowNum,
                string vendorRaw,
                string tag1Raw,
                string purchaseTypeRaw,
                DateTime? purchaseDate,
                string animalTypeRaw,
                string? tag2Raw,
                string? tag3Raw,
                string? animalType2Raw,
                decimal liveWeight,
                decimal liveRate,
                decimal hotWeightRaw,
                string? gradeRaw,
                int? healthScore,
                string? commentRaw,
                string? animalControlNoRaw,
                string? officeUse2Raw,
                string? stateRaw,
                string? buyerRaw,
                string? vetRaw,
                bool isCondemned,
                List<string> notes)
            {
                var vendorName = NormalizeRequiredField(vendorRaw);
                var tag1 = NormalizeRequiredField(tag1Raw);
                var purchaseTypeValue = NormalizeRequiredField(purchaseTypeRaw);

                if (string.IsNullOrEmpty(tag1))
                {
                    if (string.IsNullOrEmpty(vendorName) && string.IsNullOrEmpty(purchaseTypeValue) && !purchaseDate.HasValue)
                    {
                        return;
                    }

                    vm.TotalRows++;
                    vm.Rows.Add(new ExcelPreviewRow
                    {
                        RowNum = rowNum,
                        VendorName = vendorName,
                        TagNumber1 = "",
                        Status = "Error",
                        StatusNote = "Missing required field: Tag Number One"
                    });
                    return;
                }

                vm.TotalRows++;

                var missing = new List<string>();
                if (string.IsNullOrEmpty(vendorName)) missing.Add("Vendor");
                if (string.IsNullOrEmpty(purchaseTypeValue)) missing.Add("Purchase Type");
                if (!purchaseDate.HasValue) missing.Add("Purchase Date");

                if (missing.Count > 0)
                {
                    vm.Rows.Add(new ExcelPreviewRow
                    {
                        RowNum = rowNum,
                        VendorName = vendorName,
                        TagNumber1 = tag1,
                        Status = "Error",
                        StatusNote = "Missing required field(s): " + string.Join(", ", missing)
                    });
                    return;
                }

                var localNotes = new List<string>(notes);

                var vendorSafe = TrimWithNote(vendorName, maxVendorName, "Vendor", localNotes) ?? vendorName;
                var tag1Safe = TrimWithNote(tag1, maxTag, "Tag Number One", localNotes) ?? tag1;
                var tag2Safe = TrimWithNote(tag2Raw, maxTag, "Tag Number Two", localNotes);
                var tag3Safe = TrimWithNote(tag3Raw, maxTag, "Tag 3", localNotes);
                var animalType = TrimWithNote(animalTypeRaw, maxAnimalType, "Animal Type", localNotes) ?? "Cow";
                var animalType2 = TrimWithNote(animalType2Raw, maxAnimalType, "Animal Type 2", localNotes);
                var grade = TrimWithNote(gradeRaw, maxGrade, "Grade", localNotes);
                var comment = TrimWithNote(commentRaw, maxComment, "Comment", localNotes);
                var animalCtrlNo = TrimWithNote(animalControlNoRaw, maxControlNo, "Animal Control Number", localNotes);
                var officeUse2 = TrimWithNote(officeUse2Raw, maxOfficeUse2, "Office Use 2", localNotes);
                var state = TrimWithNote(stateRaw, maxState, "State", localNotes);
                var buyer = TrimWithNote(buyerRaw, maxBuyer, "Buyer", localNotes);
                var vet = TrimWithNote(vetRaw, maxVet, "Vet Name", localNotes);

                if (string.IsNullOrWhiteSpace(animalType)) animalType = "Cow";
                if (animalType.StartsWith("Str", StringComparison.OrdinalIgnoreCase)) animalType = "Steer";

                var purchaseType = purchaseTypeValue.Contains("consignment", StringComparison.OrdinalIgnoreCase)
                    ? "Consignment Bill"
                    : "Sale Bill";
                var safePurchaseDate = purchaseDate ?? DateTime.Today;

                var inFileKey = $"{vendorSafe}|{tag1Safe}";
                if (!inFileDupes.Add(inFileKey))
                {
                    vm.Rows.Add(new ExcelPreviewRow
                    {
                        RowNum = rowNum,
                        VendorName = vendorSafe,
                        TagNumber1 = tag1Safe,
                        TagNumber2 = tag2Safe,
                        Tag3 = tag3Safe,
                        AnimalType = animalType,
                        AnimalType2 = animalType2,
                        PurchaseType = purchaseType,
                        PurchaseDate = safePurchaseDate,
                        LiveWeight = liveWeight,
                        LiveRate = liveRate,
                        KillDate = null,
                        HotWeight = hotWeightRaw > 0 ? hotWeightRaw : null,
                        Grade = grade,
                        HealthScore = healthScore,
                        Comment = comment,
                        AnimalControlNumber = animalCtrlNo,
                        OfficeUse2 = officeUse2,
                        State = state,
                        BuyerName = buyer,
                        VetName = vet,
                        IsCondemned = isCondemned,
                        Status = "Duplicate",
                        StatusNote = "Duplicate tag in uploaded file for this vendor"
                    });
                    return;
                }

                var vendor = vendors.FirstOrDefault(v =>
                    v.VendorName.Equals(vendorSafe, StringComparison.OrdinalIgnoreCase));

                string status = "OK";
                string? statusNote = localNotes.Any() ? string.Join("; ", localNotes) : null;

                if (vendor?.VendorID > 0)
                {
                    // in-memory lookup against preloaded set.
                    var exists = existingTagVendorKeys.Contains((tag1Safe.Trim(), vendor.VendorID));
                    if (exists)
                    {
                        status = "Duplicate";
                        statusNote = string.IsNullOrEmpty(statusNote)
                            ? "Tag already exists for this vendor"
                            : statusNote + "; Tag already exists for this vendor";
                    }
                }

                vm.Rows.Add(new ExcelPreviewRow
                {
                    RowNum = rowNum,
                    VendorName = vendorSafe,
                    TagNumber1 = tag1Safe,
                    TagNumber2 = tag2Safe,
                    Tag3 = tag3Safe,
                    AnimalType = animalType,
                    AnimalType2 = animalType2,
                    PurchaseType = purchaseType,
                    PurchaseDate = safePurchaseDate,
                    LiveWeight = liveWeight,
                    LiveRate = liveRate,
                    KillDate = null,
                    HotWeight = hotWeightRaw > 0 ? hotWeightRaw : null,
                    Grade = grade,
                    HealthScore = healthScore,
                    Comment = comment,
                    AnimalControlNumber = animalCtrlNo,
                    OfficeUse2 = officeUse2,
                    State = state,
                    BuyerName = buyer,
                    VetName = vet,
                    IsCondemned = isCondemned,
                    Status = status,
                    StatusNote = statusNote,
                });
            }

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                if (ext == ".xlsx")
                {
                    using var wb = new XLWorkbook(stream);
                    var ws = wb.Worksheets.First();

                    var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 30;
                    for (int c = 1; c <= lastCol; c++)
                    {
                        var h = NormalizeHeader(ws.Cell(1, c).GetString());
                        if (!string.IsNullOrEmpty(h)) colMap[h] = c;
                    }

                    int Col(params string[] names)
                    {
                        foreach (var n in names)
                            if (colMap.TryGetValue(n.ToLowerInvariant(), out int c)) return c;
                        return -1;
                    }

                    (string Value, bool FormulaFallback) ReadText(int row, int col)
                    {
                        if (col < 0) return ("", false);
                        var cell = ws.Cell(row, col);
                        if (cell == null) return ("", false);

                        try
                        {
                            if (cell.DataType == XLDataType.Number)
                            {
                                var d = cell.GetDouble();
                                var num = d == Math.Floor(d) ? ((long)d).ToString() : d.ToString(CultureInfo.InvariantCulture);
                                return (num, false);
                            }

                            if (cell.DataType == XLDataType.Boolean) return (cell.GetBoolean().ToString(), false);
                            if (cell.DataType == XLDataType.DateTime)
                                return (cell.GetDateTime().ToString("MM/dd/yyyy", CultureInfo.InvariantCulture), false);

                            var cached = cell.CachedValue.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(cached)) return (cached, false);

                            var display = cell.GetFormattedString().Trim();
                            bool usedFallback = cell.HasFormula && !string.IsNullOrEmpty(display);
                            return (display, usedFallback);
                        }
                        catch
                        {
                            return (cell.GetString().Trim(), false);
                        }
                    }

                    DateTime? ReadDate(int row, int col, out bool formulaFallback)
                    {
                        formulaFallback = false;
                        if (col < 0) return null;

                        var cell = ws.Cell(row, col);
                        if (cell == null) return null;

                        try
                        {
                            if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();

                            if (cell.DataType == XLDataType.Number)
                            {
                                var rawNum = cell.GetDouble();
                                if (rawNum > 0 && rawNum < 2958465) return DateTime.FromOADate(rawNum);
                            }

                            var cached = cell.CachedValue.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(cached) && DateTime.TryParse(cached, out var cachedDate))
                                return cachedDate;

                            var display = cell.GetFormattedString().Trim();
                            if (DateTime.TryParse(display, out var displayDate))
                            {
                                formulaFallback = cell.HasFormula && string.IsNullOrEmpty(cached);
                                return displayDate;
                            }
                        }
                        catch
                        {
                        }

                        return null;
                    }

                    int colAnimalType = Col("animal type");
                    int colTag1 = Col("tag number one", "tag one", "tag 1");
                    int colTag2 = Col("tag number two", "tag two", "tag 2");
                    int colTag3 = Col("tag 3", "tag3");
                    int colPurchDate = Col("purchase date");
                    int colPurchType = Col("purchase type");
                    int colVendor = Col("vendor");
                    int colLiveWeight = Col("live weight");
                    int colLiveRate = Col("live rate");
                    int colHotWeight = Col("hot weight");
                    int colGrade = Col("grade");
                    int colHS = Col("h s", "hs", "health score");
                    int colComment = Col("comments", "comment");
                    int colACN = Col("animal control number");
                    int colOfficeUse2 = Col("office use 2");
                    int colState = Col("state");
                    int colBuyer = Col("buyer");
                    int colAnimalType2 = Col("animal type 2");
                    int colVetName = Col("vet name");

                    int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
                    for (int row = 2; row <= lastRow; row++)
                    {
                        var notes = new List<string>();

                        var (tag1Raw, tag1Formula) = ReadText(row, colTag1);
                        if (tag1Formula) notes.Add("Tag Number One used formula fallback");

                        var (vendorRaw, vendorFormula) = ReadText(row, colVendor);
                        if (vendorFormula) notes.Add("Vendor used formula fallback");

                        var (purchaseTypeRaw, purchTypeFormula) = ReadText(row, colPurchType);
                        if (purchTypeFormula) notes.Add("Purchase Type used formula fallback");

                        var purchaseDate = ReadDate(row, colPurchDate, out var purchaseDateFormula);
                        if (purchaseDateFormula) notes.Add("Purchase Date used formula fallback");

                        var (liveWeightRaw, lwFormula) = ReadText(row, colLiveWeight);
                        if (lwFormula) notes.Add("Live Weight used formula fallback");

                        var (liveRateRaw, lrFormula) = ReadText(row, colLiveRate);
                        if (lrFormula) notes.Add("Live Rate used formula fallback");

                        var (hotWeightRawText, hwFormula) = ReadText(row, colHotWeight);
                        if (hwFormula) notes.Add("Hot Weight used formula fallback");

                        var (animalTypeRaw, _) = ReadText(row, colAnimalType);
                        var (tag2Raw, _) = ReadText(row, colTag2);
                        var (tag3Raw, _) = ReadText(row, colTag3);
                        var (animalType2Raw, _) = ReadText(row, colAnimalType2);
                        var (gradeRaw, _) = ReadText(row, colGrade);
                        var (hsRaw, hsFormula) = ReadText(row, colHS);
                        if (hsFormula) notes.Add("Health Score used formula fallback");

                        var (commentRaw, _) = ReadText(row, colComment);
                        var (acnRaw, _) = ReadText(row, colACN);
                        var (officeUse2Raw, _) = ReadText(row, colOfficeUse2);
                        var (stateRaw, _) = ReadText(row, colState);
                        var (buyerRaw, _) = ReadText(row, colBuyer);
                        var (vetRaw, _) = ReadText(row, colVetName);

                        decimal liveWeight = ParseDecimal(liveWeightRaw);
                        decimal liveRate = ParseDecimal(liveRateRaw);
                        decimal hotWeight = ParseDecimal(hotWeightRawText);

                        int? hs = null;
                        if (int.TryParse(hsRaw, out var hsValue) && hsValue > 0) hs = hsValue;

                        var commentClean = NullIfEmpty(commentRaw);
                        bool isCond = !string.IsNullOrEmpty(commentClean)
                            && commentClean.Contains("cond", StringComparison.OrdinalIgnoreCase);

                        await AddPreviewRowAsync(
                            row,
                            vendorRaw,
                            tag1Raw,
                            purchaseTypeRaw,
                            purchaseDate,
                            animalTypeRaw,
                            NullIfEmpty(tag2Raw),
                            NullIfEmpty(tag3Raw),
                            NullIfEmpty(animalType2Raw),
                            liveWeight,
                            liveRate,
                            hotWeight,
                            NullIfEmpty(gradeRaw),
                            hs,
                            commentClean,
                            NullIfEmpty(acnRaw),
                            NullIfEmpty(officeUse2Raw),
                            NullIfEmpty(stateRaw),
                            NullIfEmpty(buyerRaw),
                            NullIfEmpty(vetRaw),
                            isCond,
                            notes);
                    }
                }
                else
                {
                    using var workbook = new HSSFWorkbook(stream);
                    var sheet = workbook.GetSheetAt(0);
                    var formatter = new DataFormatter(CultureInfo.InvariantCulture);

                    var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    var headerRow = sheet.GetRow(0);
                    int lastCol = headerRow?.LastCellNum ?? 30;
                    for (int c = 0; c < lastCol; c++)
                    {
                        var header = NormalizeHeader(formatter.FormatCellValue(headerRow?.GetCell(c)));
                        if (!string.IsNullOrEmpty(header)) colMap[header] = c;
                    }

                    int Col(params string[] names)
                    {
                        foreach (var n in names)
                            if (colMap.TryGetValue(n.ToLowerInvariant(), out var c)) return c;
                        return -1;
                    }

                    ICell? CellAt(int row, int col)
                    {
                        if (col < 0) return null;
                        var r = sheet.GetRow(row);
                        return r?.GetCell(col, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    }

                    (string Value, bool FormulaFallback) ReadText(int row, int col)
                    {
                        var cell = CellAt(row, col);
                        if (cell == null) return ("", false);

                        try
                        {
                            switch (cell.CellType)
                            {
                                case CellType.Numeric:
                                    if (DateUtil.IsCellDateFormatted(cell))
                                        return ($"{cell.DateCellValue:MM/dd/yyyy}", false);
                                    var n = cell.NumericCellValue;
                                    return (Math.Abs(n % 1) < 0.0000001
                                        ? ((long)n).ToString()
                                        : n.ToString(CultureInfo.InvariantCulture), false);

                                case CellType.Boolean:
                                    return (cell.BooleanCellValue.ToString(), false);

                                case CellType.String:
                                    return ((cell.StringCellValue ?? "").Trim(), false);

                                case CellType.Formula:
                                    switch (cell.CachedFormulaResultType)
                                    {
                                        case CellType.Numeric:
                                            if (DateUtil.IsCellDateFormatted(cell))
                                                return ($"{cell.DateCellValue:MM/dd/yyyy}", false);
                                            var cachedNum = cell.NumericCellValue;
                                            return (Math.Abs(cachedNum % 1) < 0.0000001
                                                ? ((long)cachedNum).ToString()
                                                : cachedNum.ToString(CultureInfo.InvariantCulture), false);

                                        case CellType.String:
                                            return ((cell.StringCellValue ?? "").Trim(), false);

                                        case CellType.Boolean:
                                            return (cell.BooleanCellValue.ToString(), false);
                                    }

                                    var display = formatter.FormatCellValue(cell)?.Trim() ?? "";
                                    return (display, !string.IsNullOrEmpty(display));
                            }

                            return ((formatter.FormatCellValue(cell) ?? "").Trim(), false);
                        }
                        catch
                        {
                            return ((formatter.FormatCellValue(cell) ?? "").Trim(), false);
                        }
                    }

                    DateTime? ReadDate(int row, int col, out bool formulaFallback)
                    {
                        formulaFallback = false;
                        var cell = CellAt(row, col);
                        if (cell == null) return null;

                        try
                        {
                            if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
                                return cell.DateCellValue;

                            if (cell.CellType == CellType.Formula
                                && cell.CachedFormulaResultType == CellType.Numeric
                                && DateUtil.IsCellDateFormatted(cell))
                                return cell.DateCellValue;

                            var raw = (formatter.FormatCellValue(cell) ?? "").Trim();
                            if (DateTime.TryParse(raw, out var parsed))
                            {
                                formulaFallback = cell.CellType == CellType.Formula;
                                return parsed;
                            }
                        }
                        catch
                        {
                        }

                        return null;
                    }

                    int colAnimalType = Col("animal type");
                    int colTag1 = Col("tag number one", "tag one", "tag 1");
                    int colTag2 = Col("tag number two", "tag two", "tag 2");
                    int colTag3 = Col("tag 3", "tag3");
                    int colPurchDate = Col("purchase date");
                    int colPurchType = Col("purchase type");
                    int colVendor = Col("vendor");
                    int colLiveWeight = Col("live weight");
                    int colLiveRate = Col("live rate");
                    int colHotWeight = Col("hot weight");
                    int colGrade = Col("grade");
                    int colHS = Col("h s", "hs", "health score");
                    int colComment = Col("comments", "comment");
                    int colACN = Col("animal control number");
                    int colOfficeUse2 = Col("office use 2");
                    int colState = Col("state");
                    int colBuyer = Col("buyer");
                    int colAnimalType2 = Col("animal type 2");
                    int colVetName = Col("vet name");

                    int lastRow = sheet.LastRowNum;
                    for (int row = 1; row <= lastRow; row++)
                    {
                        var notes = new List<string>();

                        var (tag1Raw, tag1Formula) = ReadText(row, colTag1);
                        if (tag1Formula) notes.Add("Tag Number One used formula fallback");

                        var (vendorRaw, vendorFormula) = ReadText(row, colVendor);
                        if (vendorFormula) notes.Add("Vendor used formula fallback");

                        var (purchaseTypeRaw, purchTypeFormula) = ReadText(row, colPurchType);
                        if (purchTypeFormula) notes.Add("Purchase Type used formula fallback");

                        var purchaseDate = ReadDate(row, colPurchDate, out var purchaseDateFormula);
                        if (purchaseDateFormula) notes.Add("Purchase Date used formula fallback");

                        var (liveWeightRaw, lwFormula) = ReadText(row, colLiveWeight);
                        if (lwFormula) notes.Add("Live Weight used formula fallback");

                        var (liveRateRaw, lrFormula) = ReadText(row, colLiveRate);
                        if (lrFormula) notes.Add("Live Rate used formula fallback");

                        var (hotWeightRawText, hwFormula) = ReadText(row, colHotWeight);
                        if (hwFormula) notes.Add("Hot Weight used formula fallback");

                        var (animalTypeRaw, _) = ReadText(row, colAnimalType);
                        var (tag2Raw, _) = ReadText(row, colTag2);
                        var (tag3Raw, _) = ReadText(row, colTag3);
                        var (animalType2Raw, _) = ReadText(row, colAnimalType2);
                        var (gradeRaw, _) = ReadText(row, colGrade);
                        var (hsRaw, hsFormula) = ReadText(row, colHS);
                        if (hsFormula) notes.Add("Health Score used formula fallback");

                        var (commentRaw, _) = ReadText(row, colComment);
                        var (acnRaw, _) = ReadText(row, colACN);
                        var (officeUse2Raw, _) = ReadText(row, colOfficeUse2);
                        var (stateRaw, _) = ReadText(row, colState);
                        var (buyerRaw, _) = ReadText(row, colBuyer);
                        var (vetRaw, _) = ReadText(row, colVetName);

                        decimal liveWeight = ParseDecimal(liveWeightRaw);
                        decimal liveRate = ParseDecimal(liveRateRaw);
                        decimal hotWeight = ParseDecimal(hotWeightRawText);

                        int? hs = null;
                        if (int.TryParse(hsRaw, out var hsValue) && hsValue > 0) hs = hsValue;

                        var commentClean = NullIfEmpty(commentRaw);
                        bool isCond = !string.IsNullOrEmpty(commentClean)
                            && commentClean.Contains("cond", StringComparison.OrdinalIgnoreCase);

                        await AddPreviewRowAsync(
                            row + 1,
                            vendorRaw,
                            tag1Raw,
                            purchaseTypeRaw,
                            purchaseDate,
                            animalTypeRaw,
                            NullIfEmpty(tag2Raw),
                            NullIfEmpty(tag3Raw),
                            NullIfEmpty(animalType2Raw),
                            liveWeight,
                            liveRate,
                            hotWeight,
                            NullIfEmpty(gradeRaw),
                            hs,
                            commentClean,
                            NullIfEmpty(acnRaw),
                            NullIfEmpty(officeUse2Raw),
                            NullIfEmpty(stateRaw),
                            NullIfEmpty(buyerRaw),
                            NullIfEmpty(vetRaw),
                            isCond,
                            notes);
                    }
                }
            }
            catch (Exception ex)
            {
                vm.Errors.Add($"Parse error: {ex.Message}");
                return View(vm);
            }

            // Store preview in TempData (for confirm) and Session (for AJAX fix/delete)
            var excelJson = System.Text.Json.JsonSerializer.Serialize(vm);
            TempData["ExcelPreview"] = excelJson;
            await StagingBridge.WriteAsync(
                HttpContext.Session, _stagingService,
                StagingBridge.Types.Excel,
                StagingBridge.GetUserKey(HttpContext),
                excelJson,
                sourceFileName: file.FileName);
            return View("ExcelPreview", vm);
        }

        // EXCEL IMPORT - CONFIRM & SAVE 
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ExcelConfirm()
        {
            var json = TempData["ExcelPreview"] as string
                    ?? HttpContext.Session.GetString("ExcelPreview");
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
                    KillDate            = null,
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
                    KillStatus          = "Pending",
                    SaleBillRef         = billRef,
                });
            }

            var (imported, skipped, errors) = await _animalService.BulkImportAsync(toImport);
            int dupeCount = vm.Rows.Count(r => r.Status == "Duplicate");

            TempData["SuccessMessage"] =
                $"Import complete: {imported} animals saved, {skipped + dupeCount} skipped.";
            
            //mark staging batch as cleared so it doesn't restore next time
            await StagingBridge.ClearAsync(
                HttpContext.Session, _stagingService,
                StagingBridge.Types.Excel,
                StagingBridge.GetUserKey(HttpContext));

            return RedirectToAction("Index", "Animal", new { status = "pending" });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ExcelClearSession()
        {
            await StagingBridge.ClearAsync(
                HttpContext.Session, _stagingService,
                StagingBridge.Types.Excel,
                StagingBridge.GetUserKey(HttpContext));

            TempData.Remove("ExcelPreview");
            TempData["SuccessMessage"] = "Excel import session cleared.";
            return RedirectToAction(nameof(Excel));
        }

        // =============================================================================
        // HOT SCALE AUTO-PULL HELPERS
        // =============================================================================

        // ---------------------------------------------------------------------
        // ReadHotWeightRowsFromHotScaleDbAsync
        //
        // Runs the production Hot Weight report SQL against the Hot Scale DB
        // (configured via the "HotScale" connection string in appsettings) for
        // a single date.  Returns the rows as a list of HotScaleRow records,
        // ready to be projected into a synthetic workbook.
        //
        // Same query the SSRS "HotScaleReport" runs in production. Both
        // @reportdatestart and @reportdateend bind to the same date so the
        // shape matches the user-picked-date-only behaviour of the SSRS report.
        //
        // Throws on connection or execution failure - the caller is expected
        // to catch and present a friendly inline warning to the user.
        // ---------------------------------------------------------------------
        private async Task<List<HotScaleRow>> ReadHotWeightRowsFromHotScaleDbAsync(DateTime forDate)
        {
            var connStr = _configuration.GetConnectionString("HotScale");
            if (string.IsNullOrWhiteSpace(connStr))
                throw new InvalidOperationException("HotScale connection string is not configured.");

            const string sql = @"
SELECT
    CAST(tbl_animal_master.LiveWeightDate AS time(0)) AS timekilled,
    tbl_animal_master.AnimalNumber,
    tbl_animal_master.SexCode,
    tbl_animal_master.AnimalType,
    tbl_animal_master.BackTag,
    lot.KILL_LOT_NUMBER,
    tbl_animal_master.Tag1,
    tbl_animal_master.Tag2,
    tbl_animal_master.ProgramCode,
    tbl_animal_master.CountryCode,
    tbl_animal_master.HotWeightSide1,
    tbl_animal_master.HotWeightSide2,
    CAST(tbl_animal_master.LiveWeightDate AS date) AS datekilled,
    tbl_animal_master.HotWeightSide1 + tbl_animal_master.HotWeightSide2 AS hottotal,
    NULLIF(tbl_animal_master.LiveWeight, 0.00) AS liveweight,
    (tbl_animal_master.HotWeightSide1 + tbl_animal_master.HotWeightSide2) / tbl_animal_master.LiveWeight * 100 AS yieldpercent,
    side1.QUALITY_GRADE_CODE AS side1grade,
    side2.QUALITY_GRADE_CODE AS side2grade,
    tbl_animal_master.HealthScore,
    CASE WHEN side1.TRIM_GRADE_FRONT = 1 THEN ' ' WHEN side1.TRIM_GRADE_FRONT = 2 THEN 'LT' WHEN side1.TRIM_GRADE_FRONT = 3 THEN 'HT' END AS side1front,
    CASE WHEN side1.TRIM_GRADE_HIND  = 1 THEN ' ' WHEN side1.TRIM_GRADE_HIND  = 2 THEN 'LT' WHEN side1.TRIM_GRADE_HIND  = 3 THEN 'HT' END AS side1hind,
    CASE WHEN side2.TRIM_GRADE_FRONT = 1 THEN ' ' WHEN side2.TRIM_GRADE_FRONT = 2 THEN 'LT' WHEN side2.TRIM_GRADE_FRONT = 3 THEN 'HT' END AS side2front,
    CASE WHEN side2.TRIM_GRADE_HIND  = 1 THEN ' ' WHEN side2.TRIM_GRADE_HIND  = 2 THEN 'LT' WHEN side2.TRIM_GRADE_HIND  = 3 THEN 'HT' END AS side2hind
FROM tbl_animal_master
LEFT OUTER JOIN tbl_hot_weights_history AS side1 ON tbl_animal_master.BarCodeSide1 = side1.ANIMAL_BARCODE
LEFT OUTER JOIN tbl_hot_weights_history AS side2 ON tbl_animal_master.BarCodeSide2 = side2.ANIMAL_BARCODE
INNER JOIN tbl_animal_kill_lots AS lot ON tbl_animal_master.AnimalNumber = lot.ANIMAL_NUMBER
WHERE (CAST(tbl_animal_master.DateKilled AS date) BETWEEN @reportdatestart AND @reportdateend)
  AND (
        -- Production-killed: carcass weighed (both sides), graded, scored.
        -- Picks up rows where Hot Scale has finished its measurement workflow.
        (
            tbl_animal_master.HotWeightSide1 > 0
            AND tbl_animal_master.HotWeightSide2 > 0
            AND ISNULL(side1.QUALITY_GRADE_CODE, '') <> ''
            AND ISNULL(tbl_animal_master.HealthScore, 0) > 0
        )
        OR
        -- Condemned: no carcass to weigh, condemnation marker on either side.
        -- Either side1 or side2 grade starts with 'X' (XML, XO, XTOX, etc.).
        (
            ISNULL(tbl_animal_master.HotWeightSide1, 0) = 0
            AND ISNULL(tbl_animal_master.HotWeightSide2, 0) = 0
            AND (
                side1.QUALITY_GRADE_CODE LIKE 'X%'
                OR side2.QUALITY_GRADE_CODE LIKE 'X%'
            )
        )
      )
ORDER BY tbl_animal_master.AnimalNumber";

            var rows = new List<HotScaleRow>();

            await using var conn = new SqlConnection(connStr);
            await conn.OpenAsync();
            await using var cmd = new SqlCommand(sql, conn) { CommandTimeout = 60 };
            cmd.Parameters.Add(new SqlParameter("@reportdatestart", SqlDbType.Date) { Value = forDate.Date });
            cmd.Parameters.Add(new SqlParameter("@reportdateend",   SqlDbType.Date) { Value = forDate.Date });

            await using var reader = await cmd.ExecuteReaderAsync();

            // Resolve column ordinals once
            int oTimeKilled  = SafeOrdinal(reader, "timekilled");
            int oAnimalNo    = SafeOrdinal(reader, "AnimalNumber");
            int oSexCode     = SafeOrdinal(reader, "SexCode");
            int oAnimalType  = SafeOrdinal(reader, "AnimalType");
            int oBackTag     = SafeOrdinal(reader, "BackTag");
            int oLotNumber   = SafeOrdinal(reader, "KILL_LOT_NUMBER");
            int oTag1        = SafeOrdinal(reader, "Tag1");
            int oTag2        = SafeOrdinal(reader, "Tag2");
            int oProgramCode = SafeOrdinal(reader, "ProgramCode");
            int oCountryCode = SafeOrdinal(reader, "CountryCode");
            int oSide1       = SafeOrdinal(reader, "HotWeightSide1");
            int oSide2       = SafeOrdinal(reader, "HotWeightSide2");
            int oLiveWeight  = SafeOrdinal(reader, "liveweight");
            int oSide1Grade  = SafeOrdinal(reader, "side1grade");
            int oSide2Grade  = SafeOrdinal(reader, "side2grade");
            int oHealthScore = SafeOrdinal(reader, "HealthScore");

            while (await reader.ReadAsync())
            {
                rows.Add(new HotScaleRow
                {
                    TimeKilled    = oTimeKilled  >= 0 && !reader.IsDBNull(oTimeKilled)  ? reader.GetTimeSpan(oTimeKilled).ToString(@"hh\:mm\:ss") : "",
                    AnimalNumber  = oAnimalNo    >= 0 && !reader.IsDBNull(oAnimalNo)    ? reader.GetValue(oAnimalNo)?.ToString() ?? "" : "",
                    SexCode       = oSexCode     >= 0 && !reader.IsDBNull(oSexCode)     ? reader.GetString(oSexCode) : "",
                    AnimalType    = oAnimalType  >= 0 && !reader.IsDBNull(oAnimalType)  ? reader.GetString(oAnimalType) : "",
                    BackTag       = oBackTag     >= 0 && !reader.IsDBNull(oBackTag)     ? reader.GetValue(oBackTag)?.ToString() ?? "" : "",
                    KillLotNumber = oLotNumber   >= 0 && !reader.IsDBNull(oLotNumber)   ? reader.GetValue(oLotNumber)?.ToString() ?? "" : "",
                    Tag1          = oTag1        >= 0 && !reader.IsDBNull(oTag1)        ? reader.GetValue(oTag1)?.ToString() ?? "" : "",
                    Tag2          = oTag2        >= 0 && !reader.IsDBNull(oTag2)        ? reader.GetValue(oTag2)?.ToString() ?? "" : "",
                    ProgramCode   = oProgramCode >= 0 && !reader.IsDBNull(oProgramCode) ? reader.GetString(oProgramCode) : "",
                    CountryCode   = oCountryCode >= 0 && !reader.IsDBNull(oCountryCode) ? reader.GetString(oCountryCode) : "",
                    Side1         = oSide1       >= 0 && !reader.IsDBNull(oSide1)       ? Convert.ToDecimal(reader.GetValue(oSide1)) : (decimal?)null,
                    Side2         = oSide2       >= 0 && !reader.IsDBNull(oSide2)       ? Convert.ToDecimal(reader.GetValue(oSide2)) : (decimal?)null,
                    LiveWeight    = oLiveWeight  >= 0 && !reader.IsDBNull(oLiveWeight)  ? Convert.ToDecimal(reader.GetValue(oLiveWeight)) : (decimal?)null,
                    Side1Grade    = oSide1Grade  >= 0 && !reader.IsDBNull(oSide1Grade)  ? reader.GetString(oSide1Grade)?.Trim() ?? "" : "",
                    Side2Grade    = oSide2Grade  >= 0 && !reader.IsDBNull(oSide2Grade)  ? reader.GetString(oSide2Grade)?.Trim() ?? "" : "",
                    HealthScore   = oHealthScore >= 0 && !reader.IsDBNull(oHealthScore) ? Convert.ToInt32(reader.GetValue(oHealthScore)) : (int?)null,
                });
            }

            return rows;
        }

        // Helper: returns column ordinal, or -1 if not present (defensive — query
        // shape changes shouldn't crash the import).
        private static int SafeOrdinal(System.Data.Common.DbDataReader r, string colName)
        {
            try { return r.GetOrdinal(colName); }
            catch { return -1; }
        }

        // ---------------------------------------------------------------------
        // BuildHotScaleSyntheticFile
        //
        // Converts a list of HotScaleRow into an in-memory .xlsx file (wrapped
        // in IFormFile) whose column headers exactly match the names the
        // existing Hot Weight Excel parser already recognises.
        //
        // This is a deliberate adapter: by emitting the same workbook shape an
        // SSRS export would produce, we reuse the entire 1300-line matching
        // pipeline without forking it. If the parser's column aliases ever
        // change, only the header row below needs to follow.
        // ---------------------------------------------------------------------
        private static IFormFile BuildHotScaleSyntheticFile(List<HotScaleRow> rows, DateTime forDate)
        {
            var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("HotScale");

                // Header row — names match the parser's Col(...) aliases (line ~1187).
                int c = 1;
                ws.Cell(1, c++).Value = "Number";       // → ACN
                ws.Cell(1, c++).Value = "BackTag";
                ws.Cell(1, c++).Value = "Tag1";
                ws.Cell(1, c++).Value = "Tag2";
                ws.Cell(1, c++).Value = "Side 1 Hot";
                ws.Cell(1, c++).Value = "Side 2 Hot";
                ws.Cell(1, c++).Value = "Grade";        // ← side1grade
                ws.Cell(1, c++).Value = "Grade 2";      // ← side2grade
                ws.Cell(1, c++).Value = "HealthScore";
                ws.Cell(1, c++).Value = "LiveWeight";
                ws.Cell(1, c++).Value = "Origin";       // ← CountryCode
                ws.Cell(1, c++).Value = "Lot";          // ← KILL_LOT_NUMBER
                ws.Cell(1, c++).Value = "Sex";          // ← SexCode
                ws.Cell(1, c++).Value = "Type";         // ← AnimalType
                ws.Cell(1, c++).Value = "Program";      // ← ProgramCode

                int r = 2;
                foreach (var row in rows)
                {
                    int cc = 1;
                    ws.Cell(r, cc++).Value = row.AnimalNumber  ?? "";
                    ws.Cell(r, cc++).Value = row.BackTag       ?? "";
                    ws.Cell(r, cc++).Value = row.Tag1          ?? "";
                    ws.Cell(r, cc++).Value = row.Tag2          ?? "";

                    // Numeric cells: only set when a value is present so empty
                    // cells (left as blank) are read by the parser as "no value"
                    // — matches the (v > 0 ? v : null) check in the existing
                    // GetDecimalCell helper.
                    if (row.Side1.HasValue)       ws.Cell(r, cc).Value = row.Side1.Value;
                    cc++;
                    if (row.Side2.HasValue)       ws.Cell(r, cc).Value = row.Side2.Value;
                    cc++;

                    ws.Cell(r, cc++).Value = row.Side1Grade    ?? "";
                    ws.Cell(r, cc++).Value = row.Side2Grade    ?? "";

                    if (row.HealthScore.HasValue) ws.Cell(r, cc).Value = row.HealthScore.Value;
                    cc++;
                    if (row.LiveWeight.HasValue)  ws.Cell(r, cc).Value = row.LiveWeight.Value;
                    cc++;

                    ws.Cell(r, cc++).Value = row.CountryCode   ?? "";
                    ws.Cell(r, cc++).Value = row.KillLotNumber ?? "";
                    ws.Cell(r, cc++).Value = row.SexCode       ?? "";
                    ws.Cell(r, cc++).Value = row.AnimalType    ?? "";
                    ws.Cell(r, cc++).Value = row.ProgramCode   ?? "";
                    r++;
                }

                wb.SaveAs(ms);
            }
            ms.Position = 0;

            return new FormFile(ms, 0, ms.Length, "file",
                $"HotScale-{forDate:yyyy-MM-dd}.xlsx")
            {
                Headers      = new Microsoft.AspNetCore.Http.HeaderDictionary(),
                ContentType  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            };
        }

        // ---------------------------------------------------------------------
        // MergeHotWeightStaging
        //
        // Smart-merge: takes a fresh VM (from auto-pull or upload) and an
        // existing VM (from shared staging), and produces a merged VM that
        // preserves any rows the team has already manually fixed.
        //
        // Match strategy: by (positive) ControlNo first, then by a stable
        // identifier built from the tag columns.  When an existing fixed row
        // is found, the user's edits override the fresh pipeline output.
        //
        // Loaded rows are preserved across merges, BUT only if the matching bill
        // in BarnData is still in KillStatus=Killed. If a bill that was previously
        // loaded has been deleted or reset to Pending (e.g. operator cleared bills
        // and re-imported), we drop the stale Loaded marker — otherwise the new
        // batch ghost-flags those rows as "already killed" when they aren't.
        // ---------------------------------------------------------------------
        private HotWeightImportViewModel MergeHotWeightStaging(
            HotWeightImportViewModel fresh,
            HotWeightImportViewModel? existing,
            HashSet<int>? currentlyKilledControlNos = null)
        {
            if (existing == null) return fresh;

            // Index existing rows for lookup
            var existingByCtrl = new Dictionary<int, HotWeightPreviewRow>();
            var existingByTag  = new Dictionary<string, HotWeightPreviewRow>(StringComparer.OrdinalIgnoreCase);

            foreach (var er in existing.AutoRows.Concat(existing.FlaggedRows))
            {
                if (er.ControlNo > 0 && !existingByCtrl.ContainsKey(er.ControlNo))
                    existingByCtrl[er.ControlNo] = er;

                var k = TagIdentityKey(er);
                if (!string.IsNullOrEmpty(k) && !existingByTag.ContainsKey(k))
                    existingByTag[k] = er;
            }

            int preserved = 0;
            int staleLoadedDropped = 0;

            void TryPreserve(HotWeightPreviewRow target)
            {
                HotWeightPreviewRow? existingRow = null;
                if (target.ControlNo > 0 && existingByCtrl.TryGetValue(target.ControlNo, out var byCtrl))
                    existingRow = byCtrl;
                if (existingRow == null)
                {
                    var k = TagIdentityKey(target);
                    if (!string.IsNullOrEmpty(k) && existingByTag.TryGetValue(k, out var byTag))
                        existingRow = byTag;
                }

                if (existingRow == null) return;

                // Loaded rows carry forward their Loaded status only if the
                // matched bill is still actually killed in BarnData. Otherwise
                // (bill deleted, bill reset to Pending, etc.) we drop the stale
                // marker and let it flow through as a fresh row.
                if (string.Equals(existingRow.Status, "Loaded", StringComparison.OrdinalIgnoreCase))
                {
                    bool billIsStillKilled = currentlyKilledControlNos != null
                        && target.ControlNo > 0
                        && currentlyKilledControlNos.Contains(target.ControlNo);

                    // If we have no killed-set context (e.g. callers that haven't
                    // populated it), be conservative and preserve. If we DO have
                    // the set and the bill isn't in it, drop the stale marker.
                    if (currentlyKilledControlNos == null || billIsStillKilled)
                    {
                        target.Status            = "Loaded";
                        target.IsManuallyEdited  = existingRow.IsManuallyEdited;
                        preserved++;
                        return;
                    }
                    else
                    {
                        // Stale Loaded — drop it. Don't return — fall through to the
                        // manual-edit-preservation path below in case there are still
                        // operator fixes worth keeping.
                        staleLoadedDropped++;
                    }
                }

                // Otherwise, only preserve if the user manually edited the row.
                if (!existingRow.IsManuallyEdited) return;

                target.ControlNo               = existingRow.ControlNo;
                if (!string.IsNullOrWhiteSpace(existingRow.AnimalControlNumber))
                    target.AnimalControlNumber = existingRow.AnimalControlNumber;
                target.NewHotWeight            = existingRow.NewHotWeight ?? target.NewHotWeight;
                target.NewGrade                = existingRow.NewGrade  ?? target.NewGrade;
                target.NewGrade2               = existingRow.NewGrade2 ?? target.NewGrade2;
                target.NewHealthScore          = existingRow.NewHealthScore ?? target.NewHealthScore;
                target.NewAnimalControlNumber  = existingRow.NewAnimalControlNumber ?? target.NewAnimalControlNumber;
                target.Status                  = existingRow.Status;
                target.FlagReason              = existingRow.FlagReason;
                target.IsManuallyEdited        = true;
                preserved++;
            }

            foreach (var row in fresh.AutoRows)    TryPreserve(row);
            foreach (var row in fresh.FlaggedRows) TryPreserve(row);

            // After preserving fixes, some rows that were originally Flagged in
            // 'fresh' may now be marked Status="OK" thanks to a manual fix that
            // moved them out of the flag pile. Reorganize to reflect that.
            var promoted = fresh.FlaggedRows.Where(r => r.Status == "OK").ToList();
            foreach (var p in promoted)
            {
                fresh.FlaggedRows.Remove(p);
                if (!fresh.AutoRows.Any(a => a.RowKey == p.RowKey || (p.ControlNo > 0 && a.ControlNo == p.ControlNo)))
                    fresh.AutoRows.Add(p);
            }

            // Final completeness re-check. A row reaches AutoRows/Ready only if it
            // would actually pass save-time validation. Without this, manually-edited
            // rows preserved across merges (or picker-only fixes that didn't supply
            // Side/Grade/HS values) could land in Ready with empty data — operator
            // clicks Load and the save endpoint reports "0 rows updated".
            var demoteFromAuto = fresh.AutoRows.Where(r =>
            {
                if (string.Equals(r.Status, "Loaded", StringComparison.OrdinalIgnoreCase))
                    return false; // already saved earlier — leave alone
                if (r.IsCondemned) return false; // condemned bypasses HW/Grade/HS
                if (string.IsNullOrWhiteSpace(r.AnimalControlNumber)) return true;
                if (!r.NewHotWeight.HasValue || r.NewHotWeight.Value <= 0) return true;
                if (string.IsNullOrWhiteSpace(r.NewGrade)) return true;
                if (!r.NewHealthScore.HasValue
                    || r.NewHealthScore.Value < 1
                    || r.NewHealthScore.Value > 5) return true;
                return false;
            }).ToList();

            foreach (var d in demoteFromAuto)
            {
                fresh.AutoRows.Remove(d);
                d.Status = "Flag";
                var reasons = new List<string>();
                if (string.IsNullOrWhiteSpace(d.AnimalControlNumber)) reasons.Add("ACN missing");
                if (!d.NewHotWeight.HasValue || d.NewHotWeight.Value <= 0) reasons.Add("HotWeight missing or zero");
                if (string.IsNullOrWhiteSpace(d.NewGrade)) reasons.Add("Grade missing");
                if (!d.NewHealthScore.HasValue || d.NewHealthScore.Value < 1 || d.NewHealthScore.Value > 5)
                    reasons.Add("HealthScore missing or out of 1–5");
                d.FlagReason = string.Join("; ", reasons);
                if (!fresh.FlaggedRows.Any(f => f.ControlNo == d.ControlNo))
                    fresh.FlaggedRows.Add(d);
            }

            _logger.LogInformation(
                "[HW-MERGE] Preserved {Count} manually-edited rows; dropped {Stale} stale Loaded markers; demoted {Demoted} incomplete rows. Final buckets: AutoRows={Auto} (OK={Ok}, Loaded={Loaded}), FlaggedRows={Flagged}, DupRows={Dups}, Sum={Sum}.",
                preserved, staleLoadedDropped, demoteFromAuto.Count,
                fresh.AutoRows.Count,
                fresh.AutoRows.Count(r => r.Status == "OK"),
                fresh.AutoRows.Count(r => r.Status == "Loaded"),
                fresh.FlaggedRows.Count,
                fresh.DupRows.Count,
                fresh.AutoRows.Count + fresh.FlaggedRows.Count + fresh.DupRows.Count);

            return fresh;
        }

        private static string TagIdentityKey(HotWeightPreviewRow r)
        {
            // Stable cross-pull identifier when ControlNo isn't yet assigned.
            // We use whatever tag identifiers are present, joined by '|'.
            var parts = new[] { r.FileBackTag, r.FileTag1, r.FileTag2 }
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Select(t => t!.Trim().ToUpperInvariant());
            return string.Join("|", parts);
        }

        // HOT WEIGHT — REFRESH FROM HOT SCALE (manual pull trigger).
        // Reads existing staging for smart-merge, runs the SQL query against
        // Hot Scale, builds a synthetic Excel, runs the matching pipeline,
        // smart-merges with existing staging (preserves manual fixes), then
        // redirects to GET to render. This is the only path that hits the
        // Hot Scale DB — manual refresh keeps page loads fast.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> HotWeightRefreshFromHotScale()
        {
            // 1. Read existing shared staging so we can smart-merge into it
            //    (preserves any manual fixes operators have made since last pull).
            HotWeightImportViewModel? existingVm = null;
            try
            {
                var existingJson = await StagingBridge.ReadSharedAsync(
                    _stagingService,
                    StagingBridge.Types.HotWeight,
                    StagingBridge.SharedHotWeightKey);
                if (!string.IsNullOrEmpty(existingJson))
                    existingVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(existingJson);
            }
            catch (Exception rex)
            {
                _logger.LogWarning(rex, "[HW-AUTO] Could not read existing shared staging; treating as empty.");
            }

            // 2. Resolve the date to pull. Empty/missing override → DateTime.Today.
            DateTime pullDate = DateTime.Today;
            var overrideStr = _configuration["AppSettings:HotScaleDateOverride"];
            if (!string.IsNullOrWhiteSpace(overrideStr)
                && DateTime.TryParse(overrideStr, System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var overrideDate))
            {
                pullDate = overrideDate.Date;
                _logger.LogInformation("[HW-AUTO] Using HotScaleDateOverride from config: {Date:yyyy-MM-dd}", pullDate);
            }

            // 3. Run the pull
            HotWeightImportViewModel? freshVm = null;
            int rowsPulled = 0;
            try
            {
                var rows = await ReadHotWeightRowsFromHotScaleDbAsync(pullDate);
                rowsPulled = rows.Count;
                _logger.LogInformation("[HW-AUTO] Pulled {Count} rows from Hot Scale for {Date:yyyy-MM-dd}",
                    rowsPulled, pullDate);

                if (rowsPulled > 0)
                {
                    var syntheticFile = BuildHotScaleSyntheticFile(rows, pullDate);
                    freshVm = await ParseAndMatchHotWeightFileAsync(syntheticFile, treatAsAutoPull: true);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[HW-AUTO] Hot Scale refresh failed; existing staging unchanged.");
                TempData["ErrorMessage"] = "Couldn't pull today's hot weights from Hot Scale. Existing data is unchanged.";
                return RedirectToAction(nameof(HotWeightImport));
            }

            // 4. Smart-merge & persist
            HotWeightImportViewModel finalVm;
            if (freshVm != null)
            {
                // Pre-compute the set of bills currently in KillStatus=Killed.
                // The merge uses this to validate "Loaded" markers from existing
                // staging — if the bill is no longer killed (deleted / reset to
                // Pending), the stale marker is dropped.
                HashSet<int>? killedControlNos = null;
                try
                {
                    var allKilled = await _animalService.GetAllAsync();
                    killedControlNos = allKilled
                        .Where(a => string.Equals(a.KillStatus, "Killed", StringComparison.OrdinalIgnoreCase))
                        .Select(a => a.ControlNo)
                        .ToHashSet();
                    _logger.LogInformation("[HW-AUTO] {Count} bills currently in KillStatus=Killed (used to validate Loaded markers).", killedControlNos.Count);
                }
                catch (Exception kex)
                {
                    _logger.LogWarning(kex, "[HW-AUTO] Could not load currently-killed bills; merge will preserve all Loaded markers conservatively.");
                }

                finalVm = MergeHotWeightStaging(freshVm, existingVm, killedControlNos);
                try
                {
                    var mergedJson = System.Text.Json.JsonSerializer.Serialize(finalVm);
                    await StagingBridge.WriteSharedAsync(
                        _stagingService,
                        StagingBridge.Types.HotWeight,
                        StagingBridge.SharedHotWeightKey,
                        mergedJson,
                        sourceFileName: $"HotScale {pullDate:yyyy-MM-dd} pulled {DateTime.Now:HH:mm}");
                    var uniqueAfterDedup = finalVm.TotalInExcel;
                    string dedupNote = (rowsPulled > uniqueAfterDedup)
                        ? $" ({rowsPulled - uniqueAfterDedup} duplicate scale-machine rows merged to {uniqueAfterDedup} unique animals)"
                        : "";
                    TempData["SuccessMessage"] = $"Refreshed — pulled {rowsPulled} rows from Hot Scale for {pullDate:yyyy-MM-dd}{dedupNote}.";
                }
                catch (Exception wex)
                {
                    _logger.LogWarning(wex, "[HW-AUTO] Could not write merged VM to shared staging.");
                    TempData["ErrorMessage"] = "Pull succeeded but couldn't save staging. Try again.";
                }
            }
            else if (existingVm != null)
            {
                // 0 rows from Hot Scale — keep existing staging untouched.
                TempData["InfoMessage"] = $"No animals weighed yet for {pullDate:yyyy-MM-dd}. Existing data is unchanged.";
            }
            else
            {
                // No existing staging, no fresh data — clear out so the page renders empty.
                await StagingBridge.ClearSharedAsync(
                    _stagingService,
                    StagingBridge.Types.HotWeight,
                    StagingBridge.SharedHotWeightKey);
                TempData["InfoMessage"] = $"No animals weighed yet for {pullDate:yyyy-MM-dd}.";
            }

            return RedirectToAction(nameof(HotWeightImport));
        }

    }
}

// AJAX request models 
public class ExcelFixRowRequest
{
    public int      RowNum              { get; set; }
    public string?  VendorName          { get; set; }
    public string?  TagNumber1          { get; set; }
    public string?  PurchaseType        { get; set; }
    public DateTime? PurchaseDate       { get; set; }
    public string?  TagNumber2          { get; set; }
    public string?  Tag3                { get; set; }
    public string?  AnimalType          { get; set; }
    public string?  AnimalType2         { get; set; }
    public decimal  LiveWeight          { get; set; }
    public decimal  LiveRate            { get; set; }
    public decimal? HotWeight           { get; set; }
    public string?  Grade               { get; set; }
    public int?     HealthScore         { get; set; }
    public string?  Comment             { get; set; }
    public string?  AnimalControlNumber { get; set; }
    public string?  OfficeUse2          { get; set; }
    public string?  State               { get; set; }
    public string?  BuyerName           { get; set; }
    public string?  VetName             { get; set; }
}

public class ExcelDeleteRowRequest
{
    public int RowNum { get; set; }
}

public class FixedFlaggedRow
{
    public string RowKey       { get; set; } = "";

    public int     ControlNo   { get; set; }

    public string AnimalControlNumber { get; set; } = "";
    public decimal Side1       { get; set; }
    public decimal Side2       { get; set; }
    public string  Grade       { get; set; } = "";
    public int     HealthScore { get; set; }
}

public class SaveEditsRequest
{
    public List<AnimalRowDto> Rows { get; set; } = new();
}

public class MarkKilledRequest
{
    public string KillDate { get; set; } = "";
    public List<AnimalRowDto> Rows { get; set; } = new();
}

public class AnimalRowDto
{
public int ControlNo { get; set; }

public int OriginalControlNo { get; set; } 
public string AnimalControlNumber { get; set; } = "";
public string KillDate { get; set; } = "";
public decimal LiveWeight { get; set; }
public decimal LiveRate { get; set; }
public decimal HotWeight { get; set; }
public string Grade { get; set; } = "";
public int HealthScore { get; set; }
public bool IsCondemned { get; set; }
public string State { get; set; } = "";
public string VetName { get; set; } = "";
public string OfficeUse2 { get; set; } = "";
public string Comment { get; set; } = "";
public string PurchaseType { get; set; } = "";
public string AnimalType { get; set; } = "";

public bool HwImported { get; set;}
}

// Row shape returned by the Hot Scale auto-pull SQL.
// Columns mirror the SSRS HotScaleReport so values can be projected
// directly into the synthetic workbook the existing parser expects.
public class HotScaleRow
{
    public string  TimeKilled    { get; set; } = "";
    public string  AnimalNumber  { get; set; } = "";
    public string  SexCode       { get; set; } = "";
    public string  AnimalType    { get; set; } = "";
    public string  BackTag       { get; set; } = "";
    public string  KillLotNumber { get; set; } = "";
    public string  Tag1          { get; set; } = "";
    public string  Tag2          { get; set; } = "";
    public string  ProgramCode   { get; set; } = "";
    public string  CountryCode   { get; set; } = "";
    public decimal? Side1        { get; set; }
    public decimal? Side2        { get; set; }
    public decimal? LiveWeight   { get; set; }
    public string  Side1Grade    { get; set; } = "";
    public string  Side2Grade    { get; set; } = "";
    public int?    HealthScore   { get; set; }
}