using BarnData.Core.Services;
using BarnData.Data.Entities;
using BarnData.Web.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
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

        public ImportController(IAnimalService animalService, IVendorService vendorService,
                                 IAnimalQueryService animalQueryService,
                                 IImportStagingService stagingService,
                                 ILogger<ImportController> logger)
        {
            _animalService = animalService;
            _vendorService = vendorService;
            _animalQueryService = animalQueryService;
            _stagingService = stagingService;
            _logger = logger;
        }

        //  SALE BILL IMPORT — GET 
        public IActionResult SaleBill()
        {
            return View();
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

             bool IsCompleteForKill(AnimalRowDto r) =>
                NormalizeAcn(r.AnimalControlNumber) != null
                && r.HotWeight > 0
                && !string.IsNullOrWhiteSpace(r.Grade)
                && r.HealthScore >= 1 && r.HealthScore <= 5;

            var incomplete = req.Rows.Where(r => !IsCompleteForKill(r)).ToList();
            if (incomplete.Any())
            {
                var sample = string.Join(", ", incomplete.Take(5).Select(x => x.ControlNo));
                return Json(new
                {
                    success = false,
                    message = $"Selected rows missing required fields (ACN, Hot Wt, Grade, HS). Ctrl No: {sample}"
                });
            }

            var validationErrors = req.Rows
            .Select(ValidateMarkKilledRow)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();
            if (validationErrors.Any())
            {
                return Json(new
                {
                    success = false,
                    message = string.Join(" ", validationErrors.Take(3))
                });
            }

            var animalData = req.Rows.Select(r => new KillAnimalData
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
            return Json(new {
                success  = true,
                message  = $"{count} animal{(count != 1 ? "s" : "")} marked as killed on {killDate:MM/dd/yyyy}.",
                redirect = Url.Action("Tally", "Report", new { killDate = killDate.ToString("yyyy-MM-dd") })
            });
        }

        //  MARK AS KILLED 
        public async Task<IActionResult> MarkKilled(int? vendorId, string? vendorIds)
        {
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
                        IsCondemned         = a.IsCondemned,
                        HotWeight    = hw != null ? hw.NewHotWeight  : a.HotWeight,
                        Grade        = hw != null ? hw.NewGrade       : a.Grade,
                        HealthScore  = hw != null ? hw.NewHealthScore : a.HealthScore,
                        HwImported   = hw != null,
                    };
                }).ToList()
            };

            return View(vm);
        }

        //  MARK AS KILLED — FAST, PAGINATED  (Phase 2b)
        public async Task<IActionResult> MarkKilledFast(
            string? vendorIds,
            int page = 1,
            int pageSize = 100,
            string? q = null)
        {
            // Sanity: reasonable page size limits
            if (page < 1) page = 1;
            if (pageSize < 100) pageSize = 100;
            if (pageSize > 500) pageSize = 500;

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

            // Hot Weight pre-fill (same logic as the legacy MarkKilled action)
            var hwLoaded = (TempData["HWLoaded"] as string == "1")
                        || (HttpContext.Session.GetString("HWLoaded") == "1");
            var hwJson = TempData.Peek("HWPreview") as string
                    ?? HttpContext.Session.GetString("HWPreview");
            HttpContext.Session.Remove("HWLoaded");
            var hwLookup = new Dictionary<int, HotWeightPreviewRow>();

            if (hwLoaded && !string.IsNullOrEmpty(hwJson))
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
                        TempData["HWLoadedCount"] = hwLookup.Count.ToString();
                        TempData["HWFlaggedCount"] = hwVm.FlaggedRows.Count(r => !hwLookup.ContainsKey(r.ControlNo)).ToString();
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "[MKFast-GET] Exception deserializing HWPreview JSON");
                }
            }

            // Build the same MarkKilledViewModel the legacy page uses, so we can
            // reuse PendingAnimalRow / MarkKilledViewModel shape and keep the save
            // endpoints fully compatible.
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
                        IsCondemned         = a.IsCondemned,
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
        public async Task<IActionResult> HotWeightImport()
        {
            var json = await StagingBridge.ReadAsync(
                HttpContext.Session, _stagingService,
                StagingBridge.Types.HotWeight,
                StagingBridge.GetUserKey(HttpContext));

            HotWeightImportViewModel? vm = null;
            if (!string.IsNullOrEmpty(json))
            {
                try { vm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(json); }
                catch { /* corrupt payload — show empty tabs */ }
            }

            if (vm != null) ViewBag.RestoredFromStaging = true;
            return View("HotWeightPreview", vm ?? new HotWeightImportViewModel());
        }

        // HOT WEIGHT IMPORT — PREVIEW POST
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

            var bullGrades = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "BB", "LB", "UB", "FB" };

            var cowGrades = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "CN", "SH", "CT", "B1", "B2", "BR" };


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
                    return View("HotWeightPreview", vm);
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
                    return View("HotWeightPreview", vm);
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
                    return View("HotWeightPreview", vm);
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
                    return View("HotWeightPreview", vm);
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
                    var unassigned = hits
                    .Where(a => IsAcnMissing(a.AnimalControlNumber))
                    .ToList();

                    if (!unassigned.Any() && hits.Any())
                    {
                        reason = $"{context} only matches already-assigned ACN records";
                    }

                    return unassigned;
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

                    //Direct ACN match 
                    if (!string.IsNullOrEmpty(acn))
                        acnAnimals.TryGetValue(acn, out animal);

                    
                    
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

                                matchedControlNos.Add(bestAnimal.ControlNo);
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

                    goto DoneMatch;
                    AddFlag:
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
                        Status         = "Flag",
                        FlagReason     = flagReason,
                        // Store candidates if this is a multi-match flag (for picker UI)
                        Candidates     = _candidateBuffer
                    });
                    _candidateBuffer = null;
                    continue;
                    DoneMatch:
                    // Dedup - silently skip if we already matched this animal
                    // (hot scale machine sometimes writes 2 rows per animal - both have same BackTag)
                    if (matchedControlNos.Contains(animal.ControlNo))
                    {
                        vm.TotalInExcel--;   // don't count the silent duplicate in total
                        continue;
                    }
                    matchedControlNos.Add(animal.ControlNo);

                    // Use the matched animal's ACN (or write from Excel if tag-matched)
                    var resolvedAcn = !string.IsNullOrEmpty(animal.AnimalControlNumber)
                        ? animal.AnimalControlNumber.TrimStart('0')
                        : acn;

                    vm.Matched++;

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

                    // Side checks
                    bool side1Ok = s1.HasValue && s1.Value > 0;
                    bool side2Ok = s2.HasValue && s2.Value > 0;

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

                    // HealthScore validation
                    if (!hs.HasValue)
                        flags.Add("HealthScore missing");
                    else if (hs.Value < 1 || hs.Value > 5)
                        flags.Add($"HealthScore {hs} out of range 1–5");

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
                        row.Status = "OK";
                        vm.AutoRows.Add(row);
                    }
                }
                var hwJson = System.Text.Json.JsonSerializer.Serialize(vm);
                TempData["HWPreview"] = hwJson;
                //HttpContext.Session.SetString("HWPreview", hwJson);
                await StagingBridge.WriteAsync(
                    HttpContext.Session, _stagingService,
                    StagingBridge.Types.HotWeight,
                    StagingBridge.GetUserKey(HttpContext),
                    hwJson,
                    sourceFileName: file?.FileName);
                _logger.LogInformation("[HW-IMPORT] Parsed: TotalInExcel={Total}, AutoRows={Auto}, FlaggedRows={Flagged}, JsonLength={Len}",
                    vm.TotalInExcel, vm.AutoRows.Count, vm.FlaggedRows.Count, hwJson.Length);
                _logger.LogInformation("[HW-IMPORT] Sample AutoRow ControlNos: {Ids}",
                    string.Join(", ", vm.AutoRows.Take(5).Select(r => $"{r.ControlNo}={r.NewHotWeight}")));
            }
            catch (Exception ex)
            {
                vm.Errors.Add($"Parse error: {ex.Message}");
                return View("HotWeightPreview", vm);
            }

            return View("HotWeightPreview", vm);
        }

        // HOT WEIGHT - LOAD SELECTED ROWS INTO MARK AS KILLED (no DB write yet)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult HotWeightLoadToMarkKilled([FromForm] string? selectedControlNos, [FromForm] string? fixedFlaggedJson)
        {
            var json = HttpContext.Session.GetString("HWPreview")
                    ?? TempData.Peek("HWPreview") as string;
            if (string.IsNullOrEmpty(json))
            { TempData["ErrorMessage"] = "Preview session expired. Please re-upload the file."; return RedirectToAction(nameof(HotWeightImport)); }

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

            // Merge any fixed flagged rows from the JS into the vm
            if (!string.IsNullOrEmpty(fixedFlaggedJson))
            {
                try
                {
                    var fixedRows = System.Text.Json.JsonSerializer.Deserialize<List<FixedFlaggedRow>>(fixedFlaggedJson);
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
                                    existing.Side1 = fix.Side1;
                                    existing.Side2 = fix.Side2;
                                    existing.NewHotWeight = fix.Side1 + fix.Side2;
                                    existing.NewGrade = fix.Grade;
                                    existing.NewHealthScore = fix.HealthScore;
                                    existing.Status = "OK";
                                    existing.FlagReason = "";
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

            // Pass the filtered/merged subset to Mark as Killed via TempData ONLY
            TempData["HWLoaded"]  = "1";
            TempData["HWPreview"] = json;   // filtered subset for Mark as Killed
            HttpContext.Session.SetString("HWLoaded", "1");

            // ── Update master session — mark loaded rows as "Loaded", keep flagged rows intact ──
            var masterJson = HttpContext.Session.GetString("HWPreview");
            if (!string.IsNullOrEmpty(masterJson))
            {
                try
                {
                    var masterVm = System.Text.Json.JsonSerializer.Deserialize<HotWeightImportViewModel>(masterJson);
                    if (masterVm != null)
                    {
                        // If selectedControlNos is provided use it; otherwise mark ALL OK rows as Loaded
                        HashSet<int> loadedIds;
                        if (!string.IsNullOrEmpty(selectedControlNos))
                        {
                            loadedIds = selectedControlNos.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(s => int.TryParse(s.Trim(), out int id) ? id : 0)
                                .Where(id => id > 0).ToHashSet();
                        }
                        else
                        {
                            // "Load all" - mark every OK (non-Loaded) row as Loaded
                            loadedIds = masterVm.AutoRows
                                .Where(r => r.Status == "OK")
                                .Select(r => r.ControlNo).ToHashSet();
                        }

                        int marked = 0;
                        foreach (var r in masterVm.AutoRows.Where(r => loadedIds.Contains(r.ControlNo) && r.Status == "OK"))
                        { r.Status = "Loaded"; marked++; }

                        var updatedMaster = System.Text.Json.JsonSerializer.Serialize(masterVm);
                        HttpContext.Session.SetString("HWPreview", updatedMaster);
                        _logger.LogInformation("[HW-LOAD] {Count} rows marked Loaded in master session", marked);
                    }
                }
                catch (Exception ex) { _logger.LogWarning(ex, "[HW-LOAD] Could not update master session"); }
            }

            _logger.LogInformation("[HW-LOAD] Loading into MarkKilled. Json length={Len}", json.Length);
            return RedirectToAction("MarkKilled");
        }

        // HOT WEIGHT - CLEAR SESSION
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> HotWeightClearSession()
        {
            await StagingBridge.ClearAsync(
                HttpContext.Session, _stagingService,
                StagingBridge.Types.HotWeight,
                StagingBridge.GetUserKey(HttpContext));

            HttpContext.Session.Remove("HWLoaded");
            TempData.Remove("HWPreview");
            TempData.Remove("HWLoaded");
            TempData["SuccessMessage"] = "Hot Weight import session cleared.";
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

    // Imported HW rows can save without Live Wt for consignment.
    // Live Wt will be backfilled from Hot Wt in the save mapping.
    if (!isHwImported && isConsignment && row.HotWeight > 0 && row.LiveWeight <= 0)
        return $"Ctrl No {row.ControlNo}: Live Wt is required for consignment when Hot Wt is entered.";

    if (row.HotWeight > 0 && row.LiveWeight > 0 && row.HotWeight > row.LiveWeight)
        return $"Ctrl No {row.ControlNo}: Hot Wt ({row.HotWeight:N1}) cannot exceed Live Wt ({row.LiveWeight:N1}).";

    var grade = (row.Grade ?? "").Trim().ToUpperInvariant();
    if (string.IsNullOrEmpty(grade))
        return null;

    // HW-import grade/type mismatches are already flagged during preview.
    // Do not block persistence here for imported rows.
    if (isHwImported)
        return null;

    var bullGrades = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "BB", "LB", "UB", "FB"
    };

    var cowGrades = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "CN", "SH", "CT", "B1", "B2", "BR"
    };

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

    if (hwImported)
        return null;

    var bullGrades = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "BB", "LB", "UB", "FB"
    };

    var cowGrades = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "CN", "SH", "CT", "B1", "B2", "BR"
    };

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
            // Always render ExcelPreview — upload form lives inside panel 0
            // Restore session if available, otherwise show empty tabs
            var sessionJson = HttpContext.Session.GetString("ExcelPreview");
            ExcelImportViewModel? vm = null;
            if (!string.IsNullOrEmpty(sessionJson))
            {
                try { vm = System.Text.Json.JsonSerializer.Deserialize<ExcelImportViewModel>(sessionJson); }
                catch { /* corrupt session — show empty tabs */ }
            }
            return View("ExcelPreview", vm ?? new ExcelImportViewModel());
        }

        // EXCEL IMPORT — CLEAR SESSION
        /*[HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ExcelClearSession()
        {
            HttpContext.Session.Remove("ExcelPreview");
            TempData.Remove("ExcelPreview");
            TempData["SuccessMessage"] = "Excel import session cleared.";
            return RedirectToAction(nameof(Excel));
        }*/

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

            // Show a banner indicating the preview was restored
            ViewBag.RestoredFromStaging = true;
            return View("ExcelPreview", vm);
        }

        //  EXCEL IMPORT — POST (parse & preview) 
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
public string AnimalControlNumber { get; set; } = "";
public string KillDate { get; set; } = "";
public decimal LiveWeight { get; set; }
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

