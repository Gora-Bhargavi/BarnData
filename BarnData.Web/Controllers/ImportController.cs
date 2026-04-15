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
namespace BarnData.Web.Controllers
{
    public class ImportController : Controller
    {
        private readonly IAnimalService _animalService;
        private readonly IVendorService _vendorService;

        public ImportController(IAnimalService animalService, IVendorService vendorService)
        {
            _animalService = animalService;
            _vendorService = vendorService;
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

        //  MARK AS KILLED 
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
                    IsCondemned        = a.IsCondemned,
                    HotWeight           = a.HotWeight,
                    Grade               = a.Grade,
                    HealthScore         = a.HealthScore,
                }).ToList()
            };

            return View(vm);
        }

        //Adding Post method to handle SaveMarkedEdits
        [HttpPost]
[ValidateAntiForgeryToken]
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
        return animalCtrlChanged || liveWeightChanged || hotWeightEntered || gradeEntered || hsEntered || condemnedChanged
        || stateChanged || vetChanged || office2Changed || commentChanged;
    }

    var editedIds = allIds.Where(IsEditedForSave).ToList();

    if (!editedIds.Any())
    {
        TempData["ErrorMessage"] = "No editable field changes found to save.";
        return RedirectToAction(nameof(MarkKilled), new { vendorId });
    }

    //Consignment validation: If hot weight entered, live weight is required and must be >= Hot weight
    var validationErrors = new List<string>();
        foreach (var id in editedIds)
        {
            var purchaseType = form[$"purchaseType_{id}"].FirstOrDefault() ?? "";
            var isConsignment = purchaseType.Contains("consignment", StringComparison.OrdinalIgnoreCase);

            bool hasHot = decimal.TryParse(form[$"hotWeight_{id}"], out var hot) && hot > 0;
            bool hasLive = decimal.TryParse(form[$"liveWeight_{id}"], out var live) && live > 0;

            if(!isConsignment) continue;

            if(hasHot && !hasLive)
                validationErrors.Add($"Ctrl No {id}: Live Wt is required for consignment when Hot Wt is entered.");
            else if (hasHot && hasLive && hot > live)
                validationErrors.Add($"Ctrl No {id}: Hot Wt ({hot:N1}) cannot exceed Live Wt ({live:N1})."); 
        }

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

        return new KillAnimalData
        {
            ControlNo = id,
            AnimalControlNumber = NullIfEmpty(form[$"animalCtrlNo_{id}"].FirstOrDefault()),
            KillDate = rowKillDate,
            LiveWeight = decimal.TryParse(form[$"liveWeight_{id}"], out var lw) && lw > 0 ? lw : null,
            HotWeight = decimal.TryParse(form[$"hotWeight_{id}"], out var hw) && hw > 0 ? hw : null,
            Grade = NullIfEmpty(form[$"grade_{id}"].FirstOrDefault()),
            HealthScore = int.TryParse(form[$"healthScore_{id}"], out var hs) && hs > 0 ? hs : null,
            IsCondemned = form[$"condemned_{id}"].Any(v => v == "true" || v == "on"),
            State = NullIfEmpty(form[$"state_{id}"].FirstOrDefault()),
            VetName = NullIfEmpty(form[$"vetName_{id}"].FirstOrDefault()),
            OfficeUse2 = NullIfEmpty(form[$"officeUse2_{id}"].FirstOrDefault()),
            Comment = NullIfEmpty(form[$"comment_{id}"].FirstOrDefault()),
        };
    }).ToList();

    int count = await _animalService.SaveKillDataAsync(animalData);

    TempData["SuccessMessage"] = $"{count} animal records updated. They remain pending until marked as killed.";
    return RedirectToAction(nameof(MarkKilled), new { vendorId });
}
        [HttpPost]
        [ValidateAntiForgeryToken]
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

            var validationErrors = new List<string>();
                foreach (var id in selectedIds)
                {
                    var purchaseType = form[$"purchaseType_{id}"].FirstOrDefault() ?? "";
                    var isConsignment = purchaseType.Contains("consignment", StringComparison.OrdinalIgnoreCase);
    
                    bool hasHot = decimal.TryParse(form[$"hotWeight_{id}"], out var hot) && hot > 0;
                    bool hasLive = decimal.TryParse(form[$"liveWeight_{id}"], out var live) && live > 0;
    
                    if(!isConsignment) continue;
    
                    if(hasHot && !hasLive)
                        validationErrors.Add($"Ctrl No {id}: Live Wt is required for consignment when Hot Wt is entered.");
                    else if (hasHot && hasLive && hot > live)
                        validationErrors.Add($"Ctrl No {id}: Hot Wt ({hot:N1}) cannot exceed Live Wt ({live:N1})."); 
                }

                if(validationErrors.Any())
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

                return new KillAnimalData
                {
                    ControlNo = id,
                    AnimalControlNumber = NullIfEmpty(form[$"animalCtrlNo_{id}"].FirstOrDefault()),
                    KillDate = rowKillDate ?? defaultKillDate,
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

        //  Helper 
        private static string? NullIfEmpty(string? s)
            => string.IsNullOrWhiteSpace(s) ? null : s;
        
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
            return View();
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
                    var exists = await _animalService.IsTagDuplicateAsync(tag1Safe, vendor.VendorID);
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

            // Store preview in TempData as JSON for confirm step
            TempData["ExcelPreview"] = System.Text.Json.JsonSerializer.Serialize(vm);
            return View("ExcelPreview", vm);
        }

        // EXCEL IMPORT - CONFIRM & SAVE 
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

            return RedirectToAction("Index", "Animal", new { status = "pending" });
        }
    }
}
