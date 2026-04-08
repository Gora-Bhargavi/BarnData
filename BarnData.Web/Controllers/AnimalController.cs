using BarnData.Core.Services;
using BarnData.Data.Entities;
using BarnData.Web.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace BarnData.Web.Controllers
{
    public class AnimalController : Controller
    {
        private readonly IAnimalService _animalService;
        private readonly IVendorService _vendorService;

        public AnimalController(IAnimalService animalService, IVendorService vendorService)
        {
            _animalService = animalService;
            _vendorService = vendorService;
        }

        // INDEX — show ALL animals with status filter 
        public async Task<IActionResult> Index(DateTime? killDate, int? vendorId, string? status)
        {
            var vendors = await _vendorService.GetAllActiveAsync();
            IEnumerable<BarnData.Data.Entities.Animal> animals;

            // Default: show all pending animals (most useful daily view)
            if (status == "killed" && killDate.HasValue)
            {
                animals = await _animalService.GetByKillDateAsync(killDate.Value, vendorId);
                ViewBag.StatusFilter = "killed";
                ViewBag.KillDate = killDate.Value.ToString("yyyy-MM-dd");
            }
            else if (status == "killed")
            {
                animals = await _animalService.GetByKillDateAsync(DateTime.Today, vendorId);
                ViewBag.StatusFilter = "killed";
                ViewBag.KillDate = DateTime.Today.ToString("yyyy-MM-dd");
            }
            else if (status == "all")
            {
                animals = await _animalService.GetAllAsync(vendorId);
                ViewBag.StatusFilter = "all";
                ViewBag.KillDate = DateTime.Today.ToString("yyyy-MM-dd");
            }
            else
            {
                // Default: pending animals
                animals = await _animalService.GetPendingAsync(vendorId);
                ViewBag.StatusFilter = "pending";
                ViewBag.KillDate = DateTime.Today.ToString("yyyy-MM-dd");
            }

            ViewBag.VendorId   = vendorId;
            ViewBag.VendorList = new SelectList(vendors, "VendorID", "VendorName", vendorId);
            ViewBag.TotalCount = animals.Count();
            ViewBag.TotalLiveWeight = animals.Sum(a => a.LiveWeight);
            ViewBag.TotalHotWeight  = animals.Sum(a => a.HotWeight ?? 0);

            return View(animals);
        }

        //  CREATE GET — blank entry form 
        public async Task<IActionResult> Create()
        {
            var vm = new AnimalViewModel
            {
                KillDate     = DateTime.Today,
                PurchaseDate = DateTime.Today,
            };

            await PopulateVendorDropdown(vm);
            return View(vm);
        }

        // CREATE POST — save new animal record
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(AnimalViewModel vm)
        {
            // Handle vendor — either selected from list (VendorID > 0)
            // or typed as free text (VendorID = 0, VendorNameFreeText has value)
            if (vm.VendorID == 0 && !string.IsNullOrWhiteSpace(vm.VendorNameFreeText))
            {
                vm.VendorID = await _vendorService.GetOrCreateAsync(vm.VendorNameFreeText.Trim());
            }

            // Clear VendorID model error — it's set via JS hidden field
            ModelState.Remove("VendorID");

            // Manual vendor validation
            if (vm.VendorID == 0 && string.IsNullOrWhiteSpace(vm.VendorNameFreeText))
            {
                ModelState.AddModelError("VendorID", "Vendor is required.");
            }

            if (!ModelState.IsValid)
            {
                await PopulateVendorDropdown(vm);
                return View(vm);
            }

            // Check weight warning — does not block, just flags
            vm.ShowWeightWarning = _animalService.IsWeightOutOfRange(vm.LiveWeight);

            // If weight is out of range AND user hasn't confirmed yet — show warning
            if (vm.ShowWeightWarning && !vm.WeightWarningConfirmed)
            {
                await PopulateVendorDropdown(vm);
                ModelState.AddModelError("LiveWeight",
                    $"Live weight {vm.LiveWeight} lbs is outside the expected range (300–2,500 lbs). " +
                    "Please confirm this is correct by checking the box below.");
                return View(vm);
            }

            var animal = MapToEntity(vm);
            var (success, error) = await _animalService.CreateAsync(animal);

            if (!success)
            {
                ModelState.AddModelError("TagNumber1", error);
                await PopulateVendorDropdown(vm);
                return View(vm);
            }

            TempData["SuccessMessage"] = $"Record saved — Control No. {animal.ControlNo}. Tag: {animal.TagNumber1}";

            // Save & Add Another — carry sticky fields to the next form
            if (Request.Form.ContainsKey("saveAndAdd"))
            {
                // Store sticky fields in TempData to pre-populate next form
                TempData["StickyVendorId"]    = vm.VendorID;
                TempData["StickyPurchaseType"]= vm.PurchaseType;
                TempData["StickyPurchaseDate"]= vm.PurchaseDate.ToString("yyyy-MM-dd");
                TempData["StickyKillDate"]    = vm.KillDate?.ToString("yyyy-MM-dd");
                TempData["StickyLiveRate"]    = vm.LiveRate.ToString();
                TempData["StickyConsRate"]    = vm.ConsignmentRate?.ToString();
                TempData["StickyProgramCode"] = vm.ProgramCode;
                return RedirectToAction(nameof(CreateSticky));
            }

            return RedirectToAction(nameof(Index), new { status = "pending" });
        }

        //  CREATE STICKY — blank form with fields pre-filled 
        public async Task<IActionResult> CreateSticky()
        {
            var vm = new AnimalViewModel
            {
                PurchaseDate = DateTime.Today,
                KillDate     = DateTime.Today,
            };

            // Restore sticky fields from TempData
            if (TempData["StickyVendorId"] is int vendorId && vendorId > 0)
                vm.VendorID = vendorId;

            if (TempData["StickyPurchaseType"] is string pt && !string.IsNullOrEmpty(pt))
                vm.PurchaseType = pt;

            if (TempData["StickyPurchaseDate"] is string pd && DateTime.TryParse(pd, out var purchDate))
                vm.PurchaseDate = purchDate;

            if (TempData["StickyKillDate"] is string kd && DateTime.TryParse(kd, out var killDate))
                vm.KillDate = killDate;

            if (TempData["StickyLiveRate"] is string lr && decimal.TryParse(lr, out var liveRate))
                vm.LiveRate = liveRate;

            if (TempData["StickyConsRate"] is string cr && decimal.TryParse(cr, out var consRate))
                vm.ConsignmentRate = consRate;

            if (TempData["StickyProgramCode"] is string prog && !string.IsNullOrEmpty(prog))
                vm.ProgramCode = prog;

            // Keep StickyVendorId in TempData for the view to restore vendor search text
            TempData.Keep("StickyVendorId");

            await PopulateVendorDropdown(vm);
            return View("Create", vm);
        }

        // ── EDIT GET — pre-filled form ────────────────────────────────────
        public async Task<IActionResult> Edit(int id)
        {
            var animal = await _animalService.GetByControlNoAsync(id);
            if (animal == null) return NotFound();

            var vm = MapToViewModel(animal);
            await PopulateVendorDropdown(vm);
            return View(vm);
        }

        // ── EDIT POST — save changes ──────────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, AnimalViewModel vm)
        {
            if (id != vm.ControlNo) return BadRequest();

            // Handle vendor free-text entry
            if (vm.VendorID == 0 && !string.IsNullOrWhiteSpace(vm.VendorNameFreeText))
            {
                vm.VendorID = await _vendorService.GetOrCreateAsync(vm.VendorNameFreeText.Trim());
            }

            ModelState.Remove("VendorID");

            if (!ModelState.IsValid)
            {
                await PopulateVendorDropdown(vm);
                return View(vm);
            }

            vm.ShowWeightWarning = _animalService.IsWeightOutOfRange(vm.LiveWeight);
            if (vm.ShowWeightWarning && !vm.WeightWarningConfirmed)
            {
                await PopulateVendorDropdown(vm);
                ModelState.AddModelError("LiveWeight",
                    $"Live weight {vm.LiveWeight} lbs is outside the expected range (300–2,500 lbs). " +
                    "Confirm this is correct by checking the box below.");
                return View(vm);
            }

            var animal = MapToEntity(vm);
            var (success, error) = await _animalService.UpdateAsync(animal);

            if (!success)
            {
                ModelState.AddModelError("TagNumber1", error);
                await PopulateVendorDropdown(vm);
                return View(vm);
            }

            TempData["SuccessMessage"] = $"Animal record #{id} updated successfully.";
            return RedirectToAction(nameof(Index),
                new { killDate = vm.KillDate.HasValue ? vm.KillDate.Value.ToString("yyyy-MM-dd") : DateTime.Today.ToString("yyyy-MM-dd") });
        }

        // ── DETAIL — read-only view ───────────────────────────────────────
        public async Task<IActionResult> Detail(int id)
        {
            var animal = await _animalService.GetByControlNoAsync(id);
            if (animal == null) return NotFound();

            return View(animal);
        }

        // ── DELETE POST — soft delete ─────────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id, string killDate)
        {
            await _animalService.DeleteAsync(id);
            TempData["SuccessMessage"] = $"Animal record #{id} removed from today's list.";
            return RedirectToAction(nameof(Index), new { killDate });
        }

        // ── VENDOR SEARCH — called via AJAX as user types ────────────────
        [HttpGet]
        public async Task<IActionResult> SearchVendors(string term)
        {
            var vendors = await _vendorService.GetAllActiveAsync();
            var matches = vendors
                .Where(v => string.IsNullOrEmpty(term) ||
                            v.VendorName.Contains(term, StringComparison.OrdinalIgnoreCase))
                .Select(v => new { id = v.VendorID, name = v.VendorName })
                .Take(10)
                .ToList();
            return Json(matches);
        }

        // ── TAG DUPLICATE CHECK — called via AJAX on blur ─────────────────
        [HttpGet]
        public async Task<IActionResult> CheckTag(
            string tag1, int vendorId, int? controlNo = null)
        {
            if (string.IsNullOrWhiteSpace(tag1) || vendorId == 0)
                return Json(new { isDuplicate = false });

            bool isDuplicate = await _animalService.IsTagDuplicateAsync(
                tag1, vendorId, controlNo);

            return Json(new { isDuplicate });
        }

        // ── EXPORT EXCEL ──────────────────────────────────────────────────
        [HttpGet]
        public async Task<IActionResult> Export(
            int? vendorId, string? status,
            DateTime? killDateFrom, DateTime? killDateTo,
            DateTime? purchDateFrom, DateTime? purchDateTo)
        {
            var filter = new ExportFilter
            {
                VendorId      = vendorId,
                Status        = string.IsNullOrEmpty(status) ? null : status,
                KillDateFrom  = killDateFrom,
                KillDateTo    = killDateTo,
                PurchDateFrom = purchDateFrom,
                PurchDateTo   = purchDateTo,
            };

            var animals = (await _animalService.GetFilteredAsync(filter)).ToList();

            using var wb = new ClosedXML.Excel.XLWorkbook();
            var ws = wb.Worksheets.Add("Animals");

            // Headers
            var headers = new[]
            {
                "Control No", "Animal Type", "Tag Number One", "Tag Number Two",
                "Purchase Date", "Purchase Type", "Vendor", "Live Weight", "Live Rate",
                "Kill Date", "Hot Weight", "Grade", "H S", "Comments",
                "Animal Control Number", "Tag 3", "Office Use 2", "State",
                "Buyer", "Animal Type 2", "Vet Name", "Kill Status", "Program Code",
                "Is Condemned"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#1e2f47");
                cell.Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
            }

            // Data rows
            int row = 2;
            foreach (var a in animals)
            {
                ws.Cell(row, 1).Value  = a.ControlNo;
                ws.Cell(row, 2).Value  = a.AnimalType;
                ws.Cell(row, 3).Value  = a.TagNumber1;
                ws.Cell(row, 4).Value  = a.TagNumber2 ?? "";
                ws.Cell(row, 5).Value  = a.PurchaseDate.ToString("MM/dd/yyyy");
                ws.Cell(row, 6).Value  = a.PurchaseType;
                ws.Cell(row, 7).Value  = a.Vendor?.VendorName ?? "";
                ws.Cell(row, 8).Value  = a.LiveWeight;
                ws.Cell(row, 9).Value  = a.LiveRate;
                ws.Cell(row, 10).Value = a.KillDate.HasValue ? a.KillDate.Value.ToString("MM/dd/yyyy") : "";
                ws.Cell(row, 11).Value = a.HotWeight.HasValue ? a.HotWeight.Value : 0;
                ws.Cell(row, 12).Value = a.Grade ?? "";
                ws.Cell(row, 13).Value = a.HealthScore.HasValue ? a.HealthScore.Value : 0;
                ws.Cell(row, 14).Value = a.Comment ?? "";
                ws.Cell(row, 15).Value = a.AnimalControlNumber ?? "";
                ws.Cell(row, 16).Value = a.Tag3 ?? "";
                ws.Cell(row, 17).Value = a.OfficeUse2 ?? "";
                ws.Cell(row, 18).Value = a.State ?? "";
                ws.Cell(row, 19).Value = a.BuyerName ?? "";
                ws.Cell(row, 20).Value = a.AnimalType2 ?? "";
                ws.Cell(row, 21).Value = a.VetName ?? "";
                ws.Cell(row, 22).Value = a.KillStatus;
                ws.Cell(row, 23).Value = a.ProgramCode;
                ws.Cell(row, 24).Value = a.IsCondemned ? "Yes" : "No";
                row++;
            }

            ws.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            stream.Position = 0;

            var fileName = $"Animals_Export_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        // ── EXPORT PAGE (filter form) ─────────────────────────────────────
        public async Task<IActionResult> ExportPage()
        {
            var vendors = await _vendorService.GetAllActiveAsync();
            ViewBag.VendorList = new Microsoft.AspNetCore.Mvc.Rendering.SelectList(
                vendors, "VendorID", "VendorName");
            ViewData["Title"]    = "Export Animals";
            ViewData["Subtitle"] = "Choose filters and download as Excel";
            return View();
        }
        private async Task PopulateVendorDropdown(AnimalViewModel vm)
        {
            var vendors = await _vendorService.GetAllActiveAsync();
            vm.VendorList = vendors.Select(v =>
                new SelectListItem(v.VendorName, v.VendorID.ToString()));
        }

        private static Animal MapToEntity(AnimalViewModel vm) => new()
        {
            ControlNo            = vm.ControlNo,
            VendorID             = vm.VendorID,
            TagNumber1           = vm.TagNumber1,
            TagNumber2           = string.IsNullOrEmpty(vm.TagNumber2) ? null : vm.TagNumber2,
            Tag3                 = vm.Tag3,
            AnimalType           = vm.AnimalType,
            AnimalType2          = vm.AnimalType2,
            ProgramCode          = vm.ProgramCode,
            PurchaseDate         = vm.PurchaseDate,
            PurchaseType         = vm.PurchaseType,
            LiveWeight           = vm.LiveWeight,
            LiveRate             = vm.LiveRate,
            ConsignmentRate      = vm.ConsignmentRate,
            KillDate             = vm.KillDate,
            HotWeight            = vm.HotWeight,
            Grade                = string.IsNullOrEmpty(vm.Grade) ? null : vm.Grade,
            Grade2               = vm.Grade2,
            HealthScore          = vm.HealthScore,
            FetalBlood           = vm.FetalBlood,
            Comment              = vm.Comment,
            AnimalControlNumber  = vm.AnimalControlNumber,
            State                = vm.State,
            BuyerName            = vm.BuyerName,
            VetName              = vm.VetName,
            OfficeUse2           = vm.OfficeUse2,
            KillStatus           = vm.KillStatus,
            Origin               = vm.Origin,
            IsCondemned          = vm.IsCondemned,
        };

        private static AnimalViewModel MapToViewModel(Animal a) => new()
        {
            ControlNo            = a.ControlNo,
            VendorID             = a.VendorID,
            TagNumber1           = a.TagNumber1,
            TagNumber2           = a.TagNumber2,
            Tag3                 = a.Tag3,
            AnimalType           = a.AnimalType,
            AnimalType2          = a.AnimalType2,
            ProgramCode          = a.ProgramCode,
            PurchaseDate         = a.PurchaseDate,
            PurchaseType         = a.PurchaseType,
            LiveWeight           = a.LiveWeight,
            LiveRate             = a.LiveRate,
            ConsignmentRate      = a.ConsignmentRate,
            KillDate             = a.KillDate,
            HotWeight            = a.HotWeight,
            Grade                = a.Grade,
            Grade2               = a.Grade2,
            HealthScore          = a.HealthScore,
            FetalBlood           = a.FetalBlood,
            Comment              = a.Comment,
            AnimalControlNumber  = a.AnimalControlNumber,
            State                = a.State,
            BuyerName            = a.BuyerName,
            VetName              = a.VetName,
            OfficeUse2           = a.OfficeUse2,
            KillStatus           = a.KillStatus,
            Origin               = a.Origin,
            IsCondemned          = a.IsCondemned,
        };
    }
}
