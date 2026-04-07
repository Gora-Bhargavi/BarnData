using BarnData.Core.Services;
using BarnData.Data.Entities;
using BarnData.Web.Models;
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

        // ── INDEX — show ALL animals with status filter ──────────────────
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

        // ── CREATE GET — blank entry form ─────────────────────────────────
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

        // ── CREATE POST — save new animal record ──────────────────────────
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

            TempData["SuccessMessage"] = $"Animal record saved. Control No: {animal.ControlNo}";

            // If Save & Add Another was clicked — go back to blank Create form
            if (Request.Form.ContainsKey("saveAndAdd"))
                return RedirectToAction(nameof(Create));

            return RedirectToAction(nameof(Index), new { status = "pending" });
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

        // ── HELPERS ───────────────────────────────────────────────────────
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
