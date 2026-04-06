using BarnData.Core.Services;
using BarnData.Data.Entities;
using BarnData.Web.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.Text.Json;

namespace BarnData.Web.Controllers
{
    public class AnimalController : Controller
    {
        private const string CreateCarryForwardKey = "CreateCarryForward";

        private sealed class CreateCarryForwardState
        {
            public int VendorID { get; set; }
            public string? VendorNameFreeText { get; set; }
            public string PurchaseType { get; set; } = string.Empty;
            public DateTime PurchaseDate { get; set; }
            public string AnimalType { get; set; } = string.Empty;
            public string? AnimalType2 { get; set; }
            public string ProgramCode { get; set; } = string.Empty;
            public DateTime KillDate { get; set; }
            public string Grade { get; set; } = string.Empty;
            public string? Grade2 { get; set; }
            public int HealthScore { get; set; }
            public string? Origin { get; set; }
            public string? State { get; set; }
            public string? BuyerName { get; set; }
            public string? VetName { get; set; }
            public string? OfficeUse2 { get; set; }
        }
        private readonly IAnimalService _animalService;
        private readonly IVendorService _vendorService;

        public AnimalController(IAnimalService animalService, IVendorService vendorService)
        {
            _animalService = animalService;
            _vendorService = vendorService;
        }

        // ── INDEX — animal list filtered by kill date ─────────────────────
        public async Task<IActionResult> Index(DateTime? killDate, int? vendorId)
        {
            var date = killDate ?? DateTime.Today;
            var vendors = await _vendorService.GetAllActiveAsync();
            var animals = await _animalService.GetByKillDateAsync(date, vendorId);

            ViewBag.KillDate   = date.ToString("yyyy-MM-dd");
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
            var vm = BuildViewModelFromCarryForward() ?? new AnimalViewModel
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
        public async Task<IActionResult> Create(AnimalViewModel vm, string? submitMode)
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
            vm.ShowWeightWarning = vm.LiveWeight.HasValue &&
                _animalService.IsWeightOutOfRange(vm.LiveWeight.Value);

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

            if (string.Equals(submitMode, "save-add-another", StringComparison.OrdinalIgnoreCase))
            {
                TempData["SuccessMessage"] = $"Animal record saved. Control No: {animal.ControlNo}. Ready for the next animal.";
                TempData[CreateCarryForwardKey] = JsonSerializer.Serialize(BuildCarryForwardModel(vm));
                return RedirectToAction(nameof(Create));
            }

            TempData["SuccessMessage"] = $"Animal record saved. Control No: {animal.ControlNo}";
            return RedirectToAction(nameof(Index),
                new { killDate = vm.KillDate.ToString("yyyy-MM-dd") });
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

            vm.ShowWeightWarning = vm.LiveWeight.HasValue &&
                _animalService.IsWeightOutOfRange(vm.LiveWeight.Value);
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
                new { killDate = vm.KillDate.ToString("yyyy-MM-dd") });
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
            string tag1, string killDate, int vendorId, int? controlNo = null)
        {
            if (string.IsNullOrWhiteSpace(tag1) || vendorId == 0)
                return Json(new { isDuplicate = false });

            if (!DateTime.TryParse(killDate, out var date))
                date = DateTime.Today;

            bool isDuplicate = await _animalService.IsTagDuplicateAsync(
                tag1, date, vendorId, controlNo);

            return Json(new { isDuplicate });
        }

        // ── HELPERS ───────────────────────────────────────────────────────
        private async Task PopulateVendorDropdown(AnimalViewModel vm)
        {
            var vendors = await _vendorService.GetAllActiveAsync();
            vm.VendorList = vendors.Select(v =>
                new SelectListItem(v.VendorName, v.VendorID.ToString()));

            if (vm.VendorID > 0 && string.IsNullOrWhiteSpace(vm.VendorNameFreeText))
            {
                vm.VendorNameFreeText = vendors
                    .FirstOrDefault(v => v.VendorID == vm.VendorID)
                    ?.VendorName;
            }
        }

        private AnimalViewModel? BuildViewModelFromCarryForward()
        {
            if (!TempData.TryGetValue(CreateCarryForwardKey, out var raw) || raw is not string json)
            {
                return null;
            }

            try
            {
                var state = JsonSerializer.Deserialize<CreateCarryForwardState>(json);
                if (state == null)
                {
                    TempData.Remove(CreateCarryForwardKey);
                    return null;
                }

                return new AnimalViewModel
                {
                    VendorID = state.VendorID,
                    VendorNameFreeText = state.VendorNameFreeText,
                    PurchaseType = state.PurchaseType,
                    PurchaseDate = state.PurchaseDate,
                    LiveRate = null,
                    AnimalType = state.AnimalType,
                    AnimalType2 = state.AnimalType2,
                    ProgramCode = state.ProgramCode,
                    KillDate = state.KillDate,
                    Grade = state.Grade,
                    Grade2 = state.Grade2,
                    HealthScore = state.HealthScore,
                    Origin = state.Origin,
                    State = state.State,
                    BuyerName = state.BuyerName,
                    VetName = state.VetName,
                    OfficeUse2 = state.OfficeUse2,
                    TagNumber1 = string.Empty,
                    TagNumber2 = null,
                    Tag3 = null,
                    AnimalControlNumber = string.Empty,
                    LiveWeight = null,
                    HotWeight = null,
                    FetalBlood = null,
                    Comment = null,
                    IsCondemned = false,
                    ShowWeightWarning = false,
                    WeightWarningConfirmed = false
                };
            }
            catch (JsonException)
            {
                TempData.Remove(CreateCarryForwardKey);
                return null;
            }
        }

        private static CreateCarryForwardState BuildCarryForwardModel(AnimalViewModel vm) => new()
        {
            VendorID             = vm.VendorID,
            VendorNameFreeText   = vm.VendorNameFreeText,
            PurchaseType         = vm.PurchaseType,
            PurchaseDate         = vm.PurchaseDate,
            AnimalType           = vm.AnimalType,
            AnimalType2          = vm.AnimalType2,
            ProgramCode          = vm.ProgramCode,
            KillDate             = vm.KillDate,
            Grade                = vm.Grade,
            Grade2               = vm.Grade2,
            HealthScore          = vm.HealthScore,
            Origin               = vm.Origin,
            State                = vm.State,
            BuyerName            = vm.BuyerName,
            VetName              = vm.VetName,
            OfficeUse2           = vm.OfficeUse2
        };

        private static Animal MapToEntity(AnimalViewModel vm) => new()
        {
            ControlNo            = vm.ControlNo,
            VendorID             = vm.VendorID,
            TagNumber1           = vm.TagNumber1,
            TagNumber2           = vm.TagNumber2 ?? string.Empty,
            Tag3                 = vm.Tag3,
            AnimalType           = vm.AnimalType,
            AnimalType2          = vm.AnimalType2,
            ProgramCode          = vm.ProgramCode,
            PurchaseDate         = vm.PurchaseDate,
            PurchaseType         = vm.PurchaseType,
            LiveWeight           = vm.LiveWeight!.Value,
            LiveRate             = vm.LiveRate!.Value,
            KillDate             = vm.KillDate,
            HotWeight            = vm.HotWeight,
            Grade                = vm.Grade,
            Grade2               = vm.Grade2,
            HealthScore          = vm.HealthScore,
            FetalBlood           = vm.FetalBlood,
            Comment              = vm.Comment ?? string.Empty,
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
