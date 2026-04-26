using BarnData.Core.Services;
using BarnData.Data.Entities;
using BarnData.Web.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Rotativa.AspNetCore;

namespace BarnData.Web.Controllers
{
    public class ReportController : Controller
    {
        private readonly IAnimalService _animalService;
        private readonly IVendorService _vendorService;

        public ReportController(IAnimalService animalService, IVendorService vendorService)
        {
            _animalService = animalService;
            _vendorService = vendorService;
        }

        // Phase 2e — parse a comma-separated vendor ids string (e.g. "3,7,12") into a list.
        // Returns an empty list for null/empty input.
        private static List<int> ParseVendorIds(string? vendorIds)
        {
            if (string.IsNullOrWhiteSpace(vendorIds)) return new List<int>();
            return vendorIds
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(s => int.TryParse(s.Trim(), out var id) ? id : 0)
                .Where(id => id > 0)
                .Distinct()
                .ToList();
        }

        // Phase 2e — load animals for a kill-date tally honoring either legacy
        // single-vendor id OR new multi-vendor ids. If both are supplied,
        // multi wins. If neither, returns all vendors for that kill date.
        private async Task<List<Animal>> LoadKilledAsync(DateTime date, int? legacyVendorId, List<int> multiIds)
        {
            // Multi-vendor path
            if (multiIds.Count > 0)
            {
                var results = new List<Animal>();
                foreach (var vid in multiIds)
                {
                    var filter = new ExportFilter
                    {
                        Status       = "Killed",
                        VendorId     = vid,
                        KillDateFrom = date,
                        KillDateTo   = date,
                    };
                    results.AddRange(await _animalService.GetFilteredAsync(filter));
                }
                // De-dupe defensively by ControlNo in case a bug ever made an animal appear for two vendors.
                return results
                    .GroupBy(a => a.ControlNo)
                    .Select(g => g.First())
                    .ToList();
            }

            // Legacy single-vendor path (back-compat for bookmarked URLs)
            var single = new ExportFilter
            {
                Status       = "Killed",
                VendorId     = legacyVendorId,
                KillDateFrom = date,
                KillDateTo   = date,
            };
            return (await _animalService.GetFilteredAsync(single)).ToList();
        }

        //  PAGE 1: KILLED ANIMALS LIST 
        public async Task<IActionResult> Tally(DateTime? killDate, int? vendorId, string? vendorIds)
        {
            var date    = killDate ?? DateTime.Today;
            var vendors = await _vendorService.GetAllActiveAsync();
            var multi   = ParseVendorIds(vendorIds);
            var animals = await LoadKilledAsync(date, vendorId, multi);

            ViewBag.KillDate   = date;
            ViewBag.KillDateStr = date.ToString("yyyy-MM-dd");
            ViewBag.VendorId   = vendorId;
            ViewBag.VendorIds  = vendorIds ?? "";
            ViewBag.VendorList = vendors.Select(v =>
                new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem(
                    v.VendorName, v.VendorID.ToString(),
                    multi.Contains(v.VendorID) || v.VendorID == vendorId));
            ViewBag.SelectedVendorNames = (multi.Count > 0
                    ? vendors.Where(v => multi.Contains(v.VendorID))
                    : vendors.Where(v => v.VendorID == vendorId))
                .Select(v => v.VendorName).ToList();
            ViewBag.TotalCount     = animals.Count;
            ViewBag.TotalCondemned = animals.Count(a => a.IsCondemned);
            ViewBag.TotalPassed    = animals.Count(a => !a.IsCondemned);
            ViewBag.TotalLiveWt    = animals.Sum(a => a.LiveWeight);
            ViewBag.TotalHotWt     = animals.Sum(a => a.HotWeight ?? 0);
            ViewBag.TotalCost      = animals.Sum(a => a.SaleCost);

            ViewData["Title"]    = "Nightly tally";
            ViewData["Subtitle"] = "Killed animals — " + date.ToString("dddd, MMMM d, yyyy");
            return View(animals);
        }

        //  PAGE 2: FILTER VIEW 
        public async Task<IActionResult> FilterView(
            int? vendorId, string? status,
            DateTime? killDateFrom, DateTime? killDateTo,
            DateTime? purchDateFrom, DateTime? purchDateTo)
        {
            var vendors = await _vendorService.GetAllActiveAsync();
            List<Animal> animals = new();

            bool hasFilter = vendorId.HasValue
                || !string.IsNullOrEmpty(status)
                || killDateFrom.HasValue || killDateTo.HasValue
                || purchDateFrom.HasValue || purchDateTo.HasValue;

            if (hasFilter)
            {
                var filter = new ExportFilter
                {
                    VendorId      = vendorId,
                    Status        = string.IsNullOrEmpty(status) || status == "all" ? null : status,
                    KillDateFrom  = killDateFrom,
                    KillDateTo    = killDateTo,
                    PurchDateFrom = purchDateFrom,
                    PurchDateTo   = purchDateTo,
                };
                animals = (await _animalService.GetFilteredAsync(filter)).ToList();
            }

            ViewBag.VendorId      = vendorId;
            ViewBag.Status        = status ?? "all";
            ViewBag.KillDateFrom  = killDateFrom?.ToString("yyyy-MM-dd");
            ViewBag.KillDateTo    = killDateTo?.ToString("yyyy-MM-dd");
            ViewBag.PurchDateFrom = purchDateFrom?.ToString("yyyy-MM-dd");
            ViewBag.PurchDateTo   = purchDateTo?.ToString("yyyy-MM-dd");
            ViewBag.HasFilter     = hasFilter;
            ViewBag.VendorList    = vendors.Select(v =>
                new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem(
                    v.VendorName, v.VendorID.ToString()));
            ViewBag.TotalCount     = animals.Count;
            ViewBag.TotalKilled    = animals.Count(a => a.KillStatus == "Killed");
            ViewBag.TotalPending   = animals.Count(a => a.KillStatus == "Pending");
            ViewBag.TotalCondemned = animals.Count(a => a.IsCondemned);
            ViewBag.TotalLiveWt    = animals.Sum(a => a.LiveWeight);
            ViewBag.TotalHotWt     = animals.Sum(a => a.HotWeight ?? 0);
            ViewBag.TotalCost      = animals.Sum(a => a.SaleCost);

            ViewData["Title"]    = "Filter view";
            ViewData["Subtitle"] = "Search all animal records by any combination of filters";
            return View(animals);
        }

        //  EXPORT EXCEL — PAGE 1 (killed list)
        public async Task<IActionResult> ExportKilledExcel(DateTime? killDate, int? vendorId, string? vendorIds)
        {
            var date    = killDate ?? DateTime.Today;
            var multi   = ParseVendorIds(vendorIds);
            var animals = await LoadKilledAsync(date, vendorId, multi);
            var bytes   = BuildExcel(animals, $"Killed animals — {date:MM/dd/yyyy}");
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"KilledAnimals_{date:yyyyMMdd}.xlsx");
        }

        //  EXPORT PDF — PAGE 1 
        public async Task<IActionResult> ExportKilledPdf(DateTime? killDate, int? vendorId, string? vendorIds)
        {
            var date    = killDate ?? DateTime.Today;
            var multi   = ParseVendorIds(vendorIds);
            var animals = await LoadKilledAsync(date, vendorId, multi);

            ViewBag.KillDate       = date;
            ViewBag.TotalCount     = animals.Count;
            ViewBag.TotalCondemned = animals.Count(a => a.IsCondemned);
            ViewBag.TotalPassed    = animals.Count(a => !a.IsCondemned);
            ViewBag.TotalLiveWt    = animals.Sum(a => a.LiveWeight);
            ViewBag.TotalHotWt     = animals.Sum(a => a.HotWeight ?? 0);
            ViewBag.TotalCost      = animals.Sum(a => a.SaleCost);

            return new ViewAsPdf("TallyPrint", animals)
            {
                FileName        = $"KilledAnimals_{date:yyyyMMdd}.pdf",
                PageSize        = Rotativa.AspNetCore.Options.Size.Letter,
                PageOrientation = Rotativa.AspNetCore.Options.Orientation.Landscape,
                PageMargins     = new Rotativa.AspNetCore.Options.Margins(8, 8, 8, 8),
                CustomSwitches  = "--print-media-type"
            };
        }

        // EXPORT EXCEL — PAGE 2 (filter view) 
        public async Task<IActionResult> ExportFilterExcel(
            int? vendorId, string? status,
            DateTime? killDateFrom, DateTime? killDateTo,
            DateTime? purchDateFrom, DateTime? purchDateTo)
        {
            var filter = new ExportFilter
            {
                VendorId      = vendorId,
                Status        = string.IsNullOrEmpty(status) || status == "all" ? null : status,
                KillDateFrom  = killDateFrom,
                KillDateTo    = killDateTo,
                PurchDateFrom = purchDateFrom,
                PurchDateTo   = purchDateTo,
            };
            var animals = (await _animalService.GetFilteredAsync(filter)).ToList();
            var bytes   = BuildExcel(animals, "Filtered animals export");
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"FilteredAnimals_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
        }

        //  EXPORT PDF — PAGE 2 
        public async Task<IActionResult> ExportFilterPdf(
            int? vendorId, string? status,
            DateTime? killDateFrom, DateTime? killDateTo,
            DateTime? purchDateFrom, DateTime? purchDateTo)
        {
            var filter = new ExportFilter
            {
                VendorId      = vendorId,
                Status        = string.IsNullOrEmpty(status) || status == "all" ? null : status,
                KillDateFrom  = killDateFrom,
                KillDateTo    = killDateTo,
                PurchDateFrom = purchDateFrom,
                PurchDateTo   = purchDateTo,
            };
            var animals = (await _animalService.GetFilteredAsync(filter)).ToList();

            ViewBag.KillDate       = DateTime.Today;
            ViewBag.TotalCount     = animals.Count;
            ViewBag.TotalCondemned = animals.Count(a => a.IsCondemned);
            ViewBag.TotalPassed    = animals.Count(a => !a.IsCondemned);
            ViewBag.TotalLiveWt    = animals.Sum(a => a.LiveWeight);
            ViewBag.TotalHotWt     = animals.Sum(a => a.HotWeight ?? 0);
            ViewBag.TotalCost      = animals.Sum(a => a.SaleCost);

            return new ViewAsPdf("TallyPrint", animals)
            {
                FileName        = $"FilteredAnimals_{DateTime.Now:yyyyMMdd_HHmm}.pdf",
                PageSize        = Rotativa.AspNetCore.Options.Size.Letter,
                PageOrientation = Rotativa.AspNetCore.Options.Orientation.Landscape,
                PageMargins     = new Rotativa.AspNetCore.Options.Margins(8, 8, 8, 8),
                CustomSwitches  = "--print-media-type"
            };
        }

        //  SHARED EXCEL BUILDER 
        private static byte[] BuildExcel(List<Animal> animals, string sheetTitle)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Animals");

            var headers = new[]
            {
                "Control No", "Animal Type", "Tag Number One", "Tag Number Two",
                "Purchase Date", "Purchase Type", "Vendor",
                "Live Weight", "Live Rate", "Kill Date", "Hot Weight",
                "Grade", "H S", "Comments", "Animal Control Number",
                "Tag 3", "Office Use 2", "State", "Buyer",
                "Animal Type 2", "Vet Name", "Kill Status", "Condemned"
            };

            var navy = XLColor.FromHtml("#1e2f47");
            for (int i = 0; i < headers.Length; i++)
            {
                var c = ws.Cell(1, i + 1);
                c.Value = headers[i];
                c.Style.Font.Bold            = true;
                c.Style.Fill.BackgroundColor = navy;
                c.Style.Font.FontColor       = XLColor.White;
            }

            int row = 2;
            foreach (var a in animals)
            {
                var fill = a.IsCondemned
                    ? XLColor.FromHtml("#fee2e2")
                    : row % 2 == 0 ? XLColor.FromHtml("#f8fafc") : XLColor.White;

                void Set(int col, object? val)
                {
                    var cell = ws.Cell(row, col);
                    cell.Value = val?.ToString() ?? "";
                    cell.Style.Fill.BackgroundColor = fill;
                    cell.Style.Font.FontSize = 10;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Hair;
                }

                Set(1,  a.ControlNo);
                Set(2,  a.AnimalType);
                Set(3,  a.TagNumber1);
                Set(4,  a.TagNumber2 ?? "");
                Set(5,  a.PurchaseDate.ToString("MM/dd/yyyy"));
                Set(6,  a.PurchaseType);
                Set(7,  a.Vendor?.VendorName ?? "");
                Set(8,  a.LiveWeight);
                Set(9,  a.LiveRate);
                Set(10, a.KillDate.HasValue ? a.KillDate.Value.ToString("MM/dd/yyyy") : "");
                Set(11, a.HotWeight.HasValue ? a.HotWeight.Value : (object)"");
                Set(12, a.Grade ?? "");
                Set(13, a.HealthScore.HasValue ? a.HealthScore.Value : (object)"");
                Set(14, a.Comment ?? "");
                Set(15, a.AnimalControlNumber ?? "");
                Set(16, a.Tag3 ?? "");
                Set(17, a.OfficeUse2 ?? "");
                Set(18, a.State ?? "");
                Set(19, a.BuyerName ?? "");
                Set(20, a.AnimalType2 ?? "");
                Set(21, a.VetName ?? "");
                Set(22, a.KillStatus);
                Set(23, a.IsCondemned ? "Yes" : "No");
                row++;
            }

            // Totals row
            var totFill = XLColor.FromHtml("#e2e8f0");
            ws.Cell(row, 1).Value = "TOTAL";
            ws.Cell(row, 8).Value = animals.Sum(a => a.LiveWeight);
            ws.Cell(row, 11).Value = animals.Sum(a => a.HotWeight ?? 0);
            for (int c = 1; c <= headers.Length; c++)
            {
                ws.Cell(row, c).Style.Font.Bold = true;
                ws.Cell(row, c).Style.Fill.BackgroundColor = totFill;
            }

            ws.SheetView.FreezeRows(1);
            ws.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            return stream.ToArray();
        }
    }
}
