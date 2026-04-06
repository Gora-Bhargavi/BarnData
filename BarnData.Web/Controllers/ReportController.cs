using BarnData.Core.Services;
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

        // ── Legacy route kept for compatibility ───────────────────────────
        public IActionResult Tally(DateTime? killDate, int? vendorId)
        {
            return RedirectToAction(nameof(TallyToday));
        }

        // ── Today tally page ──────────────────────────────────────────────
        public async Task<IActionResult> TallyToday()
        {
            var summary = await _animalService.GetTodayKilledSummaryAsync();

            var vm = new TallyViewModel
            {
                Summary  = summary,
                KillDate = DateTime.Today
            };

            return View(vm);
        }

        // ── Vendor search page ────────────────────────────────────────────
        public async Task<IActionResult> VendorAnimals(string? vendorName)
        {
            var vm = new VendorAnimalsViewModel
            {
                VendorName = vendorName ?? string.Empty,
                Animals = string.IsNullOrWhiteSpace(vendorName)
                    ? Enumerable.Empty<BarnData.Data.Entities.Animal>()
                    : await _animalService.SearchAnimalsByVendorNameAsync(vendorName)
            };

            return View(vm);
        }

        public Task<IActionResult> ExportTodayPdf()
        {
            return ExportPdf(DateTime.Today, null);
        }

        public Task<IActionResult> ExportTodayExcel()
        {
            return ExportExcel(DateTime.Today, null);
        }

        public async Task<IActionResult> ExportVendorAnimalsPdf(string? vendorName)
        {
            var normalizedVendorName = vendorName ?? string.Empty;
            var vm = new VendorAnimalsViewModel
            {
                VendorName = normalizedVendorName,
                Animals = string.IsNullOrWhiteSpace(normalizedVendorName)
                    ? Enumerable.Empty<BarnData.Data.Entities.Animal>()
                    : await _animalService.SearchAnimalsByVendorNameAsync(normalizedVendorName)
            };

            return new ViewAsPdf("VendorAnimalsPrint", vm)
            {
                FileName = $"VendorAnimals_{SanitizeFileName(normalizedVendorName)}_{DateTime.Today:yyyyMMdd}.pdf",
                PageSize = Rotativa.AspNetCore.Options.Size.Letter,
                PageOrientation = Rotativa.AspNetCore.Options.Orientation.Landscape,
                PageMargins = new Rotativa.AspNetCore.Options.Margins(8, 8, 8, 8),
                CustomSwitches = "--print-media-type"
            };
        }

        public async Task<IActionResult> ExportVendorAnimalsExcel(string? vendorName)
        {
            var normalizedVendorName = vendorName ?? string.Empty;
            var animals = string.IsNullOrWhiteSpace(normalizedVendorName)
                ? Enumerable.Empty<BarnData.Data.Entities.Animal>()
                : await _animalService.SearchAnimalsByVendorNameAsync(normalizedVendorName);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Vendor Animals");

            ws.Cell("A1").Value = $"Vendor Animal Search: {normalizedVendorName}";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 13;
            ws.Range("A1:L1").Merge();

            ws.Cell("A2").Value = $"Generated: {DateTime.Now:MM/dd/yyyy h:mm tt}";
            ws.Range("A2:L2").Merge();

            var headers = new[]
            {
                "Control No.", "Vendor", "Tag 1", "Tag 2", "Type", "Program",
                "Kill Date", "Status", "Grade", "Live Wt", "Hot Wt", "Comment"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(4, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#0f1b2d");
                cell.Style.Font.FontColor = XLColor.White;
            }

            int row = 5;
            foreach (var animal in animals)
            {
                ws.Cell(row, 1).Value = animal.ControlNo;
                ws.Cell(row, 2).Value = animal.Vendor?.VendorName ?? string.Empty;
                ws.Cell(row, 3).Value = animal.TagNumber1;
                ws.Cell(row, 4).Value = animal.TagNumber2 ?? string.Empty;
                ws.Cell(row, 5).Value = animal.AnimalType;
                ws.Cell(row, 6).Value = animal.ProgramCode;
                ws.Cell(row, 7).Value = animal.KillDate.ToString("MM/dd/yyyy");
                ws.Cell(row, 8).Value = animal.KillStatus;
                ws.Cell(row, 9).Value = animal.Grade;
                ws.Cell(row, 10).Value = animal.LiveWeight;
                ws.Cell(row, 11).Value = animal.HotWeight?.ToString("N1") ?? string.Empty;
                ws.Cell(row, 12).Value = animal.Comment ?? string.Empty;
                row++;
            }

            ws.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            stream.Position = 0;

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"VendorAnimals_{SanitizeFileName(normalizedVendorName)}_{DateTime.Today:yyyyMMdd}.xlsx"
            );
        }

        // ── PDF EXPORT ────────────────────────────────────────────────────
        public async Task<IActionResult> ExportPdf(DateTime? killDate, int? vendorId)
        {
            var date    = killDate ?? DateTime.Today;
            var summary = await _animalService.GetTallySummaryAsync(date, vendorId);

            var vm = new TallyViewModel
            {
                Summary    = summary,
                KillDate   = date,
                VendorId   = vendorId,
                IsPrint    = true,
                VendorList = Enumerable.Empty<
                    Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>()
            };

            return new ViewAsPdf("TallyPrint", vm)
            {
                FileName        = $"Tally_{date:yyyyMMdd}.pdf",
                PageSize        = Rotativa.AspNetCore.Options.Size.Letter,
                PageOrientation = Rotativa.AspNetCore.Options.Orientation.Landscape,
                PageMargins     = new Rotativa.AspNetCore.Options.Margins(8, 8, 8, 8),
                CustomSwitches  = "--print-media-type"
            };
        }

        // ── EXCEL EXPORT — matches real tally report format ───────────────
        public async Task<IActionResult> ExportExcel(DateTime? killDate, int? vendorId)
        {
            var date    = killDate ?? DateTime.Today;
            var summary = await _animalService.GetTallySummaryAsync(date, vendorId);

            using var wb = new XLWorkbook();

            // ── SHEET 1: Summary by animal type ───────────────────────────
            var s1 = wb.Worksheets.Add("Summary");

            // Title
            s1.Cell("A1").Value = $"{date:M.d.yyyy} KILL";
            s1.Cell("A1").Style.Font.Bold     = true;
            s1.Cell("A1").Style.Font.FontSize = 13;
            s1.Range("A1:G1").Merge();

            // Headers
            var s1Headers = new[] { "", "Killed", "Cond", "Passed", "Dressed Wt", "Cost", "Avg Cost" };
            for (int i = 0; i < s1Headers.Length; i++)
            {
                var c = s1.Cell(2, i + 1);
                c.Value = s1Headers[i];
                c.Style.Font.Bold = true;
                c.Style.Fill.BackgroundColor = XLColor.FromHtml("#0f1b2d");
                c.Style.Font.FontColor       = XLColor.White;
                c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int s1Row = 3;
            foreach (var row in summary.ByType)
            {
                s1.Cell(s1Row, 1).Value = row.Category;
                s1.Cell(s1Row, 2).Value = row.Killed;
                s1.Cell(s1Row, 3).Value = row.Condemned > 0 ? row.Condemned : 0;
                s1.Cell(s1Row, 4).Value = row.Passed;
                s1.Cell(s1Row, 5).Value = row.DressedWt;
                s1.Cell(s1Row, 6).Value = row.Cost;
                s1.Cell(s1Row, 7).Value = row.AvgCost;
                if (row.Killed == 0)
                    s1.Row(s1Row).Style.Font.FontColor = XLColor.Gray;
                s1Row++;
            }

            // Totals row
            var totRow = s1.Row(s1Row);
            s1.Cell(s1Row, 1).Value = "TOTAL";
            s1.Cell(s1Row, 2).Value = summary.TotalAnimals;
            s1.Cell(s1Row, 3).Value = summary.TotalCondemned;
            s1.Cell(s1Row, 4).Value = summary.TotalPassed;
            s1.Cell(s1Row, 5).Value = summary.TotalHotWeight;
            s1.Cell(s1Row, 6).Value = summary.TotalSaleCost;
            s1.Cell(s1Row, 7).Value = summary.AverageDressRate;
            totRow.Style.Font.Bold = true;
            totRow.Style.Fill.BackgroundColor = XLColor.FromHtml("#e2e8f0");

            s1.Columns().AdjustToContents();

            // ── SHEET 2: Data dump (one row per animal) ────────────────────
            var s2 = wb.Worksheets.Add("Data Dump");

            var headers = new[]
            {
                "Vendor", "Control No.", "Purchase Date", "Purchase Type",
                "Tag 1", "Tag 2", "Animal Type", "Program",
                "Origin", "Live Wt", "Live Rate", "Sale Cost",
                "Hot Wt", "Yield %", "Dress Rate",
                "Grade", "Grade 2", "Health Score", "Comment", "Condemned"
            };

            var hFill = XLColor.FromHtml("#0f1b2d");
            for (int i = 0; i < headers.Length; i++)
            {
                var c = s2.Cell(1, i + 1);
                c.Value = headers[i];
                c.Style.Font.Bold     = true;
                c.Style.Font.FontSize = 10;
                c.Style.Fill.BackgroundColor = hFill;
                c.Style.Font.FontColor       = XLColor.White;
                c.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int dataRow = 2;
            foreach (var group in summary.ByVendor)
            {
                foreach (var a in group.Animals)
                {
                    decimal saleCost  = a.LiveWeight * a.LiveRate;
                    decimal yieldPct  = (a.HotWeight.HasValue && a.LiveWeight > 0)
                        ? Math.Round(a.HotWeight.Value / a.LiveWeight * 100, 2) : 0;
                    decimal dressRate = (a.HotWeight.HasValue && a.HotWeight.Value > 0)
                        ? Math.Round(saleCost / a.HotWeight.Value, 3) : 0;

                    bool isAlt = dataRow % 2 == 0;
                    var rowFill = a.IsCondemned
                        ? XLColor.FromHtml("#fee2e2")
                        : isAlt ? XLColor.FromHtml("#f8fafc") : XLColor.White;

                    void Set(int col, object? val)
                    {
                        var cell = s2.Cell(dataRow, col);
                        if (val != null) cell.Value = val.ToString();
                        cell.Style.Fill.BackgroundColor = rowFill;
                        cell.Style.Font.FontSize        = 10;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Hair;
                    }

                    Set(1,  group.VendorName);
                    Set(2,  a.ControlNo.ToString());
                    Set(3,  a.PurchaseDate.ToString("MM/dd/yyyy"));
                    Set(4,  a.PurchaseType);
                    Set(5,  a.TagNumber1);
                    Set(6,  a.TagNumber2);
                    Set(7,  a.AnimalType);
                    Set(8,  a.ProgramCode);
                    Set(9,  a.Origin ?? "");
                    Set(10, a.LiveWeight.ToString("N1"));
                    Set(11, a.LiveRate.ToString("N4"));
                    Set(12, saleCost.ToString("N2"));
                    Set(13, a.HotWeight.HasValue ? a.HotWeight.Value.ToString("N1") : "");
                    Set(14, yieldPct > 0 ? yieldPct.ToString("N2") + "%" : "");
                    Set(15, dressRate > 0 ? dressRate.ToString("N3") : "");
                    Set(16, a.Grade);
                    Set(17, a.Grade2 ?? "");
                    Set(18, a.HealthScore.ToString());
                    Set(19, a.Comment);
                    Set(20, a.IsCondemned ? "cond" : "");

                    dataRow++;
                }

                // Vendor subtotal row
                s2.Cell(dataRow, 1).Value  = $"Subtotal — {group.VendorName}";
                s2.Range(dataRow, 1, dataRow, 9).Merge();
                s2.Cell(dataRow, 10).Value = group.TotalLiveWeight.ToString("N1");
                s2.Cell(dataRow, 12).Value = group.TotalSaleCost.ToString("N2");
                s2.Cell(dataRow, 13).Value = group.TotalHotWeight.ToString("N1");
                s2.Cell(dataRow, 14).Value = group.YieldPct.ToString("N1") + "%";
                s2.Cell(dataRow, 15).Value = group.DressRate.ToString("N3");
                s2.Row(dataRow).Style.Font.Bold = true;
                s2.Row(dataRow).Style.Fill.BackgroundColor = XLColor.FromHtml("#1e2f47");
                s2.Row(dataRow).Style.Font.FontColor       = XLColor.White;
                dataRow += 2;
            }

            // Grand total
            s2.Cell(dataRow, 1).Value  = $"GRAND TOTAL — {summary.TotalAnimals} killed, {summary.TotalCondemned} condemned, {summary.TotalPassed} passed";
            s2.Range(dataRow, 1, dataRow, 9).Merge();
            s2.Cell(dataRow, 10).Value = summary.TotalLiveWeight.ToString("N1");
            s2.Cell(dataRow, 12).Value = summary.TotalSaleCost.ToString("N2");
            s2.Cell(dataRow, 13).Value = summary.TotalHotWeight.ToString("N1");
            s2.Cell(dataRow, 14).Value = summary.AverageYieldPct.ToString("N1") + "%";
            s2.Cell(dataRow, 15).Value = summary.AverageDressRate.ToString("N3");
            s2.Row(dataRow).Style.Font.Bold     = true;
            s2.Row(dataRow).Style.Font.FontSize = 11;
            s2.Row(dataRow).Style.Fill.BackgroundColor = XLColor.FromHtml("#0f1b2d");
            s2.Row(dataRow).Style.Font.FontColor       = XLColor.White;

            s2.SheetView.FreezeRows(1);
            s2.Columns().AdjustToContents();
            s2.Column(1).Width = Math.Min(s2.Column(1).Width, 30);

            // ── SHEET 3: Grade breakdown placeholder ──────────────────────
            var s3 = wb.Worksheets.Add("Grade Breakdown");
            s3.Cell("A1").Value = "Grade breakdown — available after HotScale integration (Phase 5)";
            s3.Cell("A1").Style.Font.Italic    = true;
            s3.Cell("A1").Style.Font.FontColor = XLColor.Gray;

            // Grade count headers
            var gradeHeaders = new[] { "", "CT", "CN", "B1", "B2", "BR", "Totals" };
            for (int i = 0; i < gradeHeaders.Length; i++)
            {
                var c = s3.Cell(3, i + 1);
                c.Value = gradeHeaders[i];
                c.Style.Font.Bold = true;
                c.Style.Fill.BackgroundColor = hFill;
                c.Style.Font.FontColor       = XLColor.White;
            }
            var categories = new[]
            {
                "Sale Cows","Sale Bulls","Consignment Cows","Consignment Bulls",
                "Canadian Cows","Canadian Bulls"
            };
            var grades = new[] { "CT","CN","B1","B2","BR" };
            int g3Row = 4;
            foreach (var cat in categories)
            {
                s3.Cell(g3Row, 1).Value = cat;
                // Count grades from actual data
                var animals = summary.ByVendor.SelectMany(v => v.Animals)
                    .Where(a => !a.IsCondemned).ToList();
                for (int gi = 0; gi < grades.Length; gi++)
                {
                    int count = animals.Count(a =>
                        a.Grade.Trim().Equals(grades[gi], StringComparison.OrdinalIgnoreCase));
                    s3.Cell(g3Row, gi + 2).Value = count > 0 ? count : 0;
                }
                g3Row++;
            }
            s3.Columns().AdjustToContents();

            // Return file
            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            stream.Position = 0;

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Tally_{date:yyyyMMdd}.xlsx"
            );
        }

        private static string SanitizeFileName(string value)
        {
            var fallback = string.IsNullOrWhiteSpace(value) ? "AllVendors" : value.Trim();
            foreach (var invalidChar in Path.GetInvalidFileNameChars())
            {
                fallback = fallback.Replace(invalidChar, '_');
            }

            return fallback.Replace(' ', '_');
        }
    }
}
