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

        // ── TALLY PAGE ────────────────────────────────────────────────────
        public async Task<IActionResult> Tally(DateTime? killDate, int? vendorId)
        {
            var date    = killDate ?? DateTime.Today;
            var vendors = await _vendorService.GetAllActiveAsync();
            var summary = await _animalService.GetTallySummaryAsync(date, vendorId);

            var vm = new TallyViewModel
            {
                Summary    = summary,
                KillDate   = date,
                VendorId   = vendorId,
                VendorList = vendors.Select(v =>
                    new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem(
                        v.VendorName, v.VendorID.ToString()))
            };
            return View(vm);
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

        // ── INTERIM TALLY EXCEL — killed today, blank post-kill columns ───
        public async Task<IActionResult> ExportInterimTally(DateTime? killDate)
        {
            var date    = killDate ?? DateTime.Today;
            var summary = await _animalService.GetTallySummaryAsync(date);

            using var wb = new XLWorkbook();

            // ── SHEET 1: Summary ──────────────────────────────────────────
            var s1 = wb.Worksheets.Add("Summary");
            var navy = XLColor.FromHtml("#0f1b2d");

            s1.Cell("A1").Value = $"NIGHTLY KILL SUMMARY — {date:MMMM d, yyyy}";
            s1.Cell("A1").Style.Font.Bold = true;
            s1.Cell("A1").Style.Font.FontSize = 14;
            s1.Range("A1:G1").Merge();

            s1.Cell("A2").Value = $"Generated: {DateTime.Now:MM/dd/yyyy h:mm tt}  |  Status: Interim (HotScale not yet connected)";
            s1.Cell("A2").Style.Font.Italic = true;
            s1.Cell("A2").Style.Font.FontColor = XLColor.Gray;
            s1.Range("A2:G2").Merge();

            var s1H = new[] { "Category", "Killed", "Condemned", "Passed", "Total Live Wt", "Total Sale Cost", "Avg Rate" };
            for (int i = 0; i < s1H.Length; i++)
            {
                var c = s1.Cell(4, i + 1);
                c.Value = s1H[i];
                c.Style.Font.Bold = c.Style.Font.Bold;
                c.Style.Font.Bold = true;
                c.Style.Fill.BackgroundColor = navy;
                c.Style.Font.FontColor = XLColor.White;
            }

            int r = 5;
            foreach (var row in summary.ByType.Where(t => t.Killed > 0))
            {
                s1.Cell(r, 1).Value = row.Category;
                s1.Cell(r, 2).Value = row.Killed;
                s1.Cell(r, 3).Value = row.Condemned;
                s1.Cell(r, 4).Value = row.Passed;
                if (row.DressedWt > 0) s1.Cell(r, 5).Value = row.DressedWt;
                else s1.Cell(r, 5).Value = "(no hot wt yet)";
                s1.Cell(r, 6).Value = row.Cost;
                if (row.AvgCost > 0) s1.Cell(r, 7).Value = row.AvgCost;
                r++;
            }

            // Grand total
            s1.Cell(r, 1).Value = "TOTAL";
            s1.Cell(r, 2).Value = summary.TotalAnimals;
            s1.Cell(r, 3).Value = summary.TotalCondemned;
            s1.Cell(r, 4).Value = summary.TotalPassed;
            s1.Cell(r, 5).Value = summary.TotalLiveWeight;
            s1.Cell(r, 6).Value = summary.TotalSaleCost;
            s1.Row(r).Style.Font.Bold = true;
            s1.Row(r).Style.Fill.BackgroundColor = XLColor.FromHtml("#e2e8f0");
            s1.Columns().AdjustToContents();

            // ── SHEET 2: Animal list with blank post-kill columns ──────────
            var s2 = wb.Worksheets.Add("Kill List");

            var headers = new[]
            {
                "Ctrl No.", "Vendor", "Tag 1", "Tag 2", "Animal Type", "Program",
                "Purchase Type", "Origin", "Live Wt", "Live Rate", "Sale Cost",
                "Kill Date",
                // Blank columns for scale ticket data
                "Hot Wt (from scale)", "Yield %", "Dress Rate",
                "Grade", "Grade 2", "Health Score",
                "Condemned", "Comment"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var c = s2.Cell(1, i + 1);
                c.Value = headers[i];
                c.Style.Font.Bold = true;
                c.Style.Font.FontSize = 10;
                // Blank post-kill columns in amber to show they need filling
                bool isPostKill = i >= 12 && i <= 17;
                c.Style.Fill.BackgroundColor = isPostKill
                    ? XLColor.FromHtml("#fef3c7")
                    : navy;
                c.Style.Font.FontColor = isPostKill ? XLColor.FromHtml("#92400e") : XLColor.White;
                c.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int dataRow = 2;
            foreach (var group in summary.ByVendor)
            {
                foreach (var a in group.Animals)
                {
                    bool isAlt = dataRow % 2 == 0;
                    var rowFill = a.IsCondemned
                        ? XLColor.FromHtml("#fee2e2")
                        : isAlt ? XLColor.FromHtml("#f8fafc") : XLColor.White;

                    void Set(int col, object? val, bool postKill = false)
                    {
                        var cell = s2.Cell(dataRow, col);
                        if (val != null) cell.Value = val.ToString();
                        cell.Style.Fill.BackgroundColor = postKill
                            ? XLColor.FromHtml("#fffbeb")
                            : rowFill;
                        cell.Style.Font.FontSize = 10;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Hair;
                    }

                    decimal saleCost = a.LiveWeight * a.LiveRate;

                    Set(1,  a.ControlNo);
                    Set(2,  group.VendorName);
                    Set(3,  a.TagNumber1);
                    Set(4,  a.TagNumber2 ?? "");
                    Set(5,  a.AnimalType);
                    Set(6,  a.ProgramCode);
                    Set(7,  a.PurchaseType);
                    Set(8,  a.Origin ?? "");
                    Set(9,  a.LiveWeight.ToString("N1"));
                    Set(10, a.LiveRate > 0 ? "$" + a.LiveRate.ToString("N4") : "");
                    Set(11, saleCost > 0 ? "$" + saleCost.ToString("N2") : "");
                    Set(12, a.KillDate.HasValue ? a.KillDate.Value.ToString("MM/dd/yyyy") : "");
                    // Post-kill columns — blank, amber background
                    Set(13, a.HotWeight.HasValue ? a.HotWeight.Value.ToString("N1") : "", true);
                    Set(14, "", true);  // Yield — calculated after hot weight entered
                    Set(15, "", true);  // Dress rate
                    Set(16, a.Grade ?? "", true);
                    Set(17, a.Grade2 ?? "", true);
                    Set(18, a.HealthScore.HasValue ? a.HealthScore.Value.ToString() : "", true);
                    Set(19, a.IsCondemned ? "COND" : "");
                    Set(20, a.Comment ?? "");

                    dataRow++;
                }

                // Vendor subtotal
                s2.Cell(dataRow, 1).Value = $"Subtotal — {group.VendorName}  |  {group.Count} animals, {group.Condemned} condemned";
                s2.Range(dataRow, 1, dataRow, 20).Merge();
                s2.Row(dataRow).Style.Fill.BackgroundColor = XLColor.FromHtml("#1e2f47");
                s2.Row(dataRow).Style.Font.FontColor = XLColor.White;
                s2.Row(dataRow).Style.Font.Bold = true;
                dataRow += 2;
            }

            s2.SheetView.FreezeRows(1);
            s2.Columns().AdjustToContents();
            s2.Column(2).Width = Math.Min(s2.Column(2).Width, 28);

            // ── SHEET 3: Instructions ──────────────────────────────────────
            var s3 = wb.Worksheets.Add("Instructions");
            s3.Cell("A1").Value = "How to complete this tally sheet";
            s3.Cell("A1").Style.Font.Bold = true;
            s3.Cell("A1").Style.Font.FontSize = 13;

            var steps = new[]
            {
                "1. Go to the Kill List tab.",
                "2. Columns highlighted in YELLOW must be filled from the scale tickets: Hot Wt, Yield, Dress Rate, Grade, Grade 2, Health Score.",
                "3. For each animal, enter the Hot Weight from the scale ticket.",
                "4. Yield % = Hot Wt ÷ Live Wt × 100  (or 100% for consignment animals).",
                "5. Dress Rate = Sale Cost ÷ Hot Wt.",
                "6. Enter Grade (CT, B1, B2, CN, LB, BB etc.) from scale ticket.",
                "7. Enter Health Score (1–5) from scale ticket.",
                "8. If an animal is condemned, column S already shows COND from the import.",
                "9. When HotScale is connected (Phase 5), all yellow columns will auto-fill. This manual step will be eliminated.",
            };
            for (int i = 0; i < steps.Length; i++)
            {
                s3.Cell(i + 3, 1).Value = steps[i];
                s3.Cell(i + 3, 1).Style.Font.FontSize = 11;
            }
            s3.Column(1).Width = 90;

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            stream.Position = 0;

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"KillList_{date:yyyyMMdd}.xlsx"
            );
        }
    }
}
