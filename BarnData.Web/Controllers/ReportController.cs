using BarnData.Core.Services;
using BarnData.Web.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Rotativa.AspNetCore;
using Microsoft.Extensions.Hosting;

namespace BarnData.Web.Controllers
{
    public class ReportController : Controller
    {
        private readonly IAnimalService _animalService;
        private readonly IVendorService _vendorService;
        private readonly IWebHostEnvironment _environment;

        public ReportController(
            IAnimalService animalService,
            IVendorService vendorService,
            IWebHostEnvironment environment)
        {
            _animalService = animalService;
            _vendorService = vendorService;
            _environment = environment;
        }

        // ── TALLY — main report page ──────────────────────────────────────
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
            var pdfExecutable = Path.Combine(
                _environment.WebRootPath ?? string.Empty,
                "Rotativa",
                "wkhtmltopdf.exe");

            if (!System.IO.File.Exists(pdfExecutable))
            {
                TempData["ErrorMessage"] =
                    "PDF export is not ready yet. Add wkhtmltopdf.exe under wwwroot/Rotativa and try again.";
                return RedirectToAction(nameof(Tally), new
                {
                    killDate = date.ToString("yyyy-MM-dd"),
                    vendorId
                });
            }

            var vm = new TallyViewModel
            {
                Summary  = summary,
                KillDate = date,
                VendorId = vendorId,
                IsPrint  = true,
                VendorList = Enumerable.Empty<
                    Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>()
            };

            return new ViewAsPdf("TallyPrint", vm)
            {
                FileName    = $"Tally_{date:yyyyMMdd}.pdf",
                PageSize    = Rotativa.AspNetCore.Options.Size.Letter,
                PageOrientation = Rotativa.AspNetCore.Options.Orientation.Landscape,
                PageMargins = new Rotativa.AspNetCore.Options.Margins(10, 10, 10, 10),
                CustomSwitches = "--print-media-type"
            };
        }

        // ── EXCEL EXPORT ──────────────────────────────────────────────────
        public async Task<IActionResult> ExportExcel(DateTime? killDate, int? vendorId)
        {
            var date    = killDate ?? DateTime.Today;
            var summary = await _animalService.GetTallySummaryAsync(date, vendorId);

            using var workbook  = new XLWorkbook();
            var ws = workbook.Worksheets.Add($"Tally {date:MM-dd-yyyy}");

            // ── Styles ────────────────────────────────────────────────────
            var headerFill  = XLColor.FromHtml("#0f1b2d");
            var headerFont  = XLColor.White;
            var vendorFill  = XLColor.FromHtml("#1e2f47");
            var vendorFont  = XLColor.White;
            var totalFill   = XLColor.FromHtml("#e2e8f0");
            var altFill     = XLColor.FromHtml("#f8fafc");
            var borderColor = XLColor.FromHtml("#cbd5e1");

            // ── Title block ───────────────────────────────────────────────
            ws.Cell("A1").Value = "NIGHTLY TALLY REPORT — TRAX-IT SLAUGHTER";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 14;
            ws.Range("A1:N1").Merge();

            ws.Cell("A2").Value = $"Kill Date: {date:dddd, MMMM d, yyyy}";
            ws.Cell("A2").Style.Font.Italic = true;
            ws.Range("A2:N2").Merge();

            ws.Cell("A3").Value = $"Generated: {DateTime.Now:MM/dd/yyyy h:mm tt}";
            ws.Cell("A3").Style.Font.Italic = true;
            ws.Cell("A3").Style.Font.FontColor = XLColor.Gray;
            ws.Range("A3:N3").Merge();

            // ── Column headers ────────────────────────────────────────────
            var headers = new[]
            {
                "Control No.", "Vendor", "Tag 1", "Tag 2", "Tag 3",
                "Animal Type", "Program", "Purchase Date", "Kill Date",
                "Live Wt (lbs)", "Live Rate", "Hot Wt (lbs)", "Grade", "Health Score"
            };

            int headerRow = 5;
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(headerRow, i + 1);
                cell.Value = headers[i];
                cell.Style.Fill.BackgroundColor       = headerFill;
                cell.Style.Font.FontColor             = headerFont;
                cell.Style.Font.Bold                  = true;
                cell.Style.Font.FontSize              = 10;
                cell.Style.Alignment.Horizontal       = XLAlignmentHorizontalValues.Center;
                cell.Style.Border.OutsideBorder       = XLBorderStyleValues.Thin;
                cell.Style.Border.OutsideBorderColor  = borderColor;
            }

            // ── Data rows ─────────────────────────────────────────────────
            int row = headerRow + 1;
            int seq = 0;

            foreach (var group in summary.ByVendor)
            {
                // Vendor group header row
                var vendorRange = ws.Range(row, 1, row, headers.Length);
                vendorRange.Merge();
                vendorRange.Style.Fill.BackgroundColor = vendorFill;
                vendorRange.Style.Font.FontColor       = vendorFont;
                vendorRange.Style.Font.Bold            = true;
                vendorRange.Style.Font.FontSize        = 10;

                ws.Cell(row, 1).Value =
                    $"  {group.VendorName}  —  {group.Count} animals  " +
                    $"|  Live: {group.TotalLiveWeight:N1} lbs  " +
                    $"|  Hot: {group.TotalHotWeight:N1} lbs";
                row++;

                foreach (var a in group.Animals)
                {
                    seq++;
                    bool alt = seq % 2 == 0;

                    void SetCell(int col, object? val, bool isNum = false)
                    {
                        var c = ws.Cell(row, col);
                        if (val != null) c.Value = val.ToString();
                        c.Style.Fill.BackgroundColor      = alt ? altFill : XLColor.White;
                        c.Style.Border.OutsideBorder      = XLBorderStyleValues.Hair;
                        c.Style.Border.OutsideBorderColor = borderColor;
                        c.Style.Font.FontSize             = 10;
                        if (isNum)
                            c.Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Right;
                    }

                    SetCell(1,  a.ControlNo);
                    SetCell(2,  a.Vendor?.VendorName);
                    SetCell(3,  a.TagNumber1);
                    SetCell(4,  a.TagNumber2);
                    SetCell(5,  a.Tag3);
                    SetCell(6,  a.AnimalType);
                    SetCell(7,  a.ProgramCode);
                    SetCell(8,  a.PurchaseDate.ToString("MM/dd/yyyy"));
                    SetCell(9,  a.KillDate.ToString("MM/dd/yyyy"));
                    SetCell(10, a.LiveWeight.ToString("N1"),  true);
                    SetCell(11, a.LiveRate.ToString("N4"),    true);
                    SetCell(12, a.HotWeight?.ToString("N1"),  true);
                    SetCell(13, a.Grade);
                    SetCell(14, a.HealthScore.ToString(), true);

                    row++;
                }

                // Vendor subtotal row
                var subRange = ws.Range(row, 1, row, headers.Length);
                subRange.Style.Fill.BackgroundColor = totalFill;
                subRange.Style.Font.Bold            = true;
                subRange.Style.Font.FontSize        = 10;

                ws.Cell(row, 1).Value  = $"Subtotal — {group.VendorName}";
                ws.Range(row, 1, row, 9).Merge();
                ws.Cell(row, 10).Value = group.TotalLiveWeight.ToString("N1");
                ws.Cell(row, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Cell(row, 12).Value = group.TotalHotWeight.ToString("N1");
                ws.Cell(row, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                row += 2; // blank gap between vendor groups
            }

            // ── Grand total row ───────────────────────────────────────────
            var grandRange = ws.Range(row, 1, row, headers.Length);
            grandRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#0f1b2d");
            grandRange.Style.Font.FontColor       = XLColor.White;
            grandRange.Style.Font.Bold            = true;
            grandRange.Style.Font.FontSize        = 11;

            ws.Cell(row, 1).Value  = $"GRAND TOTAL — {summary.TotalAnimals} animals";
            ws.Range(row, 1, row, 9).Merge();
            ws.Cell(row, 10).Value = summary.TotalLiveWeight.ToString("N1");
            ws.Cell(row, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            ws.Cell(row, 10).Style.Font.FontColor       = XLColor.White;
            ws.Cell(row, 12).Value = summary.TotalHotWeight.ToString("N1");
            ws.Cell(row, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            ws.Cell(row, 12).Style.Font.FontColor       = XLColor.White;

            // ── Yield row ─────────────────────────────────────────────────
            row++;
            ws.Cell(row, 1).Value = $"Average yield: {summary.AverageYieldPct:N1}%";
            ws.Cell(row, 1).Style.Font.Italic   = true;
            ws.Cell(row, 1).Style.Font.FontSize = 10;
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.Gray;
            ws.Range(row, 1, row, headers.Length).Merge();

            // ── Auto-fit columns ──────────────────────────────────────────
            ws.Columns().AdjustToContents();
            ws.Column(2).Width = Math.Min(ws.Column(2).Width, 30); // cap vendor col

            // ── Freeze header rows ────────────────────────────────────────
            ws.SheetView.FreezeRows(headerRow);

            // ── Return file ───────────────────────────────────────────────
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Tally_{date:yyyyMMdd}.xlsx"
            );
        }
    }
}
