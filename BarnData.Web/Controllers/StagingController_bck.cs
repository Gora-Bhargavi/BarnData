using BarnData.Core.Services;
using Microsoft.AspNetCore.Mvc;

namespace BarnData.Web.Controllers
{
    // Dedicated controller for persistent Excel / Hot-Weight import staging.
    // Lives alongside ImportController without modifying it.
    //
    // Endpoints (all prefixed /Staging/):
    //   GET  /Staging/Excel              — current active Excel batch + rows
    //   GET  /Staging/HotWeight          — current active HW batch + rows
    //   POST /Staging/Clear              — soft-delete the batch
    //   POST /Staging/UpdateRow          — edit a single row (status/payload)
    //   POST /Staging/DeleteRow          — drop a single row from staging
    //
    // Persistence outlives HttpContext.Session: rows survive browser close, app restart,
    // and cookie expiry. A user sees only their own batches — ownership via the
    // "bd_user" cookie (swap for real auth later).
    [Route("Staging")]
    public class StagingController : Controller
    {
        private readonly IImportStagingService _staging;
        public StagingController(IImportStagingService staging) => _staging = staging;

        private const string UserCookieName = "bd_user";

        private string CurrentUserId()
        {
            var existing = Request.Cookies[UserCookieName];
            if (!string.IsNullOrWhiteSpace(existing)) return existing;
            var newId = "user-" + Guid.NewGuid().ToString("N").Substring(0, 12);
            Response.Cookies.Append(UserCookieName, newId, new CookieOptions
            {
                Expires     = DateTimeOffset.UtcNow.AddYears(1),
                HttpOnly    = true,
                SameSite    = SameSiteMode.Lax,
                IsEssential = true
            });
            return newId;
        }

        [HttpGet("Excel")]
        public async Task<IActionResult> Excel() => await LoadBatch("Excel");

        [HttpGet("HotWeight")]
        public async Task<IActionResult> HotWeight() => await LoadBatch("HotWeight");

        private async Task<IActionResult> LoadBatch(string type)
        {
            var user  = CurrentUserId();
            var batch = await _staging.GetActiveBatchAsync(type, user);
            if (batch == null)
                return Json(new { hasStaging = false });

            var rows = await _staging.GetRowsAsync(batch.BatchID);
            return Json(new
            {
                hasStaging     = true,
                batchId        = batch.BatchID,
                createdAt      = batch.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"),
                sourceFileName = batch.SourceFileName,
                counts = new
                {
                    total     = batch.TotalRows,
                    ok        = batch.OkCount,
                    duplicate = batch.DuplicateCount,
                    error     = batch.ErrorCount,
                    flagged   = batch.FlaggedCount
                },
                rows = rows.Select(r => new
                {
                    rowId      = r.RowID,
                    rowNum     = r.RowNum,
                    status     = r.Status,
                    statusNote = r.StatusNote,
                    payload    = r.RowJson
                })
            });
        }

        [HttpPost("Clear")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Clear([FromBody] StagingClearRequest req)
        {
            if (req == null || req.BatchId <= 0)
                return Json(new { success = false, message = "Missing batch id." });

            var batch = await _staging.GetBatchAsync(req.BatchId);
            if (batch == null)
                return Json(new { success = false, message = "Batch not found." });

            // Ownership check — a user can only clear their own batch.
            var me = CurrentUserId();
            if (!string.Equals(batch.CreatedBy, me, StringComparison.Ordinal))
                return Json(new { success = false, message = "Not authorized." });

            await _staging.ClearBatchAsync(req.BatchId);
            return Json(new { success = true, message = "Staging cleared." });
        }

        [HttpPost("UpdateRow")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> UpdateRow([FromBody] StagingUpdateRowRequest req)
        {
            if (req == null || req.RowId <= 0 || string.IsNullOrWhiteSpace(req.Status))
                return Json(new { success = false, message = "Invalid request." });

            var updated = await _staging.UpdateRowAsync(
                req.RowId, req.Status!, req.StatusNote, req.Payload ?? "{}");
            return Json(new { success = updated });
        }

        [HttpPost("DeleteRow")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteRow([FromBody] StagingDeleteRowRequest req)
        {
            if (req == null || req.RowId <= 0)
                return Json(new { success = false, message = "Invalid request." });
            var ok = await _staging.DeleteRowAsync(req.RowId);
            return Json(new { success = ok });
        }
    }

    public class StagingClearRequest      { public int  BatchId { get; set; } }
    public class StagingDeleteRowRequest  { public long RowId   { get; set; } }
    public class StagingUpdateRowRequest
    {
        public long    RowId      { get; set; }
        public string? Status     { get; set; }
        public string? StatusNote { get; set; }
        public string? Payload    { get; set; }
    }
}
