using BarnData.Data.Entities;
using Microsoft.AspNetCore.Http;

namespace BarnData.Core.Services
{
    // ============================================================================
    // StagingBridge
    //
    // Bridges the existing Session["ExcelPreview"] / Session["HWPreview"]
    // caches to the persistent staging tables.
    //
    // Design: session stays as the hot cache (fast same-session performance).
    // After every session write, this also serializes the payload into the
    // staging tables so it survives browser close. On GET, if session is empty,
    // ReadAsync falls back to staging and re-hydrates session.
    //
    // Pragmatic storage model: we keep ONE row per batch (RowNum=0) holding
    // the entire preview JSON blob. This matches how existing code already
    // stores preview state (single JSON blob in session). A future phase can
    // split it into per-row staging for richer audit, but this keeps the
    // surgery minimal and 100% back-compatible with the existing import flow.
    // ============================================================================
    public static class StagingBridge
    {
        private const string ExcelBatchType     = "Excel";
        private const string HotWeightBatchType = "HotWeight";
        private const string OkStatus           = "OK";

        private static string SessionKeyFor(string batchType) =>
            batchType.Equals(ExcelBatchType, StringComparison.OrdinalIgnoreCase)
                ? "ExcelPreview"
                : "HWPreview";

        // Write JSON to both session (hot) and staging tables (persistent).
        // Returns the new batch ID, or null if staging write failed
        // (session is always updated, so failure here is non-fatal).
        public static async Task<int?> WriteAsync(
            ISession session,
            IImportStagingService stagingService,
            string batchType,
            string userKey,
            string json,
            string? sourceFileName = null)
        {
            if (string.IsNullOrEmpty(json)) return null;

            // 1. Session hot cache - always update first so same-session
            //    behaviour and performance are identical to before.
            session.SetString(SessionKeyFor(batchType), json);

            // 2. Persistent staging. Wrap in try/catch so any staging
            //    issue cannot break the import flow.
            try
            {
                var batch = await stagingService.CreateBatchAsync(
                    batchType:      batchType,
                    userId:         userKey,
                    sourceFileName: sourceFileName,
                    headerJson:     null);

                var rows = new[]
                {
                    (RowNum: 0, Status: OkStatus, StatusNote: (string?)null, RowJson: json)
                };
                await stagingService.AddRowsAsync(batch.BatchID, rows);

                return batch.BatchID;
            }
            catch
            {
                return null;
            }
        }

        // Read from session first; fall back to staging if empty.
        public static async Task<string?> ReadAsync(
            ISession session,
            IImportStagingService stagingService,
            string batchType,
            string userKey)
        {
            var fromSession = session.GetString(SessionKeyFor(batchType));
            if (!string.IsNullOrEmpty(fromSession)) return fromSession;

            try
            {
                var batch = await stagingService.GetActiveBatchAsync(batchType, userKey);
                if (batch == null) return null;

                var rows = await stagingService.GetRowsAsync(batch.BatchID);
                var row = rows.FirstOrDefault();
                if (row == null || string.IsNullOrEmpty(row.RowJson)) return null;

                // Restore hot cache so subsequent reads in this session are fast
                session.SetString(SessionKeyFor(batchType), row.RowJson);
                return row.RowJson;
            }
            catch
            {
                return null;
            }
        }

        // Clear both session and staging for this user+type.
        public static async Task ClearAsync(
            ISession session,
            IImportStagingService stagingService,
            string batchType,
            string userKey)
        {
            session.Remove(SessionKeyFor(batchType));
            try
            {
                var batch = await stagingService.GetActiveBatchAsync(batchType, userKey);
                if (batch != null)
                {
                    await stagingService.ClearBatchAsync(batch.BatchID);
                }
            }
            catch
            {
                // Session is already cleared; staging failure is non-fatal.
            }
        }

        // Has a persisted staging batch for this user+type?
        // Used by preview views to show a "restored from saved batch" banner.
        public static async Task<ImportStagingBatch?> GetActiveBatchInfoAsync(
            IImportStagingService stagingService,
            string batchType,
            string userKey)
        {
            try
            {
                return await stagingService.GetActiveBatchAsync(batchType, userKey);
            }
            catch
            {
                return null;
            }
        }

        // Derive a stable user key from the HttpContext using the
        // "bd_user" cookie pattern already established in Phase 1.
        public static string GetUserKey(HttpContext ctx)
        {
            if (ctx.Request.Cookies.TryGetValue("bd_user", out var existing) &&
                !string.IsNullOrWhiteSpace(existing))
            {
                return existing;
            }

            var fresh = Guid.NewGuid().ToString("N");
            ctx.Response.Cookies.Append("bd_user", fresh,
                new CookieOptions
                {
                    Expires     = DateTimeOffset.UtcNow.AddYears(1),
                    HttpOnly    = false,
                    IsEssential = true,
                    SameSite    = SameSiteMode.Lax,
                });
            return fresh;
        }

        // Shared key for Hot Weight staging.
        // All teammates read/write the same staging row when this key is used,
        // so collaborative review (auto-pull, smart-merge fixes, joint Load) works
        // across users. Excel staging continues to use per-user keys (GetUserKey).
        public const string SharedHotWeightKey = "GLOBAL_HW";

        // ---------------------------------------------------------------------------
        // Shared-staging overloads (skip the per-user Session hot cache).
        //
        // For shared HW workflow we cannot use the Session-string layer, because
        // each user has their own HTTP session — and a user-local cache would
        // shadow updates made by other teammates. These overloads route reads
        // and writes directly through the staging tables.
        //
        // Per-user behaviour (Excel imports + legacy HW path) is unaffected.
        // ---------------------------------------------------------------------------

        public static async Task<int?> WriteSharedAsync(
            IImportStagingService stagingService,
            string batchType,
            string sharedKey,
            string json,
            string? sourceFileName = null)
        {
            if (string.IsNullOrEmpty(json)) return null;
            try
            {
                var batch = await stagingService.CreateBatchAsync(
                    batchType:      batchType,
                    userId:         sharedKey,
                    sourceFileName: sourceFileName,
                    headerJson:     null);

                var rows = new[]
                {
                    (RowNum: 0, Status: OkStatus, StatusNote: (string?)null, RowJson: json)
                };
                await stagingService.AddRowsAsync(batch.BatchID, rows);
                return batch.BatchID;
            }
            catch
            {
                return null;
            }
        }

        public static async Task<string?> ReadSharedAsync(
            IImportStagingService stagingService,
            string batchType,
            string sharedKey)
        {
            try
            {
                var batch = await stagingService.GetActiveBatchAsync(batchType, sharedKey);
                if (batch == null) return null;
                var rows = await stagingService.GetRowsAsync(batch.BatchID);
                var row = rows.FirstOrDefault();
                if (row == null || string.IsNullOrEmpty(row.RowJson)) return null;
                return row.RowJson;
            }
            catch
            {
                return null;
            }
        }

        public static async Task ClearSharedAsync(
            IImportStagingService stagingService,
            string batchType,
            string sharedKey)
        {
            try
            {
                var batch = await stagingService.GetActiveBatchAsync(batchType, sharedKey);
                if (batch != null)
                {
                    await stagingService.ClearBatchAsync(batch.BatchID);
                }
            }
            catch
            {
                // Non-fatal — caller can ignore.
            }
        }

        public static class Types
        {
            public const string Excel     = ExcelBatchType;
            public const string HotWeight = HotWeightBatchType;
        }
    }
}
