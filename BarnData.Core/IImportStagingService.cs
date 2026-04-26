using BarnData.Data.Entities;

namespace BarnData.Core.Services
{
    // Persistent import staging — replaces HttpContext.Session for Excel & HW imports.
    // Rows survive browser close, app restart, and session expiry.
    // A user sees only their own batches (filtered by CreatedBy cookie).
    public interface IImportStagingService
    {
        // Create a new Active batch for this user + type.
        // Any existing Active batch for the same user+type is auto-cleared first.
        Task<ImportStagingBatch> CreateBatchAsync(
            string batchType, string? userId, string? sourceFileName, string? headerJson);

        // Bulk-add rows to a batch. Chunks into groups of 500 to keep SQL Server
        // parameter counts within limits. Auto-refreshes batch counts.
        Task AddRowsAsync(
            int batchId,
            IEnumerable<(int RowNum, string Status, string? StatusNote, string RowJson)> rows);

        // The latest Active batch for this user + type, or null.
        Task<ImportStagingBatch?> GetActiveBatchAsync(string batchType, string? userId);

        Task<ImportStagingBatch?> GetBatchAsync(int batchId);

        // All rows for a batch, optionally filtered by status.
        Task<List<ImportStagingRow>> GetRowsAsync(int batchId, string? statusFilter = null);

        // Update a single row's payload and status (used when user fixes an error row).
        Task<bool> UpdateRowAsync(long rowId, string status, string? statusNote, string rowJson);

        Task<bool> DeleteRowAsync(long rowId);

        // Mark batch as Cleared (soft delete — rows remain for audit).
        Task ClearBatchAsync(int batchId);

        // Mark batch as Loaded — used after OK rows are pushed to the main animal table.
        Task MarkLoadedAsync(int batchId);

        // Recompute batch counts from the rows. Call after UpdateRow/DeleteRow.
        Task RefreshBatchCountsAsync(int batchId);
    }
}
