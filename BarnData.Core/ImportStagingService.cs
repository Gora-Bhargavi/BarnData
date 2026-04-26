using BarnData.Data;
using BarnData.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace BarnData.Core.Services
{
    public class ImportStagingService : IImportStagingService
    {
        private readonly BarnDataContext _db;
        public ImportStagingService(BarnDataContext db) => _db = db;

        public async Task<ImportStagingBatch> CreateBatchAsync(
            string batchType, string? userId, string? sourceFileName, string? headerJson)
        {
            // Auto-clear any existing Active batch of the same type for this user —
            // each user has at most one Active staging per type at a time.
            var existing = await _db.ImportStagingBatches
                .Where(b => b.BatchType == batchType
                         && b.Status == "Active"
                         && b.CreatedBy == userId)
                .ToListAsync();

            foreach (var b in existing)
            {
                b.Status    = "Cleared";
                b.ClearedAt = DateTime.Now;
            }

            var batch = new ImportStagingBatch
            {
                BatchType      = batchType,
                CreatedBy      = userId,
                SourceFileName = sourceFileName,
                HeaderJson     = headerJson,
                CreatedAt      = DateTime.Now,
                Status         = "Active"
            };
            _db.ImportStagingBatches.Add(batch);
            await _db.SaveChangesAsync();
            return batch;
        }

        public async Task AddRowsAsync(
            int batchId,
            IEnumerable<(int RowNum, string Status, string? StatusNote, string RowJson)> rows)
        {
            var list = rows.ToList();
            if (list.Count == 0) return;

            var entities = list.Select(r => new ImportStagingRow
            {
                BatchID    = batchId,
                RowNum     = r.RowNum,
                Status     = r.Status,
                StatusNote = r.StatusNote,
                RowJson    = r.RowJson,
                CreatedAt  = DateTime.Now,
            }).ToList();

            // Chunk into 500-row batches to keep SaveChanges parameter counts sensible.
            const int chunk = 500;
            for (int i = 0; i < entities.Count; i += chunk)
            {
                _db.ImportStagingRows.AddRange(entities.Skip(i).Take(chunk));
                await _db.SaveChangesAsync();
            }

            await RefreshBatchCountsAsync(batchId);
        }

        public async Task<ImportStagingBatch?> GetActiveBatchAsync(string batchType, string? userId)
        {
            return await _db.ImportStagingBatches
                .Where(b => b.BatchType == batchType
                         && b.Status    == "Active"
                         && b.CreatedBy == userId)
                .OrderByDescending(b => b.CreatedAt)
                .FirstOrDefaultAsync();
        }

        public async Task<ImportStagingBatch?> GetBatchAsync(int batchId)
            => await _db.ImportStagingBatches.FindAsync(batchId);

        public async Task<List<ImportStagingRow>> GetRowsAsync(int batchId, string? statusFilter = null)
        {
            var q = _db.ImportStagingRows.AsNoTracking().Where(r => r.BatchID == batchId);
            if (!string.IsNullOrEmpty(statusFilter))
                q = q.Where(r => r.Status == statusFilter);
            return await q.OrderBy(r => r.RowNum).ToListAsync();
        }

        public async Task<bool> UpdateRowAsync(long rowId, string status, string? statusNote, string rowJson)
        {
            var row = await _db.ImportStagingRows.FindAsync(rowId);
            if (row == null) return false;

            row.Status     = status;
            row.StatusNote = statusNote;
            row.RowJson    = rowJson;
            row.UpdatedAt  = DateTime.Now;
            await _db.SaveChangesAsync();

            await RefreshBatchCountsAsync(row.BatchID);
            return true;
        }

        public async Task<bool> DeleteRowAsync(long rowId)
        {
            var row = await _db.ImportStagingRows.FindAsync(rowId);
            if (row == null) return false;
            var batchId = row.BatchID;
            _db.ImportStagingRows.Remove(row);
            await _db.SaveChangesAsync();
            await RefreshBatchCountsAsync(batchId);
            return true;
        }

        public async Task ClearBatchAsync(int batchId)
        {
            var batch = await _db.ImportStagingBatches.FindAsync(batchId);
            if (batch == null) return;
            batch.Status    = "Cleared";
            batch.ClearedAt = DateTime.Now;
            await _db.SaveChangesAsync();
        }

        public async Task MarkLoadedAsync(int batchId)
        {
            var batch = await _db.ImportStagingBatches.FindAsync(batchId);
            if (batch == null) return;
            batch.Status   = "Loaded";
            batch.LoadedAt = DateTime.Now;
            await _db.SaveChangesAsync();
        }

        public async Task RefreshBatchCountsAsync(int batchId)
        {
            var batch = await _db.ImportStagingBatches.FindAsync(batchId);
            if (batch == null) return;

            // One aggregate query — faster than per-status count round-trips.
            var counts = await _db.ImportStagingRows
                .Where(r => r.BatchID == batchId)
                .GroupBy(r => r.Status)
                .Select(g => new { Status = g.Key, Count = g.Count() })
                .ToListAsync();

            batch.TotalRows      = counts.Sum(c => c.Count);
            batch.OkCount        = counts.FirstOrDefault(c => c.Status == "OK")?.Count        ?? 0;
            batch.DuplicateCount = counts.FirstOrDefault(c => c.Status == "Duplicate")?.Count ?? 0;
            batch.ErrorCount     = counts.FirstOrDefault(c => c.Status == "Error")?.Count     ?? 0;
            batch.FlaggedCount   = counts.FirstOrDefault(c => c.Status == "Flag")?.Count      ?? 0;

            await _db.SaveChangesAsync();
        }
    }
}
