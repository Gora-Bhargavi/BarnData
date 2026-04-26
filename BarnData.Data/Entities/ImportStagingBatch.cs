using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace BarnData.Data.Entities
{
    // A persisted import staging session. Survives browser close / app restart.
    // Each upload = one batch. Per-row detail lives in ImportStagingRow.
    // Status lifecycle: Active → Loaded (pushed to animal table) OR Cleared (discarded).
    [Table("tbl_import_staging_batch")]
    public class ImportStagingBatch
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int BatchID { get; set; }

        // "Excel" or "HotWeight"
        [Required]
        [MaxLength(20)]
        public string BatchType { get; set; } = "Excel";

        [MaxLength(50)]
        public string? CreatedBy { get; set; }

        [MaxLength(260)]
        public string? SourceFileName { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime? LoadedAt { get; set; }
        public DateTime? ClearedAt { get; set; }

        // Active | Loaded | Cleared
        [Required]
        [MaxLength(20)]
        public string Status { get; set; } = "Active";

        public int TotalRows { get; set; }
        public int OkCount { get; set; }
        public int DuplicateCount { get; set; }
        public int ErrorCount { get; set; }
        public int FlaggedCount { get; set; }

        // Optional header-level JSON (original filename, headers, file-wide errors).
        // Kept small — per-row detail lives in ImportStagingRow.
        public string? HeaderJson { get; set; }
    }

    [Table("tbl_import_staging_row")]
    public class ImportStagingRow
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public long RowID { get; set; }

        public int BatchID { get; set; }

        // Excel row number (1-based)
        public int RowNum { get; set; }

        // OK | Duplicate | Error | Flag | Loaded
        [Required]
        [MaxLength(20)]
        public string Status { get; set; } = "OK";

        [MaxLength(500)]
        public string? StatusNote { get; set; }

        // Full row payload as JSON (one column so Excel + HW rows share one table).
        // Deserialized into ExcelPreviewRow / HotWeightPreviewRow on read.
        [Required]
        public string RowJson { get; set; } = "{}";

        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime? UpdatedAt { get; set; }

        [ForeignKey("BatchID")]
        public ImportStagingBatch? Batch { get; set; }
    }
}
