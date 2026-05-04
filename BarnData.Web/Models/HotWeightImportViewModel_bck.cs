namespace BarnData.Web.Models
{
    public class HotWeightImportViewModel
    {
        public string? FileName { get; set; }
        public List<HotWeightPreviewRow> AutoRows { get; set; } = new();
        public List<HotWeightPreviewRow> FlaggedRows { get; set; } = new();
        public List<HotWeightPreviewRow> DupRows { get; set; } = new();   // exact duplicates (same HW+Grade+HS+ACN already in DB)
        public int TotalInExcel { get; set; }
        public int Matched { get; set; }
        public int AlreadyHasData { get; set; }
        public List<string> Errors { get; set; } = new();
    }

    public class HotWeightPreviewRow
    {
        public string RowKey { get; set; } = string.Empty;
        // From system
        public int ControlNo { get; set; }
        public string AnimalControlNumber { get; set; } = string.Empty;
        public string? CurrentHotWeight { get; set; }
        public string? CurrentGrade { get; set; }
        public string? CurrentHealthScore { get; set; }

        // From Excel
        public decimal? Side1 { get; set; }
        public decimal? Side2 { get; set; }
        public decimal? NewHotWeight { get; set; }   // Side1 + Side2 if both valid
        public string? NewGrade { get; set; }
        public string? NewGrade2 { get; set; }       // Grade 2 from Hot Scale
        public int? NewHealthScore { get; set; }
        public decimal? FileLiveWeight { get; set; } // live weight from Hot Scale file (for display + weight picking)
        public string? FileLot { get; set; }         // Lot number (display only)
        public string? FileSex { get; set; }         // Sex F/B (display only)
        public string? FileType { get; set; }        // Type DCOW/BULL etc (display only)
        public string? FileOrigin { get; set; }      // Origin US/CA → stored in DB

        public string? FileBackTag { get; set; }      

        public string? FileTag1 { get; set; }

        public string? FileTag2 { get; set; }

        public string? FileProgram { get; set; }

        // Multi-match candidates (serialised for picker UI)
        public List<HwCandidate>? Candidates { get; set; }

        // Status
        public string Status { get; set; } = "OK";   // OK | Flag | Loaded | Dup
        public string FlagReason { get; set; } = string.Empty;
        public string? TrimComment { get; set; }
        // Match info
        public string MatchMethod { get; set; } = "ACN";
        public string? NewAnimalControlNumber { get; set; }
    }

    // Candidate animal for multi-match picker — all 24 DB fields
    public class HwCandidate
    {
        public int     ControlNo            { get; set; }
        public string  AnimalControlNumber  { get; set; } = "";
        public string  Tag1                 { get; set; } = "";
        public string  Tag2                 { get; set; } = "";
        public string  Tag3                 { get; set; } = "";
        public string  VendorName           { get; set; } = "";
        public decimal LiveWeight           { get; set; }
        public decimal WeightDiff           { get; set; }
        public string  AnimalType           { get; set; } = "";
        public string  AnimalType2          { get; set; } = "";
        public string  ProgramCode          { get; set; } = "";
        public string  PurchaseDate         { get; set; } = "";
        public string  PurchaseType         { get; set; } = "";
        public decimal LiveRate             { get; set; }
        public decimal? ConsignmentRate     { get; set; }
        public string  Grade                { get; set; } = "";
        public string  Grade2               { get; set; } = "";
        public int?    HealthScore          { get; set; }
        public string  Comment              { get; set; } = "";
        public string  State                { get; set; } = "";
        public string  BuyerName            { get; set; } = "";
        public string  VetName              { get; set; } = "";
        public string  Origin               { get; set; } = "";
        public string  KillStatus           { get; set; } = "";
        public string  CreatedAt            { get; set; } = "";
    }

    public class HotWeightImportResult
    {
        public int Updated { get; set; }
        public int Failed { get; set; }
        public int Flagged { get; set; }
        public List<string> Errors { get; set; } = new();
        public string ImportedBy { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public DateTime ImportedAt { get; set; } = DateTime.Now;
    }
}
