using BarnData.Data.Entities;

namespace BarnData.Core.Services
{
    public interface IAnimalService
    {
        Task<IEnumerable<Animal>> GetByKillDateAsync(DateTime killDate, int? vendorId = null);
        Task<IEnumerable<Animal>> GetPendingAsync(int? vendorId = null);
        Task<IEnumerable<Animal>> GetAllAsync(int? vendorId = null);
        // Multi-vendor overloads
        Task<IEnumerable<Animal>> GetPendingByVendorsAsync(IEnumerable<int> vendorIds);
        Task<IEnumerable<Animal>> GetAllByVendorsAsync(IEnumerable<int> vendorIds);
        Task<IEnumerable<Animal>> GetByKillDateByVendorsAsync(DateTime killDate, IEnumerable<int> vendorIds);
        // Tag-based lookup for ACN auto-match
        Task<IEnumerable<Animal>> GetByTagsAsync(IEnumerable<string> tags);
        Task<Animal?> GetByControlNoAsync(int controlNo);
        Task<IEnumerable<Animal>> GetByTagSuffixAsync(string suffix);
    Task<IEnumerable<Animal>> GetByTagPatternAsync(string pattern);
    Task<IEnumerable<Animal>> GetAllPendingAsync();
    Task<(IEnumerable<Animal> Items, int TotalCount)> GetPendingPagedAsync(int? vendorId, int page, int pageSize);
    Task<bool> IsTagDuplicateAsync(string tag1, int vendorId, int? excludeControlNo = null);

    // fast duplicate detection during Excel upload.
        Task<HashSet<(string Tag, int VendorId)>> GetAllTagVendorKeysAsync();
        bool IsWeightOutOfRange(decimal liveWeight);
        Task<(bool Success, string ErrorMessage)> CreateAsync(Animal animal);
        Task<(int Imported, int Skipped, List<string> Errors)> BulkImportAsync(IEnumerable<Animal> animals);
        Task<int> MarkKilledAsync(IEnumerable<int> controlNos, DateTime killDate);
        Task<int> MarkKilledWithDataAsync(IEnumerable<KillAnimalData> animalData, DateTime killDate);
        Task<int> SaveKillDataAsync(IEnumerable<KillAnimalData> animalData);
        Task<(bool Success, string ErrorMessage)> UpdateAsync(Animal animal);
        Task<bool> DeleteAsync(int controlNo);
        Task<TallySummary> GetTallySummaryAsync(DateTime killDate, int? vendorId = null);
        Task<IEnumerable<Animal>> GetFilteredAsync(ExportFilter filter);

        // Hot Weight bulk import: match by AnimalControlNumber, apply rules
        Task<(int Updated, int Failed, List<string> Errors)> BulkUpdateHotWeightAsync(
            IEnumerable<HotWeightUpdateData> updates);

        // Fetch animals by AnimalControlNumbers for the import preview
        Task<IEnumerable<Animal>> GetByAnimalControlNumbersAsync(IEnumerable<string> acns);
    }

    public class HotWeightUpdateData
    {
        public int      ControlNo     { get; set; }   // system PK — resolved during preview
        public string   ACN           { get; set; } = string.Empty;
        public decimal? HotWeight     { get; set; }   // null = do not update
        public string?  Grade         { get; set; }
        public int?     HealthScore   { get; set; }
        public bool     ForceOverwrite { get; set; } = false;
        public string   ImportedBy    { get; set; } = string.Empty;
        public string   ImportFile    { get; set; } = string.Empty;
    }

    //  Per-animal kill data 
    public class KillAnimalData
    {
        public int      ControlNo   { get; set; }
        public decimal? HotWeight   { get; set; }
        public string?  Grade       { get; set; }
        public int?     HealthScore { get; set; }
        public bool     IsCondemned { get; set; }

        public string? AnimalControlNumber {get; set;}

        public DateTime? KillDate { get; set; }

        public decimal? LiveWeight {get; set;}

        public string? State {get; set;}
        public string? VetName {get; set;}

        public string? OfficeUse2 {get; set;}

        public string? Comment {get; set;}
    }

    public class TallySummary
    {
        public DateTime KillDate         { get; set; }
        public int      TotalAnimals     { get; set; }
        public int      TotalCondemned   { get; set; }
        public int      TotalPassed      { get; set; }
        public decimal  TotalLiveWeight  { get; set; }
        public decimal  TotalHotWeight   { get; set; }
        public decimal  TotalSaleCost    { get; set; }
        public decimal  AverageYieldPct  { get; set; }
        public decimal  AverageDressRate { get; set; }
        public decimal  AverageCost      { get; set; }
        public IEnumerable<VendorGroup>    ByVendor { get; set; } = new List<VendorGroup>();
        public IEnumerable<TypeSummaryRow> ByType   { get; set; } = new List<TypeSummaryRow>();
    }

    public class VendorGroup
    {
        public string  VendorName      { get; set; } = string.Empty;
        public int     Count           { get; set; }
        public int     Condemned       { get; set; }
        public int     Passed          { get; set; }
        public decimal TotalLiveWeight { get; set; }
        public decimal TotalHotWeight  { get; set; }
        public decimal TotalSaleCost   { get; set; }
        public decimal AvgCost         { get; set; }
        public decimal YieldPct        { get; set; }
        public decimal DressRate       { get; set; }
        public IEnumerable<Animal> Animals { get; set; } = new List<Animal>();
    }

    public class TypeSummaryRow
    {
        public string  Category  { get; set; } = string.Empty;
        public int     Killed    { get; set; }
        public int     Condemned { get; set; }
        public int     Passed    { get; set; }
        public decimal DressedWt { get; set; }
        public decimal Cost      { get; set; }
        public decimal AvgCost   { get; set; }
    }

    

    public class ExportFilter
    {
        public int?      VendorId      { get; set; }
        public string?   Status        { get; set; }  // Pending / Killed / Flagged / (null = all)
        public DateTime? KillDateFrom  { get; set; }
        public DateTime? KillDateTo    { get; set; }
        public DateTime? PurchDateFrom { get; set; }
        public DateTime? PurchDateTo   { get; set; }
    }
}

