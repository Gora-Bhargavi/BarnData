using BarnData.Data.Entities;

namespace BarnData.Core.Services
{
    public interface IAnimalService
    {
        // ── Queries ────────────────────────────────────────────────────────
        Task<IEnumerable<Animal>> GetByKillDateAsync(DateTime killDate, int? vendorId = null);
        Task<Animal?> GetByControlNoAsync(int controlNo);

        // ── Validation ────────────────────────────────────────────────────
        Task<bool> IsTagDuplicateAsync(string tag1, DateTime killDate, int vendorId, int? excludeControlNo = null);
        bool IsWeightOutOfRange(decimal liveWeight);

        // ── Write operations ──────────────────────────────────────────────
        Task<(bool Success, string ErrorMessage)> CreateAsync(Animal animal);
        Task<(bool Success, string ErrorMessage)> UpdateAsync(Animal animal);
        Task<bool> DeleteAsync(int controlNo);

        // ── Tally ─────────────────────────────────────────────────────────
        Task<TallySummary> GetTallySummaryAsync(DateTime killDate, int? vendorId = null);
    }

    public class TallySummary
    {
        public DateTime KillDate { get; set; }
        public int TotalAnimals { get; set; }
        public decimal TotalLiveWeight { get; set; }
        public decimal TotalHotWeight { get; set; }
        public decimal AverageYieldPct { get; set; }
        public IEnumerable<VendorGroup> ByVendor { get; set; } = new List<VendorGroup>();
    }

    public class VendorGroup
    {
        public string VendorName { get; set; } = string.Empty;
        public int Count { get; set; }
        public decimal TotalLiveWeight { get; set; }
        public decimal TotalHotWeight { get; set; }
        public IEnumerable<Animal> Animals { get; set; } = new List<Animal>();
    }
}
