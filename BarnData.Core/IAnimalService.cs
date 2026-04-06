using BarnData.Data.Entities;

namespace BarnData.Core.Services
{
    public interface IAnimalService
    {
        Task<IEnumerable<Animal>> GetByKillDateAsync(DateTime killDate, int? vendorId = null);
        Task<Animal?> GetByControlNoAsync(int controlNo);
        Task<bool> IsTagDuplicateAsync(string tag1, DateTime killDate, int vendorId, int? excludeControlNo = null);
        bool IsWeightOutOfRange(decimal liveWeight);
        Task<(bool Success, string ErrorMessage)> CreateAsync(Animal animal);
        Task<(bool Success, string ErrorMessage)> UpdateAsync(Animal animal);
        Task<bool> DeleteAsync(int controlNo);
        Task<TallySummary> GetTallySummaryAsync(DateTime killDate, int? vendorId = null);
        Task<TallySummary> GetTodayKilledSummaryAsync();
        Task<IEnumerable<Animal>> SearchAnimalsByVendorNameAsync(string vendorName);
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
}
