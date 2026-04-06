using BarnData.Data;
using BarnData.Data.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

namespace BarnData.Core.Services
{
    public class AnimalService : IAnimalService
    {
        private readonly BarnDataContext _db;
        private readonly decimal _weightMin;
        private readonly decimal _weightMax;
        private const string PendingStatus = "Pending";
        private const string KilledStatus = "Killed";
        private const string VerifiedStatus = "Verified";
        private const string FlaggedStatus = "Flagged";

        public AnimalService(BarnDataContext db, IConfiguration config)
        {
            _db = db;
            _weightMin = config.GetValue<decimal>("AppSettings:LiveWeightMinLbs", 300);
            _weightMax = config.GetValue<decimal>("AppSettings:LiveWeightMaxLbs", 2500);
        }

        // ── Get animals by kill date ──────────────────────────────────────
        public async Task<IEnumerable<Animal>> GetByKillDateAsync(
            DateTime killDate, int? vendorId = null)
        {
            var query = _db.Animals
                .Include(a => a.Vendor)
                .Where(a => a.KillDate.Date == killDate.Date);

            if (vendorId.HasValue)
                query = query.Where(a => a.VendorID == vendorId.Value);

            var animals = await query
                .OrderBy(a => a.Vendor!.VendorName)
                .ThenBy(a => a.ControlNo)
                .ToListAsync();

            var statusWasUpdated = false;
            foreach (var animal in animals)
            {
                if (SyncKillStatus(animal))
                {
                    statusWasUpdated = true;
                }
            }

            if (statusWasUpdated)
            {
                await _db.SaveChangesAsync();
            }

            return animals;
        }

        // ── Get single animal ─────────────────────────────────────────────
        public async Task<Animal?> GetByControlNoAsync(int controlNo)
        {
            return await _db.Animals
                .Include(a => a.Vendor)
                .FirstOrDefaultAsync(a => a.ControlNo == controlNo);
        }

        // ── Duplicate tag check ───────────────────────────────────────────
        public async Task<bool> IsTagDuplicateAsync(
            string tag1, DateTime killDate, int vendorId, int? excludeControlNo = null)
        {
            var query = _db.Animals.Where(a =>
                a.TagNumber1 == tag1.Trim() &&
                a.KillDate.Date == killDate.Date &&
                a.VendorID == vendorId);

            if (excludeControlNo.HasValue)
                query = query.Where(a => a.ControlNo != excludeControlNo.Value);

            return await query.AnyAsync();
        }

        // ── Weight range check ────────────────────────────────────────────
        public bool IsWeightOutOfRange(decimal liveWeight)
            => liveWeight < _weightMin || liveWeight > _weightMax;

        // ── Create ────────────────────────────────────────────────────────
        public async Task<(bool Success, string ErrorMessage)> CreateAsync(Animal animal)
        {
            bool isDuplicate = await IsTagDuplicateAsync(
                animal.TagNumber1, animal.KillDate, animal.VendorID);

            if (isDuplicate)
                return (false,
                    $"Tag '{animal.TagNumber1}' already exists for this vendor on {animal.KillDate:MM/dd/yyyy}.");

            animal.TagNumber1 = animal.TagNumber1.Trim();
            animal.TagNumber2 = (animal.TagNumber2 ?? string.Empty).Trim();
            animal.Comment = animal.Comment ?? string.Empty;
            animal.CreatedAt  = DateTime.Now;
            animal.KillStatus = GetCurrentKillStatus(animal.KillDate);

            _db.Animals.Add(animal);
            await _db.SaveChangesAsync();
            return (true, string.Empty);
        }

        // ── Update ────────────────────────────────────────────────────────
        public async Task<(bool Success, string ErrorMessage)> UpdateAsync(Animal animal)
        {
            bool isDuplicate = await IsTagDuplicateAsync(
                animal.TagNumber1, animal.KillDate, animal.VendorID, animal.ControlNo);

            if (isDuplicate)
                return (false,
                    $"Tag '{animal.TagNumber1}' already exists for this vendor on {animal.KillDate:MM/dd/yyyy}.");

            var existing = await _db.Animals.FindAsync(animal.ControlNo);
            if (existing == null) return (false, "Animal record not found.");

            existing.VendorID            = animal.VendorID;
            existing.TagNumber1          = animal.TagNumber1.Trim();
            existing.TagNumber2          = (animal.TagNumber2 ?? string.Empty).Trim();
            existing.Tag3                = animal.Tag3;
            existing.AnimalType          = animal.AnimalType;
            existing.AnimalType2         = animal.AnimalType2;
            existing.ProgramCode         = animal.ProgramCode;
            existing.PurchaseDate        = animal.PurchaseDate;
            existing.PurchaseType        = animal.PurchaseType;
            existing.LiveWeight          = animal.LiveWeight;
            existing.LiveRate            = animal.LiveRate;
            existing.KillDate            = animal.KillDate;
            existing.HotWeight           = animal.HotWeight;
            existing.Grade               = animal.Grade;
            existing.Grade2              = animal.Grade2;
            existing.HealthScore         = animal.HealthScore;
            existing.FetalBlood          = animal.FetalBlood;
            existing.Comment             = animal.Comment ?? string.Empty;
            existing.AnimalControlNumber = animal.AnimalControlNumber;
            existing.State               = animal.State;
            existing.BuyerName           = animal.BuyerName;
            existing.VetName             = animal.VetName;
            existing.OfficeUse2          = animal.OfficeUse2;
            existing.Origin              = animal.Origin;
            existing.IsCondemned         = animal.IsCondemned;
            existing.UpdatedAt           = DateTime.Now;
            SyncKillStatus(existing);

            await _db.SaveChangesAsync();
            return (true, string.Empty);
        }

        // ── Soft delete ───────────────────────────────────────────────────
        public async Task<bool> DeleteAsync(int controlNo)
        {
            var animal = await _db.Animals.FindAsync(controlNo);
            if (animal == null) return false;

            animal.KillStatus = FlaggedStatus;
            animal.UpdatedAt  = DateTime.Now;
            await _db.SaveChangesAsync();
            return true;
        }

        // ── Tally summary — full version with calculated columns ──────────
        public async Task<TallySummary> GetTallySummaryAsync(
            DateTime killDate, int? vendorId = null)
        {
            var all = (await GetByKillDateAsync(killDate, vendorId))
                .Where(a => a.KillStatus != FlaggedStatus)
                .ToList();

            // Helper: Sale Cost per animal
            static decimal SaleCost(Animal a)
                => a.LiveWeight * a.LiveRate;

            // Helper: Yield %
            static decimal YieldPct(Animal a)
                => (a.HotWeight.HasValue && a.LiveWeight > 0)
                    ? Math.Round(a.HotWeight.Value / a.LiveWeight * 100, 2)
                    : 0;

            // Helper: Dress Rate (sale cost per lb of hot weight)
            static decimal DressRate(Animal a)
                => (a.HotWeight.HasValue && a.HotWeight.Value > 0)
                    ? Math.Round(SaleCost(a) / a.HotWeight.Value, 3)
                    : 0;

            // ── Active (non-condemned) for weight/cost totals ─────────────
            var active = all.Where(a => !a.IsCondemned).ToList();

            var totalLive = active.Sum(a => a.LiveWeight);
            var totalHot  = active.Sum(a => a.HotWeight ?? 0);
            var totalCost = active.Sum(SaleCost);
            var yieldPct  = totalLive > 0
                ? Math.Round(totalHot / totalLive * 100, 1) : 0;
            var dressRate = totalHot > 0
                ? Math.Round(totalCost / totalHot, 3) : 0;
            var avgCost   = active.Count > 0
                ? Math.Round(totalCost / active.Count, 4) : 0;

            // ── By vendor ─────────────────────────────────────────────────
            var byVendor = all
                .GroupBy(a => a.Vendor?.VendorName ?? "Unknown")
                .Select(g =>
                {
                    var activeG = g.Where(a => !a.IsCondemned).ToList();
                    var lw  = activeG.Sum(a => a.LiveWeight);
                    var hw  = activeG.Sum(a => a.HotWeight ?? 0);
                    var sc  = activeG.Sum(SaleCost);
                    return new VendorGroup
                    {
                        VendorName      = g.Key,
                        Count           = g.Count(),
                        Condemned       = g.Count(a => a.IsCondemned),
                        Passed          = g.Count(a => !a.IsCondemned),
                        TotalLiveWeight = lw,
                        TotalHotWeight  = hw,
                        TotalSaleCost   = sc,
                        AvgCost         = lw > 0 ? Math.Round(sc / lw, 4) : 0,
                        YieldPct        = lw > 0 ? Math.Round(hw / lw * 100, 1) : 0,
                        DressRate       = hw > 0 ? Math.Round(sc / hw, 3) : 0,
                        Animals         = g.OrderBy(a => a.ControlNo).ToList()
                    };
                })
                .OrderBy(g => g.VendorName)
                .ToList();

            // ── By animal type (Sheet1 equivalent) ────────────────────────
            // Categories: Cows, Bulls, Cows-ABF, Bulls-ABF, Steers, Canadian Cows, Canadian Bulls
            TypeSummaryRow BuildRow(string label, IEnumerable<Animal> animals)
            {
                var activeA = animals.Where(a => !a.IsCondemned).ToList();
                var cost    = activeA.Sum(SaleCost);
                var hw      = activeA.Sum(a => a.HotWeight ?? 0);
                return new TypeSummaryRow
                {
                    Category  = label,
                    Killed    = animals.Count(),
                    Condemned = animals.Count(a => a.IsCondemned),
                    Passed    = activeA.Count,
                    DressedWt = hw,
                    Cost      = cost,
                    AvgCost   = hw > 0 ? Math.Round(cost / hw, 4) : 0,
                };
            }

            // Sale bill animals
            var saleBill = all.Where(a => a.PurchaseType == "Sale bill").ToList();
            // Consignment animals
            var consignment = all.Where(a => a.PurchaseType == "Consignment").ToList();
            // Canadian
            var canadian = all.Where(a => a.Origin == "Canada").ToList();

            var byType = new List<TypeSummaryRow>
            {
                BuildRow("Cows",
                    saleBill.Where(a => a.AnimalType.Contains("COW", StringComparison.OrdinalIgnoreCase)
                                     && a.ProgramCode != "ABF")),
                BuildRow("Bulls",
                    saleBill.Where(a => a.AnimalType.Contains("BULL", StringComparison.OrdinalIgnoreCase)
                                     && a.ProgramCode != "ABF")),
                BuildRow("Steers",
                    saleBill.Where(a => a.AnimalType.Contains("STEER", StringComparison.OrdinalIgnoreCase))),
                BuildRow("Cows-ABF",
                    all.Where(a => a.AnimalType.Contains("COW", StringComparison.OrdinalIgnoreCase)
                                && a.ProgramCode == "ABF")),
                BuildRow("Bulls-ABF",
                    all.Where(a => a.AnimalType.Contains("BULL", StringComparison.OrdinalIgnoreCase)
                                && a.ProgramCode == "ABF")),
                BuildRow("Consignment Cows",
                    consignment.Where(a => a.AnimalType.Contains("COW", StringComparison.OrdinalIgnoreCase)
                                        && a.Origin != "Canada")),
                BuildRow("Consignment Bulls",
                    consignment.Where(a => a.AnimalType.Contains("BULL", StringComparison.OrdinalIgnoreCase)
                                        && a.Origin != "Canada")),
                BuildRow("Canadian Cows",
                    canadian.Where(a => a.AnimalType.Contains("COW", StringComparison.OrdinalIgnoreCase))),
                BuildRow("Canadian Bulls",
                    canadian.Where(a => a.AnimalType.Contains("BULL", StringComparison.OrdinalIgnoreCase))),
            };

            return new TallySummary
            {
                KillDate         = killDate,
                TotalAnimals     = all.Count,
                TotalCondemned   = all.Count(a => a.IsCondemned),
                TotalPassed      = active.Count,
                TotalLiveWeight  = totalLive,
                TotalHotWeight   = totalHot,
                TotalSaleCost    = totalCost,
                AverageYieldPct  = yieldPct,
                AverageDressRate = dressRate,
                AverageCost      = avgCost,
                ByVendor         = byVendor,
                ByType           = byType,
            };
        }

        public async Task<TallySummary> GetTodayKilledSummaryAsync()
        {
            return await GetTallySummaryAsync(DateTime.Today);
        }

        public async Task<IEnumerable<Animal>> SearchAnimalsByVendorNameAsync(string vendorName)
        {
            vendorName = (vendorName ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(vendorName))
            {
                return Enumerable.Empty<Animal>();
            }

            var animals = await _db.Animals
                .Include(a => a.Vendor)
                .Where(a => a.Vendor != null &&
                            a.Vendor.VendorName.Contains(vendorName))
                .OrderBy(a => a.Vendor!.VendorName)
                .ThenByDescending(a => a.KillDate)
                .ThenBy(a => a.ControlNo)
                .ToListAsync();

            var statusWasUpdated = false;
            foreach (var animal in animals)
            {
                if (SyncKillStatus(animal))
                {
                    statusWasUpdated = true;
                }
            }

            if (statusWasUpdated)
            {
                await _db.SaveChangesAsync();
            }

            return animals;
        }

        private static string GetCurrentKillStatus(DateTime killDate)
        {
            return killDate.Date <= DateTime.Today ? KilledStatus : PendingStatus;
        }

        private static bool SyncKillStatus(Animal animal)
        {
            if (animal.KillStatus == FlaggedStatus || animal.KillStatus == VerifiedStatus)
            {
                return false;
            }

            var expectedStatus = GetCurrentKillStatus(animal.KillDate);
            if (animal.KillStatus == expectedStatus)
            {
                return false;
            }

            animal.KillStatus = expectedStatus;
            return true;
        }
    }
}
