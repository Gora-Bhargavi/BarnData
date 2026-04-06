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
        // One Tag1 can only appear once per vendor per kill date
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
        // Returns true if weight is OUTSIDE the acceptable range (warning, not block)
        public bool IsWeightOutOfRange(decimal liveWeight)
        {
            return liveWeight < _weightMin || liveWeight > _weightMax;
        }

        // ── Create animal ─────────────────────────────────────────────────
        public async Task<(bool Success, string ErrorMessage)> CreateAsync(Animal animal)
        {
            // Duplicate tag check
            bool isDuplicate = await IsTagDuplicateAsync(
                animal.TagNumber1, animal.KillDate, animal.VendorID);

            if (isDuplicate)
                return (false,
                    $"Tag Number '{animal.TagNumber1}' already exists for this vendor on {animal.KillDate:MM/dd/yyyy}. Duplicate tags are not allowed.");

            animal.TagNumber1 = animal.TagNumber1.Trim();
            animal.TagNumber2 = animal.TagNumber2?.Trim() ?? string.Empty;
            animal.Comment = animal.Comment?.Trim() ?? string.Empty;
            animal.CreatedAt = DateTime.Now;
            animal.KillStatus = GetCurrentKillStatus(animal.KillDate);

            _db.Animals.Add(animal);
            await _db.SaveChangesAsync();

            return (true, string.Empty);
        }

        // ── Update animal ─────────────────────────────────────────────────
        public async Task<(bool Success, string ErrorMessage)> UpdateAsync(Animal animal)
        {
            // Duplicate tag check — exclude the animal being edited
            bool isDuplicate = await IsTagDuplicateAsync(
                animal.TagNumber1, animal.KillDate, animal.VendorID, animal.ControlNo);

            if (isDuplicate)
                return (false,
                    $"Tag Number '{animal.TagNumber1}' already exists for this vendor on {animal.KillDate:MM/dd/yyyy}.");

            var existing = await _db.Animals.FindAsync(animal.ControlNo);
            if (existing == null)
                return (false, "Animal record not found.");

            // Update all editable fields
            existing.VendorID            = animal.VendorID;
            existing.TagNumber1          = animal.TagNumber1.Trim();
            existing.TagNumber2          = animal.TagNumber2?.Trim() ?? string.Empty;
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
            existing.Comment             = animal.Comment?.Trim() ?? string.Empty;
            existing.AnimalControlNumber = animal.AnimalControlNumber;
            existing.State               = animal.State;
            existing.BuyerName           = animal.BuyerName;
            existing.VetName             = animal.VetName;
            existing.OfficeUse2          = animal.OfficeUse2;
            existing.UpdatedAt           = DateTime.Now;
            SyncKillStatus(existing);

            await _db.SaveChangesAsync();
            return (true, string.Empty);
        }

        // ── Soft delete ───────────────────────────────────────────────────
        // Sets KillStatus = 'Flagged' — never hard deletes from the DB
        public async Task<bool> DeleteAsync(int controlNo)
        {
            var animal = await _db.Animals.FindAsync(controlNo);
            if (animal == null) return false;

            animal.KillStatus = FlaggedStatus;
            animal.UpdatedAt  = DateTime.Now;

            await _db.SaveChangesAsync();
            return true;
        }

        // ── Tally summary ─────────────────────────────────────────────────
        public async Task<TallySummary> GetTallySummaryAsync(
            DateTime killDate, int? vendorId = null)
        {
            var animals = (await GetByKillDateAsync(killDate, vendorId))
                .Where(a => a.KillStatus != FlaggedStatus)
                .ToList();

            var totalLive = animals.Sum(a => a.LiveWeight);
            var totalHot  = animals.Sum(a => a.HotWeight ?? 0);
            var yieldPct  = totalLive > 0
                ? Math.Round(totalHot / totalLive * 100, 1)
                : 0;

            var byVendor = animals
                .GroupBy(a => a.Vendor?.VendorName ?? "Unknown")
                .Select(g => new VendorGroup
                {
                    VendorName      = g.Key,
                    Count           = g.Count(),
                    TotalLiveWeight = g.Sum(a => a.LiveWeight),
                    TotalHotWeight  = g.Sum(a => a.HotWeight ?? 0),
                    Animals         = g.OrderBy(a => a.ControlNo).ToList()
                })
                .OrderBy(g => g.VendorName)
                .ToList();

            return new TallySummary
            {
                KillDate        = killDate,
                TotalAnimals    = animals.Count,
                TotalLiveWeight = totalLive,
                TotalHotWeight  = totalHot,
                AverageYieldPct = yieldPct,
                ByVendor        = byVendor
            };
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
