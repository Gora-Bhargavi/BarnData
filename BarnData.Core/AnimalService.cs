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

        public AnimalService(BarnDataContext db, IConfiguration config)
        {
            _db = db;
            _weightMin = config.GetValue<decimal>("AppSettings:LiveWeightMinLbs", 300);
            _weightMax = config.GetValue<decimal>("AppSettings:LiveWeightMaxLbs", 2500);
        }

        // ── Queries ───────────────────────────────────────────────────────
        public async Task<IEnumerable<Animal>> GetByKillDateAsync(
            DateTime killDate, int? vendorId = null)
        {
            var query = _db.Animals
                .Include(a => a.Vendor)
                .Where(a => a.KillDate.HasValue &&
                            a.KillDate.Value.Date == killDate.Date &&
                            a.KillStatus == "Killed");

            if (vendorId.HasValue)
                query = query.Where(a => a.VendorID == vendorId.Value);

            return await query
                .OrderBy(a => a.Vendor!.VendorName)
                .ThenBy(a => a.ControlNo)
                .ToListAsync();
        }

        public async Task<IEnumerable<Animal>> GetPendingAsync(int? vendorId = null)
        {
            var query = _db.Animals
                .Include(a => a.Vendor)
                .Where(a => a.KillStatus == "Pending");

            if (vendorId.HasValue)
                query = query.Where(a => a.VendorID == vendorId.Value);

            return await query
                .OrderBy(a => a.Vendor!.VendorName)
                .ThenBy(a => a.PurchaseDate)
                .ThenBy(a => a.ControlNo)
                .ToListAsync();
        }

        public async Task<IEnumerable<Animal>> GetAllAsync(int? vendorId = null)
        {
            var query = _db.Animals
                .Include(a => a.Vendor)
                .Where(a => a.KillStatus != "Flagged");

            if (vendorId.HasValue)
                query = query.Where(a => a.VendorID == vendorId.Value);

            return await query
                .OrderByDescending(a => a.CreatedAt)
                .ToListAsync();
        }

        public async Task<Animal?> GetByControlNoAsync(int controlNo)
        {
            return await _db.Animals
                .Include(a => a.Vendor)
                .FirstOrDefaultAsync(a => a.ControlNo == controlNo);
        }

        // ── Validation ────────────────────────────────────────────────────
        public async Task<bool> IsTagDuplicateAsync(
            string tag1, int vendorId, int? excludeControlNo = null)
        {
            var query = _db.Animals.Where(a =>
                a.TagNumber1 == tag1.Trim() &&
                a.VendorID == vendorId);

            if (excludeControlNo.HasValue)
                query = query.Where(a => a.ControlNo != excludeControlNo.Value);

            return await query.AnyAsync();
        }

        public bool IsWeightOutOfRange(decimal liveWeight)
            => liveWeight > 0 && (liveWeight < _weightMin || liveWeight > _weightMax);

        // ── Create ────────────────────────────────────────────────────────
        public async Task<(bool Success, string ErrorMessage)> CreateAsync(Animal animal)
        {
            bool isDuplicate = await IsTagDuplicateAsync(
                animal.TagNumber1, animal.VendorID);

            if (isDuplicate)
                return (false,
                    $"Tag '{animal.TagNumber1}' already exists for this vendor. Duplicate tags are not allowed.");

            animal.TagNumber1 = animal.TagNumber1.Trim();
            animal.TagNumber2 = animal.TagNumber2?.Trim();
            animal.CreatedAt  = DateTime.Now;
            animal.KillStatus = "Pending";

            _db.Animals.Add(animal);
            await _db.SaveChangesAsync();
            return (true, string.Empty);
        }

        // ── Bulk import from sale bill ────────────────────────────────────
        public async Task<(int Imported, int Skipped, List<string> Errors)>
            BulkImportAsync(IEnumerable<Animal> animals)
        {
            int imported = 0, skipped = 0;
            var errors = new List<string>();

            foreach (var animal in animals)
            {
                bool isDup = await IsTagDuplicateAsync(animal.TagNumber1, animal.VendorID);
                if (isDup)
                {
                    skipped++;
                    errors.Add($"Skipped tag '{animal.TagNumber1}' — already exists for {animal.Vendor?.VendorName ?? animal.VendorID.ToString()}");
                    continue;
                }

                animal.CreatedAt  = DateTime.Now;
                animal.KillStatus = "Pending";
                _db.Animals.Add(animal);
                imported++;
            }

            await _db.SaveChangesAsync();
            return (imported, skipped, errors);
        }

        // ── Mark animals as killed ────────────────────────────────────────
        public async Task<int> MarkKilledAsync(IEnumerable<int> controlNos, DateTime killDate)
        {
            var animals = await _db.Animals
                .Where(a => controlNos.Contains(a.ControlNo) && a.KillStatus == "Pending")
                .ToListAsync();

            foreach (var a in animals)
            {
                a.KillDate   = killDate;
                a.KillStatus = "Killed";
                a.UpdatedAt  = DateTime.Now;
            }

            await _db.SaveChangesAsync();
            return animals.Count;
        }

        // ── Update ────────────────────────────────────────────────────────
        public async Task<(bool Success, string ErrorMessage)> UpdateAsync(Animal animal)
        {
            bool isDuplicate = await IsTagDuplicateAsync(
                animal.TagNumber1, animal.VendorID, animal.ControlNo);

            if (isDuplicate)
                return (false,
                    $"Tag '{animal.TagNumber1}' already exists for this vendor.");

            var existing = await _db.Animals.FindAsync(animal.ControlNo);
            if (existing == null) return (false, "Animal record not found.");

            existing.VendorID            = animal.VendorID;
            existing.TagNumber1          = animal.TagNumber1.Trim();
            existing.TagNumber2          = animal.TagNumber2?.Trim();
            existing.Tag3                = animal.Tag3;
            existing.AnimalType          = animal.AnimalType;
            existing.AnimalType2         = animal.AnimalType2;
            existing.ProgramCode         = animal.ProgramCode;
            existing.PurchaseDate        = animal.PurchaseDate;
            existing.PurchaseType        = animal.PurchaseType;
            existing.LiveWeight          = animal.LiveWeight;
            existing.LiveRate            = animal.LiveRate;
            existing.ConsignmentRate     = animal.ConsignmentRate;
            existing.KillDate            = animal.KillDate;
            existing.HotWeight           = animal.HotWeight;
            existing.Grade               = animal.Grade;
            existing.Grade2              = animal.Grade2;
            existing.HealthScore         = animal.HealthScore;
            existing.FetalBlood          = animal.FetalBlood;
            existing.Comment             = animal.Comment;
            existing.AnimalControlNumber = animal.AnimalControlNumber;
            existing.State               = animal.State;
            existing.BuyerName           = animal.BuyerName;
            existing.VetName             = animal.VetName;
            existing.OfficeUse2          = animal.OfficeUse2;
            existing.Origin              = animal.Origin;
            existing.IsCondemned         = animal.IsCondemned;
            existing.UpdatedAt           = DateTime.Now;

            await _db.SaveChangesAsync();
            return (true, string.Empty);
        }

        // ── Soft delete ───────────────────────────────────────────────────
        public async Task<bool> DeleteAsync(int controlNo)
        {
            var animal = await _db.Animals.FindAsync(controlNo);
            if (animal == null) return false;
            animal.KillStatus = "Flagged";
            animal.UpdatedAt  = DateTime.Now;
            await _db.SaveChangesAsync();
            return true;
        }

        // ── Tally summary ─────────────────────────────────────────────────
        public async Task<TallySummary> GetTallySummaryAsync(
            DateTime killDate, int? vendorId = null)
        {
            var all = (await GetByKillDateAsync(killDate, vendorId))
                .Where(a => a.KillStatus != "Flagged")
                .ToList();

            var active = all.Where(a => !a.IsCondemned).ToList();

            var totalLive = active.Sum(a => a.LiveWeight);
            var totalHot  = active.Sum(a => a.HotWeight ?? 0);
            var totalCost = active.Sum(a => a.SaleCost);
            var yieldPct  = totalLive > 0 ? Math.Round(totalHot / totalLive * 100, 1) : 0;
            var dressRate = totalHot  > 0 ? Math.Round(totalCost / totalHot, 3) : 0;

            // ── By vendor ─────────────────────────────────────────────────
            var byVendor = all
                .GroupBy(a => a.Vendor?.VendorName ?? "Unknown")
                .Select(g =>
                {
                    var activeG = g.Where(a => !a.IsCondemned).ToList();
                    var lw = activeG.Sum(a => a.LiveWeight);
                    var hw = activeG.Sum(a => a.HotWeight ?? 0);
                    var sc = activeG.Sum(a => a.SaleCost);
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
            TypeSummaryRow BuildRow(string label, IEnumerable<Animal> animals)
            {
                var activeA = animals.Where(a => !a.IsCondemned).ToList();
                var cost = activeA.Sum(a => a.SaleCost);
                var hw   = activeA.Sum(a => a.HotWeight ?? 0);
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

            bool IsCow(Animal a)  => a.AnimalType.Contains("Cow",   StringComparison.OrdinalIgnoreCase);
            bool IsBull(Animal a) => a.AnimalType.Contains("Bull",  StringComparison.OrdinalIgnoreCase)
                                  || a.AnimalType.Contains("Steer", StringComparison.OrdinalIgnoreCase); // Steers count as Bulls

            var saleBill    = all.Where(a => a.PurchaseType == "Sale Bill").ToList();
            var consignment = all.Where(a => a.PurchaseType == "Consignment Bill").ToList();
            var canadian    = all.Where(a => a.Origin == "Canada").ToList();

            var byType = new List<TypeSummaryRow>
            {
                BuildRow("Cows",             saleBill.Where(a => IsCow(a)  && a.ProgramCode != "ABF")),
                BuildRow("Bulls",            saleBill.Where(a => IsBull(a) && a.ProgramCode != "ABF")),
                BuildRow("Cows-ABF",         all.Where(a => IsCow(a)  && a.ProgramCode == "ABF")),
                BuildRow("Bulls-ABF",        all.Where(a => IsBull(a) && a.ProgramCode == "ABF")),
                BuildRow("Consignment Cows", consignment.Where(a => IsCow(a)  && a.Origin != "Canada")),
                BuildRow("Consignment Bulls",consignment.Where(a => IsBull(a) && a.Origin != "Canada")),
                BuildRow("Canadian Cows",    canadian.Where(IsCow)),
                BuildRow("Canadian Bulls",   canadian.Where(IsBull)),
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
                AverageCost      = active.Count > 0 ? Math.Round(totalCost / active.Count, 4) : 0,
                ByVendor         = byVendor,
                ByType           = byType,
            };
        }
    }
}
