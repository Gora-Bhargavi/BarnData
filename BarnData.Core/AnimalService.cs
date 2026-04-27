using BarnData.Data;
using BarnData.Data.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System;
using System.Data;

namespace BarnData.Core.Services
{
    public class AnimalService : IAnimalService
    {
        private readonly BarnDataContext _db;

            private static string? NormalizeAcn(string? acn)
            {
                if (string.IsNullOrWhiteSpace(acn)) return null;
                var v = acn.Trim();
                return v.All(ch => ch == '0') ? null : v;
            }
        private readonly decimal _weightMin;
        private readonly decimal _weightMax;

        public AnimalService(BarnDataContext db, IConfiguration config)
        {
            _db = db;
            _weightMin = config.GetValue<decimal>("AppSettings:LiveWeightMinLbs", 300);
            _weightMax = config.GetValue<decimal>("AppSettings:LiveWeightMaxLbs", 2500);
        }

        //  Helper: load vendors and attach to animal list 
        private async Task AttachVendors(List<Animal> animals)
        {
            var ids = animals.Select(a => a.VendorID).Distinct().ToArray();
            if (!ids.Any()) return;
            var idList = string.Join(",", ids);
            var vendors = await _db.Vendors
                .FromSqlRaw($"SELECT * FROM tbl_vendor_master WHERE VendorID IN ({idList})")
                .ToListAsync();
            var dict = vendors.ToDictionary(v => v.VendorID);
            foreach (var a in animals)
                if (dict.TryGetValue(a.VendorID, out var v)) a.Vendor = v;
        }

        // GetPendingAsync 
        public async Task<IEnumerable<Animal>> GetPendingAsync(int? vendorId = null)
        {
            // Raw SQL to completely bypass EF Core translation issues
            var sql = vendorId.HasValue
                ? @"SELECT a.* FROM tbl_barn_animal_entry a
                    WHERE a.KillStatus = 'Pending' AND a.VendorID = {0}
                    ORDER BY a.VendorID, a.PurchaseDate, a.ControlNo"
                : @"SELECT a.* FROM tbl_barn_animal_entry a
                    WHERE a.KillStatus = 'Pending'
                    ORDER BY a.VendorID, a.PurchaseDate, a.ControlNo";

            var animals = vendorId.HasValue
                ? await _db.Animals.FromSqlRaw(sql, vendorId.Value).ToListAsync()
                : await _db.Animals.FromSqlRaw(sql).ToListAsync();

            await AttachVendors(animals);
            return animals;
        }

        // GetAllAsync
        public async Task<IEnumerable<Animal>> GetAllAsync(int? vendorId = null)
        {
            var sql = vendorId.HasValue
                ? @"SELECT a.* FROM tbl_barn_animal_entry a
                    WHERE a.KillStatus IN ('Pending','Killed','Verified')
                    AND a.VendorID = {0}
                    ORDER BY a.CreatedAt DESC"
                : @"SELECT a.* FROM tbl_barn_animal_entry a
                    WHERE a.KillStatus IN ('Pending','Killed','Verified')
                    ORDER BY a.CreatedAt DESC";

            var animals = vendorId.HasValue
                ? await _db.Animals.FromSqlRaw(sql, vendorId.Value).ToListAsync()
                : await _db.Animals.FromSqlRaw(sql).ToListAsync();

            await AttachVendors(animals);
            return animals;
        }

        //  GetByKillDateAsync 
        public async Task<IEnumerable<Animal>> GetByKillDateAsync(
            DateTime killDate, int? vendorId = null)
        {
            var dateStr = killDate.ToString("yyyy-MM-dd");
            var sql = vendorId.HasValue
                ? @"SELECT a.* FROM tbl_barn_animal_entry a
                    WHERE CAST(a.KillDate AS DATE) = {0}
                    AND a.KillStatus = 'Killed'
                    AND a.VendorID = {1}
                    ORDER BY a.VendorID, a.ControlNo"
                : @"SELECT a.* FROM tbl_barn_animal_entry a
                    WHERE CAST(a.KillDate AS DATE) = {0}
                    AND a.KillStatus = 'Killed'
                    ORDER BY a.VendorID, a.ControlNo";

            var animals = vendorId.HasValue
                ? await _db.Animals.FromSqlRaw(sql, dateStr, vendorId.Value).ToListAsync()
                : await _db.Animals.FromSqlRaw(sql, dateStr).ToListAsync();

            await AttachVendors(animals);
            return animals;
        }

        //  Multi-vendor overloads
        public async Task<IEnumerable<Animal>> GetPendingByVendorsAsync(IEnumerable<int> vendorIds)
        {
            var ids = string.Join(",", vendorIds.Select(i => i.ToString()));
            if (string.IsNullOrEmpty(ids)) return await GetPendingAsync();
            var sql = $@"SELECT a.* FROM tbl_barn_animal_entry a
                WHERE a.KillStatus = 'Pending' AND a.VendorID IN ({ids})
                ORDER BY a.VendorID, a.PurchaseDate, a.ControlNo";
            var animals = await _db.Animals.FromSqlRaw(sql).ToListAsync();
            await AttachVendors(animals);
            return animals;
        }

        public async Task<IEnumerable<Animal>> GetAllByVendorsAsync(IEnumerable<int> vendorIds)
        {
            var ids = string.Join(",", vendorIds.Select(i => i.ToString()));
            if (string.IsNullOrEmpty(ids)) return await GetAllAsync();
            var sql = $@"SELECT a.* FROM tbl_barn_animal_entry a
                WHERE a.KillStatus IN ('Pending','Killed','Verified') AND a.VendorID IN ({ids})
                ORDER BY a.CreatedAt DESC";
            var animals = await _db.Animals.FromSqlRaw(sql).ToListAsync();
            await AttachVendors(animals);
            return animals;
        }

        public async Task<IEnumerable<Animal>> GetByKillDateByVendorsAsync(DateTime killDate, IEnumerable<int> vendorIds)
        {
            var ids    = string.Join(",", vendorIds.Select(i => i.ToString()));
            var dateStr = killDate.ToString("yyyy-MM-dd");
            if (string.IsNullOrEmpty(ids)) return await GetByKillDateAsync(killDate);
            var sql = $@"SELECT a.* FROM tbl_barn_animal_entry a
                WHERE CAST(a.KillDate AS DATE) = '{dateStr}'
                AND a.KillStatus = 'Killed' AND a.VendorID IN ({ids})
                ORDER BY a.VendorID, a.ControlNo";
            var animals = await _db.Animals.FromSqlRaw(sql).ToListAsync();
            await AttachVendors(animals);
            return animals;
        }

        //  Tag-based lookup - for ACN auto-match during HW import 
        public async Task<IEnumerable<Animal>> GetByTagsAsync(IEnumerable<string> tags)
        {
            if (!tags.Any()) return Enumerable.Empty<Animal>();
            // Escape single quotes and build IN list
            var tagList = string.Join(",", tags.Select(t => "'" + t.Replace("'", "''").Trim() + "'"));
            var sql = $@"SELECT a.* FROM tbl_barn_animal_entry a
                WHERE a.TagNumber1 IN ({tagList})
                   OR a.TagNumber2 IN ({tagList})
                   OR a.Tag3 IN ({tagList})
                ORDER BY a.ControlNo";
            var animals = await _db.Animals.FromSqlRaw(sql).ToListAsync();
            await AttachVendors(animals);
            return animals;
        }

        // GetByControlNoAsync 
        public async Task<Animal?> GetByControlNoAsync(int controlNo)
        {
            var animals = await _db.Animals
                .FromSqlRaw("SELECT * FROM tbl_barn_animal_entry WHERE ControlNo = {0}", controlNo)
                .ToListAsync();
            var animal = animals.FirstOrDefault();
            if (animal != null) await AttachVendors(new List<Animal> { animal });
            return animal;
        }

        // GetByTagSuffixAsync — finds animals whose Tag1/Tag2/Tag3 ends with the given suffix
    public async Task<IEnumerable<Animal>> GetByTagSuffixAsync(string suffix)
    {
        if (string.IsNullOrWhiteSpace(suffix)) return Enumerable.Empty<Animal>();
        var s = suffix.Trim().TrimStart('0');  // strip leading zeros for comparison
        if (string.IsNullOrEmpty(s)) return Enumerable.Empty<Animal>();
        var likePat = "%" + s;
        var animals = await _db.Animals
            .FromSqlRaw(@"SELECT a.* FROM tbl_barn_animal_entry a
                WHERE a.TagNumber1 LIKE {0}
                   OR a.TagNumber2 LIKE {0}
                   OR a.Tag3       LIKE {0}
                   OR CAST(CONVERT(bigint, CASE WHEN ISNUMERIC(a.TagNumber1)=1 THEN a.TagNumber1 ELSE NULL END) AS nvarchar) LIKE {0}",
                likePat)
            .ToListAsync();
        await AttachVendors(animals);
        return animals;
    }

    // GetByTagPatternAsync — wildcard match: '?' replaced with any digit, regex applied in-memory
    public async Task<IEnumerable<Animal>> GetByTagPatternAsync(string pattern)
    {
        if (string.IsNullOrWhiteSpace(pattern)) return Enumerable.Empty<Animal>();
        // Build SQL LIKE pattern: ? → _ (single char wildcard in SQL)
        var sqlLike = pattern.Replace('?', '_');
        var animals = await _db.Animals
            .FromSqlRaw(@"SELECT a.* FROM tbl_barn_animal_entry a
                WHERE a.TagNumber1 LIKE {0}
                   OR a.TagNumber2 LIKE {0}
                   OR a.Tag3       LIKE {0}",
                sqlLike)
            .ToListAsync();
        await AttachVendors(animals);
        return animals;
    }

    // GetAllPendingAsync — pending animals for weight proximity matching (capped for performance)
    public async Task<IEnumerable<Animal>> GetAllPendingAsync()
    {
        // TOP 10000 guard — weight matching only needs current kill cycle (typically <2000 animals)
        var animals = await _db.Animals
            .FromSqlRaw(@"SELECT TOP 10000 * FROM tbl_barn_animal_entry
                WHERE KillStatus = 'Pending'
                ORDER BY CreatedAt DESC")
            .ToListAsync();
        await AttachVendors(animals);
        return animals;
    }

    // GetPendingPagedAsync — paginated pending list for Animal Index (avoids full-table load)
    public async Task<(IEnumerable<Animal> Items, int TotalCount)> GetPendingPagedAsync(
        int? vendorId, int page, int pageSize)
    {
        var offset = (page - 1) * pageSize;

        // Count via raw ADO.NET to avoid EF scalar mapping issues
        int total;
        using (var cmd = _db.Database.GetDbConnection().CreateCommand())
        {
            cmd.CommandText = vendorId.HasValue
                ? "SELECT COUNT(*) FROM tbl_barn_animal_entry WHERE KillStatus='Pending' AND VendorID=@vid"
                : "SELECT COUNT(*) FROM tbl_barn_animal_entry WHERE KillStatus='Pending'";
            if (vendorId.HasValue)
            {
                var p = cmd.CreateParameter(); p.ParameterName = "@vid"; p.Value = vendorId.Value; cmd.Parameters.Add(p);
            }
            if (cmd.Connection!.State != System.Data.ConnectionState.Open)
                await cmd.Connection.OpenAsync();
            total = Convert.ToInt32(await cmd.ExecuteScalarAsync() ?? 0);
        }

        var dataSql = vendorId.HasValue
            ? @"SELECT * FROM tbl_barn_animal_entry
                WHERE KillStatus='Pending' AND VendorID={0}
                ORDER BY CreatedAt DESC
                OFFSET {1} ROWS FETCH NEXT {2} ROWS ONLY"
            : @"SELECT * FROM tbl_barn_animal_entry
                WHERE KillStatus='Pending'
                ORDER BY CreatedAt DESC
                OFFSET {0} ROWS FETCH NEXT {1} ROWS ONLY";

        var animals = vendorId.HasValue
            ? await _db.Animals.FromSqlRaw(dataSql, vendorId.Value, offset, pageSize).ToListAsync()
            : await _db.Animals.FromSqlRaw(dataSql, offset, pageSize).ToListAsync();

        await AttachVendors(animals);
        return (animals, total);
    }

    // IsTagDuplicateAsync 
        public async Task<bool> IsTagDuplicateAsync(
            string tag1, int vendorId, int? excludeControlNo = null)
        {
            var trimmed = tag1.Trim();
            if (excludeControlNo.HasValue)
            {
                var count = await _db.Animals
                    .FromSqlRaw(
                        @"SELECT * FROM tbl_barn_animal_entry
                          WHERE TagNumber1 = {0} AND VendorID = {1} AND ControlNo != {2}",
                        trimmed, vendorId, excludeControlNo.Value)
                    .CountAsync();
                return count > 0;
            }
            else
            {
                var count = await _db.Animals
                    .FromSqlRaw(
                        @"SELECT * FROM tbl_barn_animal_entry
                          WHERE TagNumber1 = {0} AND VendorID = {1}",
                        trimmed, vendorId)
                    .CountAsync();
                return count > 0;
            }
        }

        public async Task<HashSet<(string Tag, int VendorId)>> GetAllTagVendorKeysAsync()
        {
            var rows = await _db.Animals
                .AsNoTracking()
                .Where(a => a.TagNumber1 != null && a.VendorID > 0)
                .Select(a => new { a.TagNumber1, a.VendorID })
                .ToListAsync();

            // Custom equality comparer: case-insensitive on Tag (to match SQL),
            // exact match on VendorId.
            var set = new HashSet<(string Tag, int VendorId)>(new TagVendorComparer());
            foreach (var r in rows)
            {
                set.Add((r.TagNumber1!.Trim(), r.VendorID));
            }
            return set;
        }

        // equality comparer mirroring SQL default CI collation.
        private sealed class TagVendorComparer : IEqualityComparer<(string Tag, int VendorId)>
        {
            public bool Equals((string Tag, int VendorId) a, (string Tag, int VendorId) b)
                => a.VendorId == b.VendorId
                && string.Equals(a.Tag, b.Tag, StringComparison.OrdinalIgnoreCase);

            public int GetHashCode((string Tag, int VendorId) v)
                => HashCode.Combine(
                    StringComparer.OrdinalIgnoreCase.GetHashCode(v.Tag ?? ""),
                    v.VendorId);
        }

        //  IsWeightOutOfRange 
        public bool IsWeightOutOfRange(decimal liveWeight)
            => liveWeight > 0 && (liveWeight < _weightMin || liveWeight > _weightMax);

        // CreateAsync
        public async Task<(bool Success, string ErrorMessage)> CreateAsync(Animal animal)
        {
            bool isDuplicate = await IsTagDuplicateAsync(animal.TagNumber1, animal.VendorID);
            if (isDuplicate)
                return (false, $"Tag '{animal.TagNumber1}' already exists for this vendor.");

            animal.TagNumber1 = animal.TagNumber1.Trim();
            animal.TagNumber2 = animal.TagNumber2?.Trim();
            animal.CreatedAt  = DateTime.Now;
            animal.KillStatus = animal.KillStatus ?? "Pending";

            _db.Animals.Add(animal);
            await _db.SaveChangesAsync();
            return (true, string.Empty);
        }

        //  BulkImportAsync
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
                    errors.Add($"Skipped tag '{animal.TagNumber1}' — already exists for vendor {animal.VendorID}");
                    continue;
                }

                animal.CreatedAt = DateTime.Now;
                _db.Animals.Add(animal);
                imported++;
            }

            if (imported > 0)
                await _db.SaveChangesAsync();

            return (imported, skipped, errors);
        }

        // MarkKilledAsync 
        public async Task<int> MarkKilledAsync(IEnumerable<int> controlNos, DateTime killDate)
        {
            var idArray = controlNos.ToArray();
            if (!idArray.Any()) return 0;

            var animals = await _db.Animals
                .FromSqlRaw("SELECT * FROM tbl_barn_animal_entry WHERE KillStatus = 'Pending'")
                .ToListAsync();

            var toUpdate = animals.Where(a => idArray.Contains(a.ControlNo)).ToList();

            foreach (var a in toUpdate)
            {
                a.KillDate   = killDate;
                a.KillStatus = "Killed";
                a.UpdatedAt  = DateTime.Now;
            }

            if (toUpdate.Any())
                await _db.SaveChangesAsync();

            return toUpdate.Count;
        }
        //Saving inline kill fields wityhout changing the kill status
        public async Task<int> SaveKillDataAsync(IEnumerable<KillAnimalData> animalData)
        {
            var dataList = animalData.ToList();
            if (!dataList.Any()) return 0;

            var ids = dataList.Select(d => d.ControlNo).ToArray();
            var idList = string.Join(",", ids);

            var animals = await _db.Animals
                .FromSqlRaw($"SELECT * FROM tbl_barn_animal_entry WHERE ControlNo IN ({idList})")
                .ToListAsync();

            var dataDict = dataList.ToDictionary(d => d.ControlNo);

            foreach (var a in animals)
            {
                if (!dataDict.TryGetValue(a.ControlNo, out var d)) continue;

                a.UpdatedAt   = DateTime.Now;

                a.IsCondemned = d.IsCondemned;

                //Update AnimalControlNumber only when a real value is provided.
                // Blank / null from client means "leave unchanged" — this prevents the Flagged-for-review
                // Pick flow (and any other partial save) from silently wiping an existing ACN.
                var normalizedAcn = NormalizeAcn(d.AnimalControlNumber);
                if (normalizedAcn != null)
                {
                a.AnimalControlNumber = normalizedAcn;
                }
                if (d.HotWeight.HasValue && d.HotWeight > 0)
                    a.HotWeight = d.HotWeight;
                
                if (!string.IsNullOrWhiteSpace(d.Grade))
                    a.Grade = d.Grade.Trim();
                
                if (d.HealthScore.HasValue && d.HealthScore > 0)
                    a.HealthScore = d.HealthScore;
                
                if(d.LiveWeight.HasValue && d.LiveWeight >0)
                    a.LiveWeight = d.LiveWeight.Value;

                //optional: if kill date is provided on save, persist it
                //if (d.KillDate.HasValue)
                 //   a.KillDate = d.KillDate.Value;

                if(!string.IsNullOrWhiteSpace(d.State))
                    a.State = d.State.Trim();
                if(!string.IsNullOrWhiteSpace(d.VetName))
                    a.VetName = d.VetName.Trim();
                if(!string.IsNullOrWhiteSpace(d.OfficeUse2))
                    a.OfficeUse2 = d.OfficeUse2.Trim();
                if(!string.IsNullOrWhiteSpace(d.Comment))
                    a.Comment = d.Comment.Trim();
            }
            if(animals.Any())
                await _db.SaveChangesAsync();
            
            return animals.Count;
        }

       //  MarkKilledWithDataAsync - saves HotWeight, Grade, HS, Condemned 
        public async Task<int> MarkKilledWithDataAsync(
            IEnumerable<KillAnimalData> animalData, DateTime killDate)
        {
            var dataList = animalData.ToList();
            if (!dataList.Any()) return 0;

            var ids = dataList.Select(d => d.ControlNo).ToArray();
            var idList = string.Join(",", ids);

            var animals = await _db.Animals
                .FromSqlRaw($"SELECT * FROM tbl_barn_animal_entry WHERE ControlNo IN ({idList})")
                .ToListAsync();

            var dataDict = dataList.ToDictionary(d => d.ControlNo);

            foreach (var a in animals)
            {
                if (!dataDict.TryGetValue(a.ControlNo, out var d)) continue;
                a.KillDate    = d.KillDate ?? killDate;
                a.KillStatus  = "Killed";
                a.UpdatedAt   = DateTime.Now;
                a.IsCondemned = d.IsCondemned;

                // Only overwrite ACN when a real value is provided; blank means "leave unchanged".
                var normalizedAcn = NormalizeAcn(d.AnimalControlNumber);
                if (normalizedAcn != null)
                a.AnimalControlNumber = normalizedAcn;
                        
                if (d.HotWeight.HasValue && d.HotWeight > 0)
                    a.HotWeight = d.HotWeight;
                if (!string.IsNullOrWhiteSpace(d.Grade))
                    a.Grade = d.Grade.Trim();
                if (d.HealthScore.HasValue && d.HealthScore > 0)
                    a.HealthScore = d.HealthScore;
                
                if(d.LiveWeight.HasValue && d.LiveWeight >0)
                    a.LiveWeight = d.LiveWeight.Value;
            }

            if (animals.Any())
                await _db.SaveChangesAsync();

            return animals.Count;
        }

        //  UpdateAsync 
        public async Task<(bool Success, string ErrorMessage)> UpdateAsync(Animal animal)
        {
            bool isDuplicate = await IsTagDuplicateAsync(
                animal.TagNumber1, animal.VendorID, animal.ControlNo);
            if (isDuplicate)
                return (false, $"Tag '{animal.TagNumber1}' already exists for this vendor.");

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

        //  DeleteAsync 
        public async Task<bool> DeleteAsync(int controlNo)
        {
            var animal = await _db.Animals.FindAsync(controlNo);
            if (animal == null) return false;
            animal.KillStatus = "Flagged";
            animal.UpdatedAt  = DateTime.Now;
            await _db.SaveChangesAsync();
            return true;
        }

        //  GetTallySummaryAsync 
        public async Task<TallySummary> GetTallySummaryAsync(
            DateTime killDate, int? vendorId = null)
        {
            var all = (await GetByKillDateAsync(killDate, vendorId)).ToList();
            var active = all.Where(a => !a.IsCondemned).ToList();

            var totalLive = active.Sum(a => a.LiveWeight);
            var totalHot  = active.Sum(a => a.HotWeight ?? 0);
            var totalCost = active.Sum(a => a.SaleCost);
            var yieldPct  = totalLive > 0 ? Math.Round(totalHot / totalLive * 100, 1) : 0;
            var dressRate = totalHot  > 0 ? Math.Round(totalCost / totalHot, 3)       : 0;

            bool IsCow(Animal a)  => a.AnimalType.Contains("Cow",   StringComparison.OrdinalIgnoreCase);
            bool IsBull(Animal a) => a.AnimalType.Contains("Bull",  StringComparison.OrdinalIgnoreCase)
                                  || a.AnimalType.Contains("Steer", StringComparison.OrdinalIgnoreCase);

            TypeSummaryRow BuildRow(string label, IEnumerable<Animal> src)
            {
                var list = src.ToList();
                var act  = list.Where(a => !a.IsCondemned).ToList();
                var cost = act.Sum(a => a.SaleCost);
                var hw   = act.Sum(a => a.HotWeight ?? 0);
                return new TypeSummaryRow
                {
                    Category  = label,
                    Killed    = list.Count,
                    Condemned = list.Count(a => a.IsCondemned),
                    Passed    = act.Count,
                    DressedWt = hw,
                    Cost      = cost,
                    AvgCost   = hw > 0 ? Math.Round(cost / hw, 4) : 0,
                };
            }

            var sale = all.Where(a => a.PurchaseType == "Sale Bill").ToList();
            var cons = all.Where(a => a.PurchaseType == "Consignment Bill").ToList();
            var cdn  = all.Where(a => a.Origin == "Canada").ToList();

            var byVendor = all
                .GroupBy(a => a.Vendor?.VendorName ?? "Unknown")
                .Select(g =>
                {
                    var act = g.Where(a => !a.IsCondemned).ToList();
                    var lw  = act.Sum(a => a.LiveWeight);
                    var hw  = act.Sum(a => a.HotWeight ?? 0);
                    var sc  = act.Sum(a => a.SaleCost);
                    return new VendorGroup
                    {
                        VendorName      = g.Key,
                        Count           = g.Count(),
                        Condemned       = g.Count(a => a.IsCondemned),
                        Passed          = act.Count,
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
                ByType           = new List<TypeSummaryRow>
                {
                    BuildRow("Cows",              sale.Where(a => IsCow(a)  && a.ProgramCode != "ABF")),
                    BuildRow("Bulls",             sale.Where(a => IsBull(a) && a.ProgramCode != "ABF")),
                    BuildRow("Cows ABF",          all.Where(a => IsCow(a)   && a.ProgramCode == "ABF")),
                    BuildRow("Bulls ABF",         all.Where(a => IsBull(a)  && a.ProgramCode == "ABF")),
                    BuildRow("Consignment Cows",  cons.Where(a => IsCow(a)  && a.Origin != "Canada")),
                    BuildRow("Consignment Bulls", cons.Where(a => IsBull(a) && a.Origin != "Canada")),
                    BuildRow("Canadian Cows",     cdn.Where(IsCow)),
                    BuildRow("Canadian Bulls",    cdn.Where(IsBull)),
                }
            };
        }

        // GetFilteredAsync - flexible export query 
        public async Task<IEnumerable<Animal>> GetFilteredAsync(ExportFilter f)
        {
            // Build WHERE clauses dynamically using safe raw SQL parameters
            var conditions = new List<string>();
            var parameters = new List<object>();
            int p = 0;

            // Status filter
            if (!string.IsNullOrEmpty(f.Status) && f.Status != "all")
            {
                conditions.Add($"a.KillStatus = {{{p++}}}");
                parameters.Add(f.Status);
            }
            else
            {
                // Exclude Flagged (soft-deleted) unless explicitly requested
                conditions.Add("a.KillStatus != 'Flagged'");
            }

            // Vendor filter
            if (f.VendorId.HasValue)
            {
                conditions.Add($"a.VendorID = {{{p++}}}");
                parameters.Add(f.VendorId.Value);
            }

            // Kill date range
            if (f.KillDateFrom.HasValue)
            {
                conditions.Add($"CAST(a.KillDate AS DATE) >= {{{p++}}}");
                parameters.Add(f.KillDateFrom.Value.ToString("yyyy-MM-dd"));
            }
            if (f.KillDateTo.HasValue)
            {
                conditions.Add($"CAST(a.KillDate AS DATE) <= {{{p++}}}");
                parameters.Add(f.KillDateTo.Value.ToString("yyyy-MM-dd"));
            }

            // Purchase date range
            if (f.PurchDateFrom.HasValue)
            {
                conditions.Add($"CAST(a.PurchaseDate AS DATE) >= {{{p++}}}");
                parameters.Add(f.PurchDateFrom.Value.ToString("yyyy-MM-dd"));
            }
            if (f.PurchDateTo.HasValue)
            {
                conditions.Add($"CAST(a.PurchaseDate AS DATE) <= {{{p++}}}");
                parameters.Add(f.PurchDateTo.Value.ToString("yyyy-MM-dd"));
            }

            var where = conditions.Any()
                ? "WHERE " + string.Join(" AND ", conditions)
                : "";

            var sql = $@"SELECT a.* FROM tbl_barn_animal_entry a {where} ORDER BY a.VendorID, a.PurchaseDate, a.ControlNo";

            var animals = await _db.Animals
                .FromSqlRaw(sql, parameters.ToArray())
                .ToListAsync();

            await AttachVendors(animals);
            return animals;
        }
        public async Task<IEnumerable<Animal>> GetByAnimalControlNumbersAsync(IEnumerable<string> acns)
        {
            var acnList = acns.Where(a => !string.IsNullOrWhiteSpace(a)).Distinct().ToList();
            if (!acnList.Any()) return Enumerable.Empty<Animal>();
            var animals = await _db.Animals
                .Where(a => a.AnimalControlNumber != null && acnList.Contains(a.AnimalControlNumber))
                .ToListAsync();
            await AttachVendors(animals);
            return animals;
        }

        public async Task<(int Updated, int Failed, List<string> Errors)> BulkUpdateHotWeightAsync(
            IEnumerable<HotWeightUpdateData> updates)
        {
            var updateList = updates.ToList();
            if (!updateList.Any()) return (0, 0, new List<string>());
            var ids = updateList.Select(u => u.ControlNo).Distinct().ToArray();
            var idList = string.Join(",", ids);
            var animals = await _db.Animals
                .FromSqlRaw($"SELECT * FROM tbl_barn_animal_entry WHERE ControlNo IN ({idList})")
                .ToListAsync();
            var dict = animals.ToDictionary(a => a.ControlNo);
            int updated = 0; int failed = 0;
            var errors = new List<string>();
            foreach (var upd in updateList)
            {
                if (!dict.TryGetValue(upd.ControlNo, out var animal))
                { errors.Add($"ACN {upd.ACN}: Record not found."); failed++; continue; }
                try
                {
                    if (upd.HotWeight.HasValue) animal.HotWeight = upd.HotWeight.Value;
                    if (!string.IsNullOrWhiteSpace(upd.Grade)) animal.Grade = upd.Grade.Trim().ToUpper();
                    if (upd.HealthScore.HasValue) animal.HealthScore = upd.HealthScore.Value;
                    animal.UpdatedAt = DateTime.Now;
                    updated++;
                }
                catch (Exception ex) { errors.Add($"ACN {upd.ACN}: {ex.Message}"); failed++; }
            }
            if (updated > 0) await _db.SaveChangesAsync();
            return (updated, failed, errors);
        }

    }
}