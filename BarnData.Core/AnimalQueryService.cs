/*
using BarnData.Data;
using BarnData.Data.Entities;
using Microsoft.EntityFrameworkCore;
using System.Data;

namespace BarnData.Core.Services
{
    // High-performance query service for the paginated Mark-as-Killed page.
    // Relies on IX_Animal_KillStatus_VendorID from BarnDataContext.
    public class AnimalQueryService : IAnimalQueryService
    {
        private readonly BarnDataContext _db;
        public AnimalQueryService(BarnDataContext db) => _db = db;

        public async Task<(IReadOnlyList<Animal> Items, int TotalCount)> GetPendingPagedAsync(
            IReadOnlyList<int>? vendorIds,
            int page,
            int pageSize,
            string? searchTerm = null)
        {
            if (page < 1) page = 1;
            if (pageSize < 1) pageSize = 100;
            if (pageSize > 1000) pageSize = 1000;   // hard cap to protect the server

            // Build a parameterised WHERE clause. We stick with FromSqlRaw positional
            // {0}, {1}, ... style because AnimalService.cs already uses that convention
            // and EF Core honours SQL Server parameter caching on it.
            var whereParts = new List<string> { "a.KillStatus = 'Pending'" };
            var parameters = new List<object>();
            int p = 0;

            if (vendorIds != null && vendorIds.Count > 0)
            {
                var placeholders = new List<string>();
                foreach (var id in vendorIds.Distinct())
                {
                    placeholders.Add("{" + p + "}");
                    parameters.Add(id);
                    p++;
                }
                whereParts.Add($"a.VendorID IN ({string.Join(",", placeholders)})");
            }

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                var trimmed = searchTerm.Trim();
                var pattern = "%" + trimmed + "%";

                // Tags and ACN use substring (LIKE %term%) so partial matches still work
                // (typical use case: typing "599" to find tag "23NP0599").
                // Ctrl No uses exact-equality match — without this guard, searching "344"
                // would also return Ctrl No 17344, 1344, 23440, etc.
                if (int.TryParse(trimmed, out var ctrlNoExact))
                {
                    whereParts.Add(
                        $"(a.TagNumber1 LIKE {{{p}}} OR a.TagNumber2 LIKE {{{p}}} " +
                        $"OR a.Tag3 LIKE {{{p}}} OR a.AnimalControlNumber LIKE {{{p}}} " +
                        $"OR a.ControlNo = {{{p + 1}}})");
                    parameters.Add(pattern);
                    parameters.Add(ctrlNoExact);
                    p += 2;
                }
                else
                {
                    // Non-numeric search term — Ctrl No can't match, skip that branch.
                    whereParts.Add(
                        $"(a.TagNumber1 LIKE {{{p}}} OR a.TagNumber2 LIKE {{{p}}} " +
                        $"OR a.Tag3 LIKE {{{p}}} OR a.AnimalControlNumber LIKE {{{p}}})");
                    parameters.Add(pattern);
                    p++;
                }
            }

            var whereClause = "WHERE " + string.Join(" AND ", whereParts);

            //  Count query (uses the same parameter set, translated to named params) 
            var countSql = $"SELECT COUNT(*) FROM tbl_barn_animal_entry a {whereClause}";
            int total;
            using (var cmd = _db.Database.GetDbConnection().CreateCommand())
            {
                var translated = countSql;
                for (int i = 0; i < parameters.Count; i++)
                    translated = translated.Replace("{" + i + "}", "@p" + i);
                cmd.CommandText = translated;

                for (int i = 0; i < parameters.Count; i++)
                {
                    var prm = cmd.CreateParameter();
                    prm.ParameterName = "@p" + i;
                    prm.Value = parameters[i];
                    cmd.Parameters.Add(prm);
                }
                if (cmd.Connection!.State != ConnectionState.Open)
                    await cmd.Connection.OpenAsync();
                total = Convert.ToInt32(await cmd.ExecuteScalarAsync() ?? 0);
            }

            if (total == 0) return (Array.Empty<Animal>(), 0);

            //  Paged data query 
            var offset = (page - 1) * pageSize;
            var dataSql =
                $"SELECT a.* FROM tbl_barn_animal_entry a {whereClause} " +
                "ORDER BY a.VendorID, a.PurchaseDate, a.ControlNo " +
                $"OFFSET {{{p}}} ROWS FETCH NEXT {{{p + 1}}} ROWS ONLY";

            var dataParams = new List<object>(parameters) { offset, pageSize };

            var animals = await _db.Animals
                .FromSqlRaw(dataSql, dataParams.ToArray())
                .AsNoTracking()
                .ToListAsync();

            // One-shot vendor hydration: single query, no N+1.
            if (animals.Count > 0)
            {
                var ids = animals.Select(a => a.VendorID).Distinct().ToList();
                var idList = string.Join(",", ids);
                var vendors = await _db.Vendors
                    .FromSqlRaw($"SELECT * FROM tbl_vendor_master WHERE VendorID IN ({idList})")
                    .AsNoTracking()
                    .ToListAsync();
                var dict = vendors.ToDictionary(v => v.VendorID);
                foreach (var a in animals)
                    if (dict.TryGetValue(a.VendorID, out var v)) a.Vendor = v;
            }

            return (animals, total);
        }

        public async Task<IReadOnlyList<VendorPickItem>> GetVendorPickListAsync()
        {
            // Projection-only query — small payload, benefits from IX_Vendor_Active_Name.
            var list = await _db.Vendors
                .FromSqlRaw("SELECT * FROM tbl_vendor_master WHERE IsActive = 1 ORDER BY VendorName")
                .AsNoTracking()
                .Select(v => new VendorPickItem { VendorID = v.VendorID, VendorName = v.VendorName })
                .ToListAsync();

            return list;
        }
    }
}
*/

using BarnData.Data;
using BarnData.Data.Entities;
using Microsoft.EntityFrameworkCore;
using System.Data;

namespace BarnData.Core.Services
{
    // High-performance query service for the paginated Mark-as-Killed page.
    // Relies on IX_Animal_KillStatus_VendorID from BarnDataContext.
    public class AnimalQueryService : IAnimalQueryService
    {
        private readonly BarnDataContext _db;
        public AnimalQueryService(BarnDataContext db) => _db = db;

        public async Task<(IReadOnlyList<Animal> Items, int TotalCount)> GetPendingPagedAsync(
            IReadOnlyList<int>? vendorIds,
            int page,
            int pageSize,
            string? searchTerm = null)
        {
            if (page < 1) page = 1;
            if (pageSize < 1) pageSize = 100;
            if (pageSize > 1000) pageSize = 1000;   // hard cap to protect the server

            // Build a parameterised WHERE clause. We stick with FromSqlRaw positional
            // {0}, {1}, ... style because AnimalService.cs already uses that convention
            // and EF Core honours SQL Server parameter caching on it.
            var whereParts = new List<string> { "a.KillStatus = 'Pending'" };
            var parameters = new List<object>();
            int p = 0;

            if (vendorIds != null && vendorIds.Count > 0)
            {
                var placeholders = new List<string>();
                foreach (var id in vendorIds.Distinct())
                {
                    placeholders.Add("{" + p + "}");
                    parameters.Add(id);
                    p++;
                }
                whereParts.Add($"a.VendorID IN ({string.Join(",", placeholders)})");
            }

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                var trimmed = searchTerm.Trim();
                var pattern = "%" + trimmed + "%";

                // Tags and ACN use substring (LIKE %term%) so partial matches still work
                // (typical use case: typing "599" to find tag "23NP0599").
                // Ctrl No uses exact-equality match — without this guard, searching "344"
                // would also return Ctrl No 17344, 1344, 23440, etc.
                if (int.TryParse(trimmed, out var ctrlNoExact))
                {
                    whereParts.Add(
                        $"(a.TagNumber1 LIKE {{{p}}} OR a.TagNumber2 LIKE {{{p}}} " +
                        $"OR a.Tag3 LIKE {{{p}}} OR a.AnimalControlNumber LIKE {{{p}}} " +
                        $"OR a.ControlNo = {{{p + 1}}})");
                    parameters.Add(pattern);
                    parameters.Add(ctrlNoExact);
                    p += 2;
                }
                else
                {
                    // Non-numeric search term — Ctrl No can't match, skip that branch.
                    whereParts.Add(
                        $"(a.TagNumber1 LIKE {{{p}}} OR a.TagNumber2 LIKE {{{p}}} " +
                        $"OR a.Tag3 LIKE {{{p}}} OR a.AnimalControlNumber LIKE {{{p}}})");
                    parameters.Add(pattern);
                    p++;
                }
            }

            var whereClause = "WHERE " + string.Join(" AND ", whereParts);

            //  Count query (uses the same parameter set, translated to named params) 
            var countSql = $"SELECT COUNT(*) FROM tbl_barn_animal_entry a {whereClause}";
            int total;
            using (var cmd = _db.Database.GetDbConnection().CreateCommand())
            {
                var translated = countSql;
                for (int i = 0; i < parameters.Count; i++)
                    translated = translated.Replace("{" + i + "}", "@p" + i);
                cmd.CommandText = translated;

                for (int i = 0; i < parameters.Count; i++)
                {
                    var prm = cmd.CreateParameter();
                    prm.ParameterName = "@p" + i;
                    prm.Value = parameters[i];
                    cmd.Parameters.Add(prm);
                }
                if (cmd.Connection!.State != ConnectionState.Open)
                    await cmd.Connection.OpenAsync();
                total = Convert.ToInt32(await cmd.ExecuteScalarAsync() ?? 0);
            }

            if (total == 0) return (Array.Empty<Animal>(), 0);

            //  Paged data query 
            var offset = (page - 1) * pageSize;
            var dataSql =
                $"SELECT a.* FROM tbl_barn_animal_entry a {whereClause} " +
                "ORDER BY a.VendorID, a.PurchaseDate, a.ControlNo " +
                $"OFFSET {{{p}}} ROWS FETCH NEXT {{{p + 1}}} ROWS ONLY";

            var dataParams = new List<object>(parameters) { offset, pageSize };

            var animals = await _db.Animals
                .FromSqlRaw(dataSql, dataParams.ToArray())
                .AsNoTracking()
                .ToListAsync();

            // One-shot vendor hydration: single query, no N+1.
            if (animals.Count > 0)
            {
                var ids = animals.Select(a => a.VendorID).Distinct().ToList();
                var idList = string.Join(",", ids);
                var vendors = await _db.Vendors
                    .FromSqlRaw($"SELECT * FROM tbl_vendor_master WHERE VendorID IN ({idList})")
                    .AsNoTracking()
                    .ToListAsync();
                var dict = vendors.ToDictionary(v => v.VendorID);
                foreach (var a in animals)
                    if (dict.TryGetValue(a.VendorID, out var v)) a.Vendor = v;
            }

            return (animals, total);
        }

        public async Task<IReadOnlyList<VendorPickItem>> GetVendorPickListAsync()
        {
            // Projection-only query — small payload, benefits from IX_Vendor_Active_Name.
            var list = await _db.Vendors
                .FromSqlRaw("SELECT * FROM tbl_vendor_master WHERE IsActive = 1 ORDER BY VendorName")
                .AsNoTracking()
                .Select(v => new VendorPickItem { VendorID = v.VendorID, VendorName = v.VendorName })
                .ToListAsync();

            return list;
        }
    }
}
