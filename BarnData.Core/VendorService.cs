using BarnData.Data;
using BarnData.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace BarnData.Core.Services
{
    public class VendorService : IVendorService
    {
        private readonly BarnDataContext _db;

        public VendorService(BarnDataContext db) => _db = db;

        public async Task<IEnumerable<Vendor>> GetAllActiveAsync()
        {
            // Raw SQL — avoids EF Core 8 bool/string LINQ translation issues
            return await _db.Vendors
                .FromSqlRaw("SELECT * FROM tbl_vendor_master WHERE IsActive = 1 ORDER BY VendorName")
                .ToListAsync();
        }

        public async Task<Vendor?> GetByIdAsync(int vendorId)
        {
            return await _db.Vendors.FindAsync(vendorId);
        }

        public async Task<int> GetOrCreateAsync(string vendorName)
        {
            var existing = await _db.Vendors
                .FromSqlRaw("SELECT * FROM tbl_vendor_master WHERE VendorName = {0}", vendorName)
                .FirstOrDefaultAsync();

            if (existing != null) return existing.VendorID;

            var newVendor = new Vendor
            {
                VendorName = vendorName,
                IsActive   = true,
                CreatedAt  = DateTime.Now
            };

            _db.Vendors.Add(newVendor);
            await _db.SaveChangesAsync();
            return newVendor.VendorID;
        }
    }
}
