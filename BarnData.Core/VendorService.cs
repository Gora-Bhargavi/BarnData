using BarnData.Data;
using BarnData.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace BarnData.Core.Services
{
    public class VendorService : IVendorService
    {
        private readonly BarnDataContext _db;

        public VendorService(BarnDataContext db)
        {
            _db = db;
        }

        public async Task<IEnumerable<Vendor>> GetAllActiveAsync()
        {
            return await _db.Vendors
                .Where(v => v.IsActive)
                .OrderBy(v => v.VendorName)
                .ToListAsync();
        }

        public async Task<Vendor?> GetByIdAsync(int vendorId)
        {
            return await _db.Vendors.FindAsync(vendorId);
        }

        // Find vendor by name or create a new one — used when typing a new vendor
        public async Task<int> GetOrCreateAsync(string vendorName)
        {
            var existing = await _db.Vendors
                .FirstOrDefaultAsync(v => v.VendorName == vendorName);

            if (existing != null)
                return existing.VendorID;

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
