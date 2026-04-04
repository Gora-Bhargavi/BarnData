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
    }
}
