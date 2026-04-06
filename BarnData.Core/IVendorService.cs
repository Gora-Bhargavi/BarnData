using BarnData.Data.Entities;

namespace BarnData.Core.Services
{
    public interface IVendorService
    {
        Task<IEnumerable<Vendor>> GetAllActiveAsync();
        Task<Vendor?> GetByIdAsync(int vendorId);
        Task<int> GetOrCreateAsync(string vendorName);
    }
}
