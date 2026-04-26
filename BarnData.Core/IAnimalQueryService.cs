using BarnData.Data.Entities;

namespace BarnData.Core.Services
{
    // Performance-focused paged queries. Bolts on beside the existing IAnimalService
    // without modifying it — existing code paths keep working unchanged.
    // Used by the new paginated Mark-as-Killed view.
    public interface IAnimalQueryService
    {
        // Paged pending animals with optional vendor filter and free-text search.
        // Uses IX_Animal_KillStatus_VendorID for an index seek.
        // Ordered by VendorID, PurchaseDate, ControlNo for stable pagination.
        Task<(IReadOnlyList<Animal> Items, int TotalCount)> GetPendingPagedAsync(
            IReadOnlyList<int>? vendorIds,
            int page,
            int pageSize,
            string? searchTerm = null);

        // Compact vendor list (id + name only) for the picker.
        // Uses IX_Vendor_Active_Name.
        Task<IReadOnlyList<VendorPickItem>> GetVendorPickListAsync();
    }

    // Lightweight DTO to avoid shipping full Vendor rows to the UI.
    public class VendorPickItem
    {
        public int    VendorID   { get; set; }
        public string VendorName { get; set; } = string.Empty;
    }
}
