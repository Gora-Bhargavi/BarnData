using Microsoft.AspNetCore.Mvc.Rendering;

namespace BarnData.Web.Models
{
    // Model for the reusable _VendorPicker partial view.
    public class VendorPickerModel
    {
        // Unique prefix on this page (e.g. "mk", "animal"). Used for DOM ids.
        public string Id { get; set; } = "vp";

        // Name attr for the hidden input that will submit the csv vendor ids.
        public string HiddenName { get; set; } = "vendorIds";

        // Comma-separated list of currently-selected vendor ids.
        public string? SelectedIds { get; set; }

        // Full list of vendors available. Pre-marked Selected=true where applicable.
        public IEnumerable<SelectListItem>? Vendors { get; set; }

        // Row count displayed next to "All vendors (N)".
        public int TotalCount { get; set; }

        // Small caption shown above the trigger (e.g. "Filter by vendor").
        public string Label { get; set; } = "Vendors";
    }
}
