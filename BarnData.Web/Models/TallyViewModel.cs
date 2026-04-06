using BarnData.Core.Services;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace BarnData.Web.Models
{
    public class TallyViewModel
    {
        public TallySummary Summary { get; set; } = new();
        public DateTime KillDate { get; set; } = DateTime.Today;
        public int? VendorId { get; set; }
        public bool IsPrint { get; set; }
        public IEnumerable<SelectListItem> VendorList { get; set; }
            = new List<SelectListItem>();
    }
}
