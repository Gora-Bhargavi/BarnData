using BarnData.Core.Services;
using BarnData.Data.Entities;
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

    public class VendorAnimalsViewModel
    {
        public string VendorName { get; set; } = string.Empty;
        public IEnumerable<Animal> Animals { get; set;} = new List<Animal>();
    }
}
