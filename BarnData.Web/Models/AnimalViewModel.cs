using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace BarnData.Web.Models
{
    public class AnimalViewModel : IValidatableObject
    {
        public int ControlNo { get; set; }

        // ── Vendor ────────────────────────────────────────────────────────
        [Display(Name = "Vendor")]
        public int VendorID { get; set; }

        [Required(ErrorMessage = "Purchase type is required")]
        [MaxLength(50)]
        [Display(Name = "Purchase type")]
        public string PurchaseType { get; set; } = string.Empty;

        [Required(ErrorMessage = "Purchase date is required")]
        [DataType(DataType.Date)]
        [Display(Name = "Purchase date")]
        public DateTime PurchaseDate { get; set; } = DateTime.Today;

        [Required(ErrorMessage = "Live rate is required")]
        [Range(0.0001, 9999.9999, ErrorMessage = "Live rate must be greater than 0")]
        [Display(Name = "Live rate ($/head)")]
        public decimal LiveRate { get; set; }

        // ── Tags ──────────────────────────────────────────────────────────
        [Required(ErrorMessage = "Tag Number 1 is required")]
        [MaxLength(50)]
        [Display(Name = "Tag Number 1")]
        public string TagNumber1 { get; set; } = string.Empty;

        [MaxLength(50)]
        [Display(Name = "Tag Number 2")]
        public string? TagNumber2 { get; set; }

        [MaxLength(50)]
        [Display(Name = "Tag 3")]
        public string? Tag3 { get; set; }

        [Required(ErrorMessage = "Animal Control Number is required")]
        [MaxLength(50)]
        [Display(Name = "Animal control number")]
        public string AnimalControlNumber { get; set; } = string.Empty;

        // ── Animal classification ─────────────────────────────────────────
        [Required(ErrorMessage = "Animal type is required")]
        [Display(Name = "Animal type")]
        public string AnimalType { get; set; } = string.Empty;

        [Display(Name = "Animal type 2")]
        [MaxLength(50)]
        public string? AnimalType2 { get; set; }

        [Required(ErrorMessage = "Program code is required")]
        [Display(Name = "Program code")]
        public string ProgramCode { get; set; } = string.Empty;

        // ── Weight & kill ─────────────────────────────────────────────────
        [Required(ErrorMessage = "Live weight is required")]
        [Range(0.1, 9999.9, ErrorMessage = "Live weight must be greater than 0")]
        [Display(Name = "Live weight (lbs)")]
        public decimal LiveWeight { get; set; }

        [Required(ErrorMessage = "Kill date is required")]
        [DataType(DataType.Date)]
        [Display(Name = "Kill date")]
        public DateTime KillDate { get; set; } = DateTime.Today;

        [Display(Name = "Hot weight (lbs)")]
        [Range(0.1, 9999.9, ErrorMessage = "Hot weight must be greater than 0")]
        public decimal? HotWeight { get; set; }

        // ── Grading ───────────────────────────────────────────────────────
        [Required(ErrorMessage = "Grade is required")]
        [Display(Name = "Grade")]
        public string Grade { get; set; } = string.Empty;

        // Editable by pricing staff later — not auto-filled from HotScale
        [MaxLength(10)]
        [Display(Name = "Grade 2")]
        public string? Grade2 { get; set; }

        [Required(ErrorMessage = "Health score is required")]
        [Range(1, 3, ErrorMessage = "Health score must be 1, 2, or 3")]
        [Display(Name = "Health score")]
        public int HealthScore { get; set; }

        // ── Office / additional ───────────────────────────────────────────
        [Display(Name = "Fetal blood")]
        public decimal? FetalBlood { get; set; }

        [MaxLength(500)]
        [Display(Name = "Comment")]
        public string? Comment { get; set; }

        [MaxLength(2)]
        [Display(Name = "State")]
        public string? State { get; set; }

        [MaxLength(100)]
        [Display(Name = "Buyer name")]
        public string? BuyerName { get; set; }

        [MaxLength(100)]
        [Display(Name = "Vet name")]
        public string? VetName { get; set; }

        [MaxLength(200)]
        [Display(Name = "Office use 2")]
        public string? OfficeUse2 { get; set; }

        public string KillStatus { get; set; } = "Pending";

        // ── Free-text vendor name (used when vendor is new / not in list) ───
        public string? VendorNameFreeText { get; set; }

        // ── Dropdown source lists (populated by controller) ───────────────
        public IEnumerable<SelectListItem> VendorList { get; set; }
            = new List<SelectListItem>();

        public IEnumerable<SelectListItem> AnimalTypeList { get; set; }
            = new List<SelectListItem>
            {
                new("Bull",  "Bull"),
                new("Cow",   "Cow"),
                new("Steer", "Steer"),
            };

        public IEnumerable<SelectListItem> ProgramCodeList { get; set; }
            = new List<SelectListItem>
            {
                new("ABF",  "ABF"),
                new("ABNF", "ABNF"),
                new("NYFS", "NYFS"),
                new("REG",  "REG"),
                new("AGF",  "AGF"),
            };

        public IEnumerable<SelectListItem> GradeList { get; set; }
            = new List<SelectListItem>
            {
                new("CT", "CT"),
                new("B1", "B1"),
                new("B2", "B2"),
                new("CN", "CN"),
                new("LB", "LB"),
                new("UB", "UB"),
                new("BB", "BB"),
                new("SL", "SL"),
            };

        public IEnumerable<SelectListItem> PurchaseTypeList { get; set; }
            = new List<SelectListItem>
            {
                new("Sale bill",    "Sale bill"),
                new("Consignment",  "Consignment"),
            };

        // ── Weight warning flags ───────────────────────────────────────────
        public bool ShowWeightWarning { get; set; }

        // Set to true when user checks "I confirm this weight is correct"
        public bool WeightWarningConfirmed { get; set; }

        // ── Cross-field validation ─────────────────────────────────────────
        public IEnumerable<ValidationResult> Validate(ValidationContext context)
        {
            if (KillDate < PurchaseDate)
            {
                yield return new ValidationResult(
                    "Kill date cannot be before purchase date.",
                    new[] { nameof(KillDate) }
                );
            }
        }
    }
}
