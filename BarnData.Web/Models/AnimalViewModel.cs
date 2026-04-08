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
        public string? VendorNameFreeText { get; set; }

        [Required(ErrorMessage = "Purchase type is required")]
        [Display(Name = "Purchase type")]
        public string PurchaseType { get; set; } = string.Empty;

        [Required(ErrorMessage = "Purchase date is required")]
        [DataType(DataType.Date)]
        [Display(Name = "Purchase date")]
        public DateTime PurchaseDate { get; set; } = DateTime.Today;

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

        [MaxLength(50)]
        [Display(Name = "Animal control number")]
        public string? AnimalControlNumber { get; set; }

        // ── Animal ────────────────────────────────────────────────────────
        [Required(ErrorMessage = "Animal type is required")]
        [Display(Name = "Animal type")]
        public string AnimalType { get; set; } = string.Empty;

        [Display(Name = "Animal type 2")]
        public string? AnimalType2 { get; set; }

        [Display(Name = "Program code")]
        public string ProgramCode { get; set; } = "REG";

        // ── Kill date — nullable, set on kill day ─────────────────────────
        [DataType(DataType.Date)]
        [Display(Name = "Kill date")]
        public DateTime? KillDate { get; set; }

        // ── Sale Bill fields ──────────────────────────────────────────────
        [Range(0, 9999.9, ErrorMessage = "Live weight must be 0 or more")]
        [Display(Name = "Live weight (lbs)")]
        public decimal LiveWeight { get; set; }

        [Display(Name = "Live rate ($/lb)")]
        public decimal LiveRate { get; set; }

        // ── Consignment Bill fields ───────────────────────────────────────
        [Display(Name = "Consignment rate ($/lb hot wt)")]
        public decimal? ConsignmentRate { get; set; }

        // ── Post-kill fields (filled from scale ticket / HotScale) ────────
        [Display(Name = "Hot weight (lbs)")]
        public decimal? HotWeight { get; set; }

        [Display(Name = "Grade")]
        public string? Grade { get; set; }

        [Display(Name = "Grade 2")]
        public string? Grade2 { get; set; }

        [Range(1, 5, ErrorMessage = "Health score must be 1–5")]
        [Display(Name = "Health score")]
        public int? HealthScore { get; set; }

        // ── Office ────────────────────────────────────────────────────────
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

        [Display(Name = "Origin")]
        public string? Origin { get; set; }

        [Display(Name = "Condemned")]
        public bool IsCondemned { get; set; } = false;

        public string KillStatus { get; set; } = "Pending";
        public bool ShowWeightWarning { get; set; }
        public bool WeightWarningConfirmed { get; set; }

        // ── Dropdown lists ────────────────────────────────────────────────
        public IEnumerable<SelectListItem> VendorList { get; set; } = new List<SelectListItem>();

        public IEnumerable<SelectListItem> AnimalTypeList { get; set; } = new List<SelectListItem>
        {
            new("Cow",   "Cow"),
            new("Bull",  "Bull"),
            new("Steer", "Steer → counts as Bull in tally"),
            new("Heifer","Heifer"),
        };

        public IEnumerable<SelectListItem> ProgramCodeList { get; set; } = new List<SelectListItem>
        {
            new("REG",  "REG"),
            new("ABF",  "ABF"),
            new("NYFS", "NYFS"),
            new("AGF",  "AGF"),
        };

        public IEnumerable<SelectListItem> GradeList { get; set; } = new List<SelectListItem>
        {
            new("CT", "CT"), new("B1", "B1"), new("B2", "B2"),
            new("CN", "CN"), new("LB", "LB"), new("UB", "UB"),
            new("BB", "BB"), new("BR", "BR"), new("SL", "SL"),
        };

        public IEnumerable<SelectListItem> PurchaseTypeList { get; set; } = new List<SelectListItem>
        {
            new("Sale Bill",        "Sale Bill"),
            new("Consignment Bill", "Consignment Bill"),
        };

        public IEnumerable<SelectListItem> OriginList { get; set; } = new List<SelectListItem>
        {
            new("",       "— Not applicable —"),
            new("Farmer", "Farmer"),
            new("Canada", "Canada"),
        };

        public IEnumerable<ValidationResult> Validate(ValidationContext context)
        {
            if (KillDate.HasValue && KillDate.Value < PurchaseDate)
                yield return new ValidationResult(
                    "Kill date cannot be before purchase date.",
                    new[] { nameof(KillDate) });

            // Live rate required for sale bill only if entering manually
            // (not required on import — can be 0 temporarily)
            if (PurchaseType == "Sale Bill" && LiveRate < 0)
                yield return new ValidationResult(
                    "Live rate cannot be negative.",
                    new[] { nameof(LiveRate) });

            // Consignment rate and Origin are NOT required at entry time
            // They may come from HotScale later (Phase 5)
        }
    }

    // ── Bulk import view model ─────────────────────────────────────────────
    public class SaleBillImportViewModel
    {
        public string? ImportedFile { get; set; }
        public int TotalRows { get; set; }
        public int Imported { get; set; }
        public int Skipped { get; set; }
        public List<string> Errors { get; set; } = new();
        public List<SaleBillPreviewRow> Preview { get; set; } = new();
    }

    public class SaleBillPreviewRow
    {
        public string VendorName  { get; set; } = string.Empty;
        public string Tag1        { get; set; } = string.Empty;
        public string? Tag2       { get; set; }
        public string AnimalType  { get; set; } = string.Empty;
        public decimal LiveWeight { get; set; }
        public decimal LiveRate   { get; set; }
        public decimal? ConsRate  { get; set; }
        public string PurchaseType{ get; set; } = string.Empty;
        public string? Comment    { get; set; }
        public bool IsCondemned   { get; set; }
        public string Status      { get; set; } = "OK"; // OK / Skip / Error
    }

    // ── Mark as killed view model ─────────────────────────────────────────
    public class MarkKilledViewModel
    {
        [Required]
        [DataType(DataType.Date)]
        public DateTime KillDate { get; set; } = DateTime.Today;
        public int? VendorId { get; set; }
        public List<PendingAnimalRow> Animals { get; set; } = new();
        public IEnumerable<SelectListItem> VendorList { get; set; } = new List<SelectListItem>();
    }

    public class PendingAnimalRow
    {
        public int      ControlNo    { get; set; }
        public string   VendorName   { get; set; } = string.Empty;
        public string   Tag1         { get; set; } = string.Empty;
        public string?  Tag2         { get; set; }
        public string?  Tag3         { get; set; }
        public string   AnimalType   { get; set; } = string.Empty;
        public string?  AnimalType2  { get; set; }
        public decimal  LiveWeight   { get; set; }
        public decimal  LiveRate     { get; set; }
        public string   PurchaseType { get; set; } = string.Empty;
        public DateTime PurchaseDate { get; set; }
        public string?  AnimalControlNumber { get; set; }
        public string?  Comment      { get; set; }
        public string?  State        { get; set; }
        public string?  BuyerName    { get; set; }
        public string?  VetName      { get; set; }
        public string?  OfficeUse2   { get; set; }
        public string   ProgramCode  { get; set; } = string.Empty;
        public bool     Selected     { get; set; }
        // Editable kill fields
        public decimal? HotWeight    { get; set; }
        public string?  Grade        { get; set; }
        public int?     HealthScore  { get; set; }
        public bool     IsCondemned  { get; set; }
    }

    // ── Excel Import view models ───────────────────────────────────────────
    public class ExcelImportViewModel
    {
        public string? FileName { get; set; }
        public int TotalRows   { get; set; }
        public List<ExcelPreviewRow> Rows { get; set; } = new();
        public List<string> Errors        { get; set; } = new();
    }

    public class ExcelPreviewRow
    {
        public int      RowNum             { get; set; }
        public string   VendorName         { get; set; } = string.Empty;
        public string   TagNumber1         { get; set; } = string.Empty;
        public string?  TagNumber2         { get; set; }
        public string?  Tag3               { get; set; }
        public string   AnimalType         { get; set; } = string.Empty;
        public string?  AnimalType2        { get; set; }
        public string   PurchaseType       { get; set; } = string.Empty;
        public DateTime PurchaseDate       { get; set; }
        public decimal  LiveWeight         { get; set; }
        public decimal  LiveRate           { get; set; }
        public DateTime? KillDate          { get; set; }
        public decimal? HotWeight          { get; set; }
        public string?  Grade              { get; set; }
        public int?     HealthScore        { get; set; }
        public string?  Comment            { get; set; }
        public string?  AnimalControlNumber{ get; set; }
        public string?  OfficeUse2         { get; set; }
        public string?  State              { get; set; }
        public string?  BuyerName          { get; set; }
        public string?  VetName            { get; set; }
        public bool     IsCondemned        { get; set; }
        public string   Status             { get; set; } = "OK"; // OK / Duplicate / Error
        public string?  StatusNote         { get; set; }
    }
}
