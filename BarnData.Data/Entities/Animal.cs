using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace BarnData.Data.Entities
{
    [Table("tbl_barn_animal_entry")]
    public class Animal
    {
        //  Identity
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ControlNo { get; set; }

        [Required]
        public int VendorID { get; set; }

        //  Tags 
        [Required]
        [MaxLength(50)]
        public string TagNumber1 { get; set; } = string.Empty;

        [MaxLength(50)]
        public string? TagNumber2 { get; set; }

        [MaxLength(50)]
        public string? Tag3 { get; set; }

        //  Animal classification 
        [Required]
        [MaxLength(50)]
        public string AnimalType { get; set; } = string.Empty;

        [MaxLength(50)]
        public string? AnimalType2 { get; set; }

        [Required]
        [MaxLength(20)]
        public string ProgramCode { get; set; } = "REG";

        // Purchase info 
        [Required]
        public DateTime PurchaseDate { get; set; }

        [Required]
        [MaxLength(50)]
        public string PurchaseType { get; set; } = string.Empty;

        [Column(TypeName = "decimal(8,1)")]
        public decimal LiveWeight { get; set; }

        // Sale Bill: price per lb of live weight
        [Column(TypeName = "decimal(10,4)")]
        public decimal LiveRate { get; set; }

        // Consignment Bill: price per lb of hot weight
        [Column(TypeName = "decimal(10,4)")]
        public decimal? ConsignmentRate { get; set; }

        //  Kill data — NULL until animal is actually killed ─
        public DateTime? KillDate { get; set; }

        [Column(TypeName = "decimal(8,1)")]
        public decimal? HotWeight { get; set; }

        // Grading — filled from scale ticket / HotScale 
        [MaxLength(10)]
        public string? Grade { get; set; }

        [MaxLength(10)]
        public string? Grade2 { get; set; }

        public int? HealthScore { get; set; }

        //  Additional fields 
        [Column(TypeName = "decimal(6,2)")]
        public decimal? FetalBlood { get; set; }

        [MaxLength(500)]
        public string? Comment { get; set; }

        [MaxLength(50)]
        public string? AnimalControlNumber { get; set; }

        //  Office / consignment fields 
        [MaxLength(2)]
        public string? State { get; set; }

        [MaxLength(100)]
        public string? BuyerName { get; set; }

        [MaxLength(100)]
        public string? VetName { get; set; }

        [MaxLength(200)]
        public string? OfficeUse2 { get; set; }

        // Origin: Farmer, Canada — for consignment animals
        [MaxLength(20)]
        public string? Origin { get; set; }

        // Condemned: excluded from weight/cost totals
        public bool IsCondemned { get; set; } = false;

        // Reference to imported sale bill batch
        [MaxLength(100)]
        public string? SaleBillRef { get; set; }

        //  System fields 
        [Required]
        [MaxLength(20)]
        public string KillStatus { get; set; } = "Pending";

        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime? UpdatedAt { get; set; }

        [MaxLength(50)]
        public string? CreatedBy { get; set; }

        //  Calculated properties (not stored in DB)
        [NotMapped]
        public decimal SaleCost =>
            PurchaseType == "Consignment"
                ? (HotWeight ?? 0) * (ConsignmentRate ?? 0)
                : LiveWeight * LiveRate;

        [NotMapped]
        public decimal YieldPct =>
            PurchaseType == "Consignment"
                ? 100m
                : (HotWeight.HasValue && LiveWeight > 0)
                    ? Math.Round(HotWeight.Value / LiveWeight * 100, 2)
                    : 0;

        [NotMapped]
        public decimal DressRate =>
            HotWeight.HasValue && HotWeight.Value > 0
                ? Math.Round(SaleCost / HotWeight.Value, 3)
                : 0;

        //  Navigation 
        [ForeignKey("VendorID")]
        public Vendor? Vendor { get; set; }
    }
}
