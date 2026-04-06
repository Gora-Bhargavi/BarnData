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

        [Required]
        [MaxLength(50)]
        public string TagNumber2 { get; set; } = string.Empty;

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
        public string ProgramCode { get; set; } = string.Empty;

        //  Purchase info 
        [Required]
        public DateTime PurchaseDate { get; set; }

        [Required]
        [MaxLength(50)]
        public string PurchaseType { get; set; } = string.Empty;

        [Required]
        [Column(TypeName = "decimal(8,1)")]
        public decimal LiveWeight { get; set; }

        [Required]
        [Column(TypeName = "decimal(10,4)")]
        public decimal LiveRate { get; set; }

        //  Kill data 
        [Required]
        public DateTime KillDate { get; set; }

        [Column(TypeName = "decimal(8,1)")]
        public decimal? HotWeight { get; set; }

        //  Grading 
        [Required]
        [MaxLength(10)]
        public string Grade { get; set; } = string.Empty;

        // Entered LATER by pricing staff - not from HotScale
        [MaxLength(10)]
        public string? Grade2 { get; set; }

        [Required]
        public int HealthScore { get; set; }

        //  Additional fields 
        [Column(TypeName = "decimal(6,2)")]
        public decimal? FetalBlood { get; set; }

        [Required]
        [MaxLength(500)]
        public string Comment { get; set; } = string.Empty;

        [Required]
        [MaxLength(50)]
        public string AnimalControlNumber { get; set; } = string.Empty;

        //  Office / consignment fields 
        [MaxLength(2)]
        public string? State { get; set; }

        [MaxLength(100)]
        public string? BuyerName { get; set; }

        [MaxLength(100)]
        public string? VetName { get; set; }

        [MaxLength(200)]
        public string? OfficeUse2 { get; set; }

        // System fields 
        [Required]
        [MaxLength(20)]
        public string KillStatus { get; set; } = "Pending";

        public DateTime CreatedAt { get; set; } = DateTime.Now;

        public DateTime? UpdatedAt { get; set; }

        [MaxLength(50)]
        public string? CreatedBy { get; set; }

        // Navigation property 
        [ForeignKey("VendorID")]
        public Vendor? Vendor { get; set; }
    }
}
