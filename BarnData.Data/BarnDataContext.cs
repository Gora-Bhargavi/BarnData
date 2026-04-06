using BarnData.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace BarnData.Data
{
    public class BarnDataContext : DbContext
    {
        public BarnDataContext(DbContextOptions<BarnDataContext> options)
            : base(options) { }

        //  DbSets 
        public DbSet<Animal> Animals { get; set; } = null!;
        public DbSet<Vendor> Vendors { get; set; } = null!;

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // tbl_vendor_master 
            modelBuilder.Entity<Vendor>(entity =>
            {
                entity.ToTable("tbl_vendor_master");

                entity.HasKey(v => v.VendorID);

                entity.Property(v => v.VendorName)
                    .IsRequired()
                    .HasMaxLength(150);

                entity.Property(v => v.IsActive)
                    .HasDefaultValue(true);

                entity.Property(v => v.CreatedAt)
                    .HasDefaultValueSql("GETDATE()");

                // Index: fast vendor dropdown lookups
                entity.HasIndex(v => v.VendorName)
                    .HasDatabaseName("IX_vendor_master_Name");
            });

            // tbl_animal_master 
            modelBuilder.Entity<Animal>(entity =>
            {
                entity.ToTable("tbl_barn_animal_entry");

                entity.HasKey(a => a.ControlNo);

                entity.Property(a => a.LiveWeight)
                    .HasColumnType("decimal(8,1)");

                entity.Property(a => a.LiveRate)
                    .HasColumnType("decimal(10,4)");

                entity.Property(a => a.HotWeight)
                    .HasColumnType("decimal(8,1)");

                entity.Property(a => a.FetalBlood)
                    .HasColumnType("decimal(6,2)");

                entity.Property(a => a.KillStatus)
                    .HasDefaultValue("Pending");

                entity.Property(a => a.CreatedAt)
                    .HasDefaultValueSql("GETDATE()");

                // Unique index: one Tag1 per vendor per kill date
                entity.HasIndex(a => new { a.TagNumber1, a.KillDate, a.VendorID })
                    .IsUnique()
                    .HasDatabaseName("UIX_animal_Tag1_KillDate_Vendor");

                // Index: fast tally report queries by kill date
                entity.HasIndex(a => new { a.KillDate, a.VendorID })
                    .HasDatabaseName("IX_animal_KillDate");

                // Index: future HotScale tag matching
                entity.HasIndex(a => new { a.TagNumber1, a.TagNumber2 })
                    .HasDatabaseName("IX_animal_Tags");

                // FK: animal → vendor
                entity.HasOne(a => a.Vendor)
                    .WithMany(v => v.Animals)
                    .HasForeignKey(a => a.VendorID)
                    .OnDelete(DeleteBehavior.Restrict);

                // DB-level check: kill date cannot be before purchase date
                entity.ToTable(t => t.HasCheckConstraint(
                    "CHK_animal_KillDate",
                    "[KillDate] >= [PurchaseDate]"
                ));

                // DB-level check: valid KillStatus values
                entity.ToTable(t => t.HasCheckConstraint(
                    "CHK_animal_KillStatus",
                    "[KillStatus] IN ('Pending','Killed','Verified','Flagged')"
                ));
            });
        }
    }
}
