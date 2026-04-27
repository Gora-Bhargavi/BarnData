using BarnData.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace BarnData.Data
{
    public class BarnDataContext : DbContext
    {
        public BarnDataContext(DbContextOptions<BarnDataContext> options)
            : base(options) { }

        public DbSet<Animal> Animals { get; set; } = null!;
        public DbSet<Vendor> Vendors { get; set; } = null!;

        // Persistent import staging (replaces HttpContext.Session for Excel + HW imports)
        public DbSet<ImportStagingBatch> ImportStagingBatches { get; set; } = null!;
        public DbSet<ImportStagingRow>   ImportStagingRows    { get; set; } = null!;

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            //  Vendor
            modelBuilder.Entity<Vendor>(entity =>
            {
                entity.ToTable("tbl_vendor_master");
                entity.HasKey(v => v.VendorID);
                entity.Property(v => v.VendorName).IsRequired().HasMaxLength(150);

                entity.HasMany(v => v.Animals)
                    .WithOne(a => a.Vendor)
                    .HasForeignKey(a => a.VendorID)
                    .OnDelete(DeleteBehavior.Restrict);

                // Speeds up vendor dropdown load (WHERE IsActive=1 ORDER BY VendorName)
                entity.HasIndex(v => new { v.IsActive, v.VendorName })
                      .HasDatabaseName("IX_Vendor_Active_Name");
            });

            //  Animal
            modelBuilder.Entity<Animal>(entity =>
            {
                entity.ToTable("tbl_barn_animal_entry");
                entity.HasKey(a => a.ControlNo);

                entity.Property(a => a.LiveWeight).HasColumnType("decimal(8,1)");
                entity.Property(a => a.LiveRate).HasColumnType("decimal(10,4)");
                entity.Property(a => a.HotWeight).HasColumnType("decimal(8,1)");
                entity.Property(a => a.FetalBlood).HasColumnType("decimal(6,2)");
                entity.Property(a => a.ConsignmentRate).HasColumnType("decimal(10,4)");

                // FK: animal -> vendor (kept commented out to match current behavior)
                //entity.HasOne(a => a.Vendor)
                //    .WithMany(v => v.Animals)
                //    .HasForeignKey(a => a.VendorID)
                //    .OnDelete(DeleteBehavior.Restrict);

                //  PERFORMANCE INDEXES (the main fix for Issue #1) 
                // These drop Mark-as-Killed load time from ~50s to <2s once rowcount > ~10k.
                // Composite index matches the most common WHERE:
                //    KillStatus='Pending' AND VendorID IN (...)
                entity.HasIndex(a => new { a.KillStatus, a.VendorID })
                      .HasDatabaseName("IX_Animal_KillStatus_VendorID");

                // For Tally / Killed-by-date queries.
                entity.HasIndex(a => new { a.KillStatus, a.KillDate })
                      .HasDatabaseName("IX_Animal_KillStatus_KillDate");

                // For tag lookups (HW import + search).
                entity.HasIndex(a => a.TagNumber1).HasDatabaseName("IX_Animal_Tag1");
                entity.HasIndex(a => a.AnimalControlNumber).HasDatabaseName("IX_Animal_ACN");
            });

            //  Import staging
            modelBuilder.Entity<ImportStagingBatch>(entity =>
            {
                entity.ToTable("tbl_import_staging_batch");
                entity.HasKey(b => b.BatchID);

                // Finds the latest Active batch for this user + type
                entity.HasIndex(b => new { b.BatchType, b.Status, b.CreatedBy })
                      .HasDatabaseName("IX_StagingBatch_Type_Status_User");
            });

            modelBuilder.Entity<ImportStagingRow>(entity =>
            {
                entity.ToTable("tbl_import_staging_row");
                entity.HasKey(r => r.RowID);

                entity.HasOne(r => r.Batch)
                      .WithMany()
                      .HasForeignKey(r => r.BatchID)
                      .OnDelete(DeleteBehavior.Cascade);

                // Filtering rows for a batch by status (tab switching)
                entity.HasIndex(r => new { r.BatchID, r.Status })
                      .HasDatabaseName("IX_StagingRow_Batch_Status");
            });
        }
    }
}
