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

                // FK: animal → vendor
                //entity.HasOne(a => a.Vendor)
                  //  .WithMany(v => v.Animals)
                   // .HasForeignKey(a => a.VendorID)
                    //.OnDelete(DeleteBehavior.Restrict);
            });
        }
    }
}
