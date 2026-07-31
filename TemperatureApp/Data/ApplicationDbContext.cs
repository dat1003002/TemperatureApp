using Microsoft.EntityFrameworkCore;
using TemperatureApp.Models;

namespace TemperatureApp.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<BTLH3> BTLH3 { get; set; }
        public DbSet<BTLH4> BTLH4 { get; set; }
        public DbSet<BTLH5> BTLH5 { get; set; }
        public DbSet<MS_1> MS_1 { get; set; }
        public DbSet<MS_2> MS_2 { get; set; }
        public DbSet<MS_3> MS_3 { get; set; }
        public DbSet<MS_4> MS_4 { get; set; }
        public DbSet<MS_5> MS_5 { get; set; }
        public DbSet<VIS_MHE1_1> VIS_MHE1_1 { get; set; }
        public DbSet<VIS_MHE1_2> VIS_MHE1_2 { get; set; }
        public DbSet<VIS_MHE2_1> VIS_MHE2_1 { get; set; }
        public DbSet<VIS_MHE2_2> VIS_MHE2_2 { get; set; }
        public DbSet<VIS_MHE3_1> VIS_MHE3_1 { get; set; }
        public DbSet<VIS_MHE3_2> VIS_MHE3_2 { get; set; }
        public DbSet<CBstellcord_1> CBstellcord_1 { get; set; }
        public DbSet<CBstellcord_5> CBstellcord_5 { get; set; }
        public DbSet<CBstellcord_7> CBstellcord_7 { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<BTLH3>().ToTable("BTLH#3");
            modelBuilder.Entity<BTLH4>().ToTable("BTLH#4");
            modelBuilder.Entity<BTLH5>().ToTable("BTLH#5");

            modelBuilder.Entity<VIS_MHE1_1>().ToTable("VIS#MHE1_1").HasKey(x => x.id);
            modelBuilder.Entity<VIS_MHE1_2>().ToTable("VIS#MHE1_2").HasKey(x => x.id);
            modelBuilder.Entity<VIS_MHE2_1>().ToTable("VIS#MHE2_1").HasKey(x => x.id);
            modelBuilder.Entity<VIS_MHE2_2>().ToTable("VIS#MHE2_2").HasKey(x => x.id);
            modelBuilder.Entity<VIS_MHE3_1>().ToTable("VIS#MHE3_1").HasKey(x => x.id);
            modelBuilder.Entity<VIS_MHE3_2>().ToTable("VIS#MHE3_2").HasKey(x => x.id);
        }
    }
}