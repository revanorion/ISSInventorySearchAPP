using System.Data.Entity;
using ISSIAS_Library.Entities;

namespace ISSIAS_Library
{
    public class FATSContext : DbContext
    {
        public FATSContext() : base("name=DbConnection")
        {
            //Database.SetInitializer<InventoryLabelingContext>(null);
            Database.SetInitializer<FATSContext>(null);
        }

        public virtual DbSet<FatsAsset> FatsAsset { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.HasDefaultSchema("HRASSET");
        }
        public override int SaveChanges()
        {
            // Throw if they try to call this
            throw new System.InvalidOperationException("This context is read-only.");
        }
    }
}
