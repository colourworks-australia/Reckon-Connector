using System.Data.Entity;
using ReckonDesktop.Model;
using SQLite.CodeFirst;

namespace ReckonDesktop.Repository
{
    public class ReckonDesktopContext : DbContext
    {
        //public static int RequiredDatabaseVersion = 1;
        //public DbSet<Destination> Destinations { get; set; }
        //public DbSet<FolderMapping> FolderMappings { get; set; }
        public DbSet<Settings> Settingses { get; set; }
        public DbSet<SchemaVersion> SchemaVersions { get; set; }
        public DbSet<Security> Securitys { get; set; }

        public ReckonDesktopContext()
            : base("ReckonDb")
        {
        }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            //modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
            var initializer = new ReckonInitializer(modelBuilder);
            Database.SetInitializer(initializer);
        }

        public class ReckonInitializer : SqliteCreateDatabaseIfNotExists<ReckonDesktopContext>
        {
            public ReckonInitializer(DbModelBuilder modelBuilder)
                : base(modelBuilder)
            {
            }

            protected override void Seed(ReckonDesktopContext context)
            {
                context.Set<Settings>().Add(new Settings
                {
                    Id = 1,
                    Timer = 300,
                    //LogFilePath = @"C:\Premier19\MYOBPLOG.TXT"
                });
            }
        }
    }
}
