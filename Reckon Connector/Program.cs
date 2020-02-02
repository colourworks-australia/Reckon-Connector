using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ReckonDesktop.Model;
using ReckonDesktop.Repository;

namespace Reckon_Connector
{
     static class Program
    {
        static int RequiredDatabaseVersion = 4;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                AppDomain.CurrentDomain.SetData("DataDirectory", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData));
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                CreateAndSeedDatabase();
                Application.Run(new frmMain());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        public static void Initialize(ReckonDesktopContext context)
        {
            int currentVersion = 0;
            try
            {
                if (context.SchemaVersions.Any())
                    currentVersion = context.SchemaVersions.Max(x => x.Version);
            }
            catch (Exception ex)
            {
            }
            SPWatcherContextHelper mmSqliteHelper = new SPWatcherContextHelper();
            while (currentVersion < RequiredDatabaseVersion)
            {
                currentVersion++;
                foreach (string migration in mmSqliteHelper.Migrations[currentVersion])
                {
                    if (!string.IsNullOrEmpty(migration))
                    {
                        context.Database.ExecuteSqlCommand(migration);
                    }
                }
                context.SchemaVersions.Add(new SchemaVersion { Version = currentVersion });
                context.SaveChanges();
            }
        }

        private class SPWatcherContextHelper
        {
            public SPWatcherContextHelper()
            {
                Migrations = new Dictionary<int, IList>();

                MigrationVersion1();
                MigrationVersion2();
                MigrationVersion3();
                MigrationVersion4();
                //MigrationVersion5();
            }

            public Dictionary<int, IList> Migrations { get; private set; }

            private void MigrationVersion1()
            {
                var steps = new List<string>
                {
                    ""
                };
                Migrations.Add(1, steps);
            }

            private void MigrationVersion2()
            {
                var steps = new List<string>
                {
//                    "ALTER TABLE \"Settings\" ADD COLUMN \"UseHosted\" BIT"
                };
                Migrations.Add(2, steps);
            }
            private void MigrationVersion3()
            {
                var steps = new List<string>
                {
                    //"ALTER TABLE \"Settings\" ADD COLUMN \"id_token\" TEXT",
                    //"ALTER TABLE \"Settings\" ADD COLUMN \"access_token\" TEXT"
                };
                Migrations.Add(3, steps);
            }

            private void MigrationVersion4()
            {
                var steps = new List<string>
                {
                    //"ALTER TABLE \"Settings\" ADD COLUMN \"tokenExpires\" INTEGER"
                };
                Migrations.Add(4, steps);
            }


            //private void MigrationVersion5()
            //{
            //    var steps = new List<string>
            //    {
            //        "ALTER TABLE \"FolderMappings\" ADD COLUMN \"ConnectionString\" TEXT",
            //        "ALTER TABLE \"FolderMappings\" ADD COLUMN \"ShareName\" TEXT"
            //    };
            //    Migrations.Add(5, steps);
            //}

        }

        private static ReckonDesktopContext CreateAndSeedDatabase()
        {
            var context = new ReckonDesktopContext();
            Initialize(context);
            return context;
        }
    }
}
