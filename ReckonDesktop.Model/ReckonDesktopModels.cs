using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ReckonDesktop.Model
{
    public class Settings
    {
        [Key]
        public int Id { get; set; }
        public int Timer { get; set; }
        public string FileFilter { get; set; }
        public string FolderPath { get; set; }
        public string ConnectionString { get; set; }
        public string MyobUser { get; set; }
        public string MyobPassword { get; set; }
        public string AutofileEndpoint { get; set; }
        public string AutofileUser { get; set; }
        public string AutofilePassword { get; set; }
        public string LogFilePath { get; set; }
        public bool? UseHosted { get; set; }
        public string id_token { get; set; }
        public string access_token { get; set; }
        public long? tokenExpires { get; set; }
    }

    public class SchemaVersion
    {
        [Key]
        public int Id { get; set; }
        public int Version { get; set; }
    }

    [Table("Security")]
    public class Security
    {
        [Key]
        public int Id { get; set; }
        public string Login { get; set; }
        public string Password { get; set; }
        public bool Protected { get; set; }
    }
}
