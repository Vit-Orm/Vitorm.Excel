using System.Collections.Generic;
using System.IO;

using OfficeOpenXml;


namespace Vitorm.Excel
{
    public class DbConfig
    {
        static DbConfig()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public DbConfig(string connectionString)
        {
            this.connectionString = connectionString;
        }

        public DbConfig(string connectionString, string readOnlyConnectionString)
        {
            this.connectionString = connectionString;
        }

        public DbConfig(Dictionary<string, object> config)
        {
            object value;
            if (config.TryGetValue("connectionString", out value))
                this.connectionString = value as string;
        }

        public string connectionString { get; set; }


        public virtual string databaseName => Path.GetFileNameWithoutExtension(connectionString);

        public virtual DbConfig WithDatabase(string databaseName)
        {
            var _connectionString = Path.Combine(Path.GetDirectoryName(connectionString), databaseName + ".xlsx");

            return new DbConfig(_connectionString);
        }

        internal string dbHashCode => connectionString.GetHashCode().ToString();
    }
}
