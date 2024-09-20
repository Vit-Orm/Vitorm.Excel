using System.Data;

using OfficeOpenXml;

namespace Vitorm.Excel
{
    public partial class DbContext : Vitorm.DbContext
    {
        public DbConfig dbConfig { get; protected set; }

        public DbContext(DbConfig dbConfig) : base(DbSetConstructor.CreateDbSet)
        {
            this.dbConfig = dbConfig;
        }

        public DbContext(string connectionString) : this(new DbConfig(connectionString))
        {
        }


        #region Transaction
        public virtual IDbTransaction BeginTransaction() => throw new System.NotImplementedException();
        public virtual IDbTransaction GetCurrentTransaction() => throw new System.NotImplementedException();

        #endregion



        public virtual string databaseName => dbConfig.databaseName;
        public virtual void ChangeDatabase(string databaseName)
        {
            dbConfig = dbConfig.WithDatabase(databaseName);
        }



        #region dbConnection

        protected ExcelPackage _dbConnection;
        public virtual ExcelPackage dbConnection => _dbConnection ??= new ExcelPackage(dbConfig.connectionString);
        public virtual ExcelPackage readOnlyDbConnection => dbConnection;

        #endregion


        public override void Dispose()
        {
            if (_dbConnection != null)
            {
                try
                {
                    //_dbConnection.Save();
                    _dbConnection.Dispose();
                    _dbConnection = null;
                }
                catch (System.Exception ex)
                {
                }
            }

            base.Dispose();
        }


    }
}