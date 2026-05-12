using System;
using System.Data.Entity;
using System.Data.SQLite;
using System.IO;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class AppContext : DbContext
    {
        public AppContext(string databaseFilePath) : base(CreateConnection(databaseFilePath), true)
        {
        }

        public DbSet<BoxBase> JornalNkus { get; set; }
        public DbSet<Switch> Switches { get; set; }
        public DbSet<CircuitBreaker> CircuitBreakers { get; set; }
        public DbSet<TransformerAttribute> TransformerAttributes { get; set; }
        public DbSet<Transformer> Transformers { get; set; }
        public DbSet<DirectMountingHandle> DirectMountingHandles { get; set; }
        public DbSet<DoorHandle> DoorHandles { get; set; }
        public DbSet<Stock> Stocks { get; set; }
        public DbSet<AdditionalPole> AdditionalPoles { get; set; }
        public DbSet<TwinBlockSwitch> TwinBlockSwitchs { get; set; }
        public DbSet<MaterialBox> Materials { get; set; }
        public DbSet<ExecutionBox> Executions { get; set; }
        public DbSet<ProductVendor> ProductVendors { get; set; }
        public DbSet<ProductGroup> ProductGroups { get; set; }
        public DbSet<ProductSeries> ProductSeriess { get; set; }
        public DbSet<ShuntTrip24v> ShuntTrip24vs { get; set; }
        public DbSet<ShuntTrip48v> ShuntTrip48vs { get; set; }
        public DbSet<ShuntTrip230v> ShuntTrip230vs { get; set; }
        public DbSet<UndervoltageRelease> UndervoltageReleases { get; set; }
        public DbSet<SignalContact> SignalContacts { get; set; }
        public DbSet<AuxiliaryContact> AuxiliaryContacts { get; set; }
        public DbSet<SignalOrAuxiliaryContact> SignalOrAuxiliaryContacts { get; set; }
        public DbSet<NotPriceComponent> NotPriceComponents { get; set; }
        public DbSet<Multiplicity> Multiplicities { get; set; }

        private static SQLiteConnection CreateConnection(string databaseFilePath)
        {
            if (string.IsNullOrWhiteSpace(databaseFilePath))
            {
                throw new ArgumentException("Путь к файлу базы данных не может быть пустым.", nameof(databaseFilePath));
            }

            string fullDatabaseFilePath = Path.GetFullPath(databaseFilePath);
            var connectionStringBuilder = new SQLiteConnectionStringBuilder
            {
                DataSource = fullDatabaseFilePath
            };

            return new SQLiteConnection(connectionStringBuilder.ConnectionString);
        }
    }
}
