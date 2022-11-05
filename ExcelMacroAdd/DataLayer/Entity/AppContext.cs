using System;
using System.Data.Entity;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class AppContext: DbContext
    {
        public AppContext(string path) : base("Context")
        {
            AppDomain.CurrentDomain.SetData("DataDirectory", path);
        }        
        public DbSet<JournalNku> JornalNkus { get; set; }
        public DbSet<Switch> Switchs { get; set; }
        public DbSet<Modul> Moduls { get; set; }
        public DbSet<TransformerAttribute> TransformerAttributes { get; set; }
        public DbSet<Transformer> Transformers { get; set; }
        public DbSet<DirectMountingHandle> DirectMountingHandles { get; set; }
        public DbSet<DoorHandle> DoorHandles { get; set; }
        public DbSet<Stock> Stocks { get; set; }
        public DbSet<AdditionalPole> AdditionalPoles { get; set; }
        public DbSet<TwinBlockSwitch> TwinBlockSwitchs { get; set; }
    }
}
