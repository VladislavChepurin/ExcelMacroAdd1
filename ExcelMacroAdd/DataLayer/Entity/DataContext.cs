using System;
using System.Data.Entity;

namespace ExcelMacroAdd.DataLayer.Entity
{
    public class DataContext: DbContext
    {
        public DataContext() : base("Context")
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            AppDomain.CurrentDomain.SetData("DataDirectory", path);
        }

        public DbSet<JornalNKU> JornalNKU { get; set; }
        public DbSet<Switch> Switch { get; set; }
        public DbSet<Modul> Modul { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {          
            //Настройка таблицы JornalNKUs
            modelBuilder.Entity<JornalNKU>().Property(p => p.Id).HasColumnName("id");
            modelBuilder.Entity<JornalNKU>().Property(p => p.Ip).HasColumnName("ip");
            modelBuilder.Entity<JornalNKU>().Property(p => p.Klima).HasColumnName("klima");
            modelBuilder.Entity<JornalNKU>().Property(p => p.Reserve).HasColumnName("reserve");
            modelBuilder.Entity<JornalNKU>().Property(p => p.Height).HasColumnName("height");
            modelBuilder.Entity<JornalNKU>().Property(p => p.Width).HasColumnName("width");
            modelBuilder.Entity<JornalNKU>().Property(p => p.Depth).HasColumnName("depth");
            modelBuilder.Entity<JornalNKU>().Property(p => p.Article).HasColumnName("article");
            modelBuilder.Entity<JornalNKU>().Property(p => p.Execution).HasColumnName("execution");
            modelBuilder.Entity<JornalNKU>().Property(p => p.Vendor).HasColumnName("vendor");
            
            //Настройка таблицы Switchs
            modelBuilder.Entity<Switch>().Property(p => p.Id).HasColumnName("id");
            modelBuilder.Entity<Switch>().Property(p => p.Current).HasColumnName("current");
            modelBuilder.Entity<Switch>().Property(p => p.Quantity).HasColumnName("quantity");
            modelBuilder.Entity<Switch>().Property(p => p.Iek).HasColumnName("iek");
            modelBuilder.Entity<Switch>().Property(p => p.EkfProxima).HasColumnName("ekf_proxima");
            modelBuilder.Entity<Switch>().Property(p => p.EkfAvers).HasColumnName("ekf_avers");
            modelBuilder.Entity<Switch>().Property(p => p.Keaz).HasColumnName("keaz");
            modelBuilder.Entity<Switch>().Property(p => p.Abb).HasColumnName("abb");
            modelBuilder.Entity<Switch>().Property(p => p.Dekraft).HasColumnName("dekraft");
            modelBuilder.Entity<Switch>().Property(p => p.Schneider).HasColumnName("schneider");
            modelBuilder.Entity<Switch>().Property(p => p.Tdm).HasColumnName("tdm");
            //Настройка таблицы Modul
            modelBuilder.Entity<Modul>().Property(p => p.Id).HasColumnName("id");
            modelBuilder.Entity<Modul>().Property(p => p.MaxCurrent).HasColumnName("max_current");
            modelBuilder.Entity<Modul>().Property(p => p.Current).HasColumnName("current");
            modelBuilder.Entity<Modul>().Property(p => p.Kurve).HasColumnName("kurve");
            modelBuilder.Entity<Modul>().Property(p => p.Quantity).HasColumnName("quantity");
            modelBuilder.Entity<Modul>().Property(p => p.IekVa47).HasColumnName("iek_va47");
            modelBuilder.Entity<Modul>().Property(p => p.IekVa47m).HasColumnName("iek_va47m");
            modelBuilder.Entity<Modul>().Property(p => p.EkfProxima).HasColumnName("ekf_proxima");
            modelBuilder.Entity<Modul>().Property(p => p.EkfAvers).HasColumnName("ekf_avers");
            modelBuilder.Entity<Modul>().Property(p => p.Keaz).HasColumnName("keaz");
            modelBuilder.Entity<Modul>().Property(p => p.Abb).HasColumnName("abb");
            modelBuilder.Entity<Modul>().Property(p => p.Dkc).HasColumnName("dkc");
            modelBuilder.Entity<Modul>().Property(p => p.Dekraft).HasColumnName("dekraft");
            modelBuilder.Entity<Modul>().Property(p => p.Schneider).HasColumnName("schneider");
            modelBuilder.Entity<Modul>().Property(p => p.Tdm).HasColumnName("tdm");
            modelBuilder.Entity<Modul>().Property(p => p.IekArmat).HasColumnName("iek_armat");
            base.OnModelCreating(modelBuilder);
        }
    }
}
