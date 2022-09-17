using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Data.Entity;
using System.Linq;

namespace ExcelMacroAdd.AccessLayer
{
    public class AccessData : IForm2Data, IJornalData
    {
        public ISwitch GetEntitySwitch(string current, string quantity)
        {
            using (DataContext db = new DataContext())
            {
                var switchs = db.Switch;
                return switchs.FirstOrDefault(p => p.Current == current && p.Quantity == quantity);
            }
        }

        public IModul GetEntityModul(string current, string kurve, string maxCurrent, string quantity)
        {
            using (DataContext db = new DataContext())
            {
                var moduls = db.Modul;
                return moduls.FirstOrDefault(p => p.Current == current && p.Kurve == kurve && p.MaxCurrent == maxCurrent && p.Quantity == quantity);
            }
        }

        public IJornalNKU GetEntityJornal(string sArticle)
        {
            using (DataContext db = new DataContext())
            {
                var jornalNKUs = db.JornalNKU;
                return jornalNKUs.FirstOrDefault(p => p.Article == sArticle);
            }
        }

        public void WriteUpdateDB(JornalNKU entity)
        {
            using (DataContext db = new DataContext())
            {
                if (entity != null)
                {
                    db.Entry(entity).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }
        }

        public void AddValueDB(JornalNKU entity)
        {
            using (DataContext db = new DataContext())
            {
                if (entity != null)
                {    
                    db.JornalNKU.Add(entity);                  
                    db.SaveChanges();
                }
            }
        }
    }
}
