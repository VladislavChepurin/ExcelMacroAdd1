using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Data.Entity;
using System.Threading.Tasks;

namespace ExcelMacroAdd.AccessLayer
{
    public class AccessData : IForm2Data, IJornalData
    {
        public async Task<ISwitch> GetEntitySwitch(string current, string quantity)
        {
            using (DataContext db = new DataContext())
            {
                return await db.Switch.AsNoTracking().FirstOrDefaultAsync(p => p.Current == current && p.Quantity == quantity);
            }
        }

        public async Task<IModul> GetEntityModul(string current, string kurve, string maxCurrent, string quantity)
        {
            using (DataContext db = new DataContext())
            {
                return await db.Modul.AsNoTracking().FirstOrDefaultAsync(p => p.Current == current && p.Kurve == kurve && p.MaxCurrent == maxCurrent && p.Quantity == quantity);
            }
        }

        public async Task<IJornalNKU> GetEntityJornal(string sArticle)
        {
            using (DataContext db = new DataContext())
            {
                return await db.JornalNKU.FirstOrDefaultAsync(p => p.Article == sArticle);
            }
        }

        public async void WriteUpdateDB(JornalNKU entity)
        {
            using (DataContext db = new DataContext())
            {
                if (entity != null)
                {
                    db.Entry(entity).State = EntityState.Modified;
                    await db.SaveChangesAsync();
                }
            }
        }

        public async void AddValueDB(JornalNKU entity)
        {
            using (DataContext db = new DataContext())
            {
                if (entity != null)
                {    
                    db.JornalNKU.Add(entity);
                    await db.SaveChangesAsync();
                }
            }
        }
    }
}
