using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Data.Entity;
using System.Threading.Tasks;

namespace ExcelMacroAdd.AccessLayer
{
    public class AccessData : IForm2Data, IJournalData
    {
        public async Task<ISwitch> GetEntitySwitch(string current, string quantity)
        {
            using (var db = new AppContext())
            {
                return await db.Switchs.AsNoTracking().FirstOrDefaultAsync(p => p.Current == current && p.Quantity == quantity);
            }
        }

        public async Task<IModul> GetEntityModule(string current, string curve, string maxCurrent, string quantity)
        {
            using (var db = new AppContext())
            {
                return await db.Moduls.AsNoTracking().FirstOrDefaultAsync(p => p.Current == current && p.Kurve == curve && p.MaxCurrent == maxCurrent && p.Quantity == quantity);
            }
        }

        public async Task<IJournalNku> GetEntityJournal(string sArticle)
        {
            using (AppContext db = new AppContext())
            {
                return await db.JornalNkus.FirstOrDefaultAsync(p => p.Article == sArticle);
            }
        }

        public async void WriteUpdateDB(JournalNku entity)
        {
            using (AppContext db = new AppContext())
            {
                if (entity != null)
                {
                    db.Entry(entity).State = EntityState.Modified;
                    await db.SaveChangesAsync();
                }
            }
        }

        public async void AddValueDB(JournalNku entity)
        {
            using (AppContext db = new AppContext())
            {
                if (entity != null)
                {    
                    db.JornalNkus.Add(entity);
                    await db.SaveChangesAsync();
                }
            }
        }
    }
}
