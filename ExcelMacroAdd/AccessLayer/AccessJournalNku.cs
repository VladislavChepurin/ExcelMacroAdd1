using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Data.Entity;
using System.Threading.Tasks;

namespace ExcelMacroAdd.AccessLayer
{
    public class AccessJournalNku
    {
        private readonly AppContext context;

        public AccessJournalNku(AppContext context)
        {
            this.context = context;
        }

        public async Task<IJournalNku> GetEntityJournal(string sArticle)
        {
            return await context.JornalNkus.FirstOrDefaultAsync(p => p.Article == sArticle);
        }

        public async void WriteUpdateDb(JournalNku entity)
        {
            if (entity != null)
            {
                context.Entry(entity).State = EntityState.Modified;
                await context.SaveChangesAsync();
            }
        }

        public async void AddValueDb(JournalNku entity)
        {
            if (entity != null)
            {
                context.JornalNkus.Add(entity);
                await context.SaveChangesAsync();
            }
        }
    }
}
