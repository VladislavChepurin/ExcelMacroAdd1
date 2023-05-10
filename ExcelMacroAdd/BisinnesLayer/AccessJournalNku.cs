using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Data.Entity;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessJournalNku
    {
        private readonly AppContext context;

        public AccessJournalNku(AppContext context)
        {
            this.context = context;
        }

        public async Task<IBoxBase> GetEntityJournal(string sArticle)
        {
            return await context.JornalNkus.FirstOrDefaultAsync(p => p.Article == sArticle);
        }

        public async void WriteUpdateDb(BoxBase entity)
        {
            if (entity != null)
            {
                context.Entry(entity).State = EntityState.Modified;
                await context.SaveChangesAsync();
            }
        }

        public async Task<IExecution> GetExecutionEntityByName(string execution)
        {             
            return await context.Executions.FirstOrDefaultAsync(p => p.ExecutionValue == execution);
        }

        public async Task<IExecution> GetExecutionEntityById(int id)
        {
            return await context.Executions.FirstOrDefaultAsync(p => p.Id == id);
        }

        public async void AddValueDb(BoxBase entity)
        {
            if (entity != null)
            {
                context.JornalNkus.Add(entity);
                await context.SaveChangesAsync();
            }
        }
    }
}
