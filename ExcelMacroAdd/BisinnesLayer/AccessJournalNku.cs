using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using Microsoft.Extensions.Caching.Memory;
using System.Data.Entity;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessJournalNku
    {
        private readonly AppContext context;
        private readonly IMemoryCache cache;
        public AccessJournalNku(AppContext context, IMemoryCache cache)
        {
            this.context = context;
            this.cache = cache;
        }

        public async Task<IBoxBase> GetEntityJournal(string sArticle)
        {
            cache.TryGetValue(sArticle, out IBoxBase boxBase);
            if (boxBase == null)
            {
                boxBase = await context.JornalNkus.Include(p => p.MaterialBox).Include(p => p.ExecutionBox).FirstOrDefaultAsync(p => p.Article == sArticle) as IBoxBase;
                cache.Set(sArticle, boxBase, new MemoryCacheEntryOptions().SetAbsoluteExpiration(System.TimeSpan.FromMinutes(5)));
            }
            return boxBase;
        }

        public async void WriteUpdateDb(BoxBase entity)
        {
            if (entity != null)
            {
                context.Entry(entity).State = EntityState.Modified;
                await context.SaveChangesAsync();
            }
        }

        public async Task<IMaterialBox> GetMaterialEntityByName(string material)
        {
            return await context.Materials.FirstOrDefaultAsync(p => p.MaterialValue == material);
        }

        public async Task<IExecutionBox> GetExecutionEntityByName(string execution)
        {
            return await context.Executions.FirstOrDefaultAsync(p => p.ExecutionValue == execution);
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
