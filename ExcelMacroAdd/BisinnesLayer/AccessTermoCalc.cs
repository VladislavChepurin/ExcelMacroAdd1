using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Data.Entity;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessTermoCalc
    {
        private readonly AppContext context;

        public AccessTermoCalc(AppContext context)
        {
            this.context = context;
        }

        public async Task<IBoxBase> GetEntityJournal(string sArticle)
        {
            return await context.JornalNkus.FirstOrDefaultAsync(p => p.Article == sArticle);
        }
    }
}
