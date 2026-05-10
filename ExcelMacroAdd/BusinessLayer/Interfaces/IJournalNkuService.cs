using ExcelMacroAdd.DataLayer.Interfaces;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BusinessLayer.Interfaces
{
    public interface IJournalNkuService
    {
        Task<IBoxBase> GetEntityJournal(string sArticle);

        Task<IReadOnlyDictionary<string, IBoxBase>> GetEntityJournalBatch(IEnumerable<string> articles);
    }
}
