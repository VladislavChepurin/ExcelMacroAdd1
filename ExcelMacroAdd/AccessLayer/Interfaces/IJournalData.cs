using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Threading.Tasks;

namespace ExcelMacroAdd.AccessLayer.Interfaces
{
    public interface IJournalData
    {
        Task<IJournalNku> GetEntityJournal(string sArticle);

        void WriteUpdateDB(JournalNku entity);

        void AddValueDB(JournalNku entity);
    }
}
