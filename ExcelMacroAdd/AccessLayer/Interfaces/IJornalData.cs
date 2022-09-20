using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Threading.Tasks;

namespace ExcelMacroAdd.AccessLayer.Interfaces
{
    public interface IJornalData
    {
        Task<IJornalNKU> GetEntityJornal(string sArticle);

        void WriteUpdateDB(JornalNKU entity);

        void AddValueDB(JornalNKU entity);
    }
}
