using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;

namespace ExcelMacroAdd.AccessLayer.Interfaces
{
    public interface IJornalData
    {
        IJornalNKU GetEntityJornal(string sArticle);

        void WriteUpdateDB(JornalNKU entity);

        void AddValueDB(JornalNKU entity);

    }
}
