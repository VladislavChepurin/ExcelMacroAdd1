using ExcelMacroAdd.DataLayer.Entity;

namespace ExcelMacroAdd.DataLayer.Infrastructure
{
    public interface IAppContextFactory
    {
        AppContext Create();
    }
}
