using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessCircuitBreaker
    {
        private readonly AppContext context;
        public AccessCircuitBreaker(AppContext context)
        {
            this.context = context;
        }

        public async Task<ISwitch> GetEntitySwitch(string current, string quantity)
        {
            return await context.Switchs.AsNoTracking().FirstOrDefaultAsync(p => p.Current == current && p.Quantity == quantity);
        }

        public async Task<IModul> GetEntityModule(string current, string curve, string maxCurrent, string quantity)
        {
            return await context.Moduls.AsNoTracking().FirstOrDefaultAsync(p => p.Current == current && p.Kurve == curve && p.MaxCurrent == maxCurrent && p.Quantity == quantity);
        }

        public string[] GetCircuitCurrentItems()
        {
            return context.Moduls
                .Select(p => p.Current)
                .ToHashSet()
                .ToArray();
        }

        public string[] GetCircuitCurveItems()
        {
            return context.Moduls
                .Select(p => p.Kurve)
                .ToHashSet()
                .ToArray();
        }
        public string[] GetCircuitMaxCurrentItems()
        {
            return context.Moduls
                .Select(p => p.MaxCurrent)
                .ToHashSet()
                .ToArray();
        }

        public string[] GetCircuitPolesItems()
        {
            return context.Moduls
                .Select(p => p.Quantity)
                .ToHashSet()
                .ToArray();
        }

        public string[] GetCircuitSwitchsItems()
        {
            return context.Switchs
                .Select(p => p.Current)
                .ToHashSet()
                .ToArray();
        }
        public string[] GetSwitchsPolesItems()
        {
            return context.Switchs
                .Select(p => p.Quantity)
                .ToHashSet()
                .ToArray();
        }
    }
}
