using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.DataLayer.Interfaces;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using ExcelMacroAdd.UserVariables;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.AccessLayer
{
    public class AccessData : IForm2Data, IJournalData, IForm4Data
    {
        private readonly AppContext context;
        public AccessData(AppContext context)
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

        public async Task<IJournalNku> GetEntityJournal(string sArticle)
        {
            return await context.JornalNkus.FirstOrDefaultAsync(p => p.Article == sArticle);
        }

        public async void WriteUpdateDB(JournalNku entity)
        {
            if (entity != null)
            {
                context.Entry(entity).State = EntityState.Modified;
                await context.SaveChangesAsync();
            }
        }

        public async void AddValueDB(JournalNku entity)
        {
            if (entity != null)
            {
                context.JornalNkus.Add(entity);
                await context.SaveChangesAsync();
            }
            
        }

        public string[] GetComboBox2Items(string current)
        {
            return context.Transformers
                .Where(p => p.Current == current)
                .Select(p => p.Bus)
                .ToHashSet()
                .ToArray();
        }

        public string[] GetComboBox3Items(string current, string bus)
        {
            return context.Transformers
                .Where(p => p.Current == current && p.Bus == bus)
                .Select(p => p.Accuracy)
                .ToHashSet()
                .ToArray();
        }

        public string[] GetComboBox4Items(string current, string bus, string accuracy)
        {
            return context.Transformers
                    .Where(p => p.Current == current && p.Bus == bus && p.Accuracy == accuracy)
                    .Select(p => p.Power)
                    .ToHashSet()
                    .ToArray();
        }

        public StructTransformer GetTransformerArticle(string current, string bus, string accuracy, string power)
        {
            var trans = context.Transformers
                    .Where(t => t.Current == current
                                && t.Bus == bus
                                && t.Accuracy == accuracy
                                && t.Power == power)
                    .Select(t => new { IekTti = t.Iek, EkfTte = t.Ekf, KeazTtk = t.Keaz, TdmTtn = t.Tdm, IekTop = t.IekTopTpsh, DekTop = t.DekraftTopTpsh })
                    .FirstOrDefault();

                return new StructTransformer() { IekTti = trans?.IekTti, EkfTte = trans?.EkfTte, KeazTtk = trans?.KeazTtk, TdmTtn = trans?.TdmTtn, IekTop = trans?.IekTop, DekTop = trans?.DekTop };
        }
    }
}
