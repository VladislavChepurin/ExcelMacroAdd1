using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.UserVariables;
using System.Linq;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessTransformer
    {
        private readonly AppContext context;

        public AccessTransformer(AppContext context)
        {
            this.context = context;
        }

        public string[] GetComboBox2Items(string current)
        {
            return context.Transformers
                .Where(p => p.Current == current)
                .Select(p => p.TransformerAttribute.Bus)
                .ToHashSet()
                .ToArray();
        }

        public string[] GetComboBox3Items(string current, string bus)
        {
            return context.Transformers
                .Where(p => p.Current == current && p.TransformerAttribute.Bus == bus)
                .Select(p => p.Accuracy)
                .ToHashSet()
                .ToArray();
        }

        public string[] GetComboBox4Items(string current, string bus, string accuracy)
        {
            return context.Transformers
                .Where(p => p.Current == current && p.TransformerAttribute.Bus == bus && p.Accuracy == accuracy)
                .Select(p => p.Power)
                .ToHashSet()
                .ToArray();
        }

        public StructTransformer GetArticle(string current, string bus, string accuracy, string power)
        {
            var trans = context.Transformers
                .Where(t => t.Current == current
                            && t.TransformerAttribute.Bus == bus
                            && t.Accuracy == accuracy
                            && t.Power == power)
                .Select(t => new { IekTti = t.Iek, EkfTte = t.Ekf, KeazTtk = t.Keaz, TdmTtn = t.Tdm, IekTop = t.IekTopTpsh, DekTop = t.DekraftTopTpsh })
                .FirstOrDefault();
            return new StructTransformer() { IekTti = trans?.IekTti, EkfTte = trans?.EkfTte, KeazTtk = trans?.KeazTtk, TdmTtn = trans?.TdmTtn, IekTop = trans?.IekTop, DekTop = trans?.DekTop };
        }

        public byte[] GetBlobPictureDb(string current, string bus, string accuracy, string power)
        {
            return context.Transformers
                .Where(t => t.Current == current
                            && t.TransformerAttribute.Bus == bus
                            && t.Accuracy == accuracy
                            && t.Power == power)
                .Select(s => s.TransformerAttribute.Picture
                ).FirstOrDefault();
        }
    }
}
