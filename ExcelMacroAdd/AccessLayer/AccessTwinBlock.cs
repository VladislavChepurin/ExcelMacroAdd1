using System.Linq;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.AccessLayer
{
    public class AccessTwinBlock
    {
        private readonly AppContext context;
        public AccessTwinBlock(AppContext context)
        {
            this.context = context;
        }
        public string[] GetComboBox1Items()
        {
            return context.TwinBlockSwitchs
                .AsNoTracking()
                .Select(p => p.Current)
                .ToHashSet()
                .ToArray();
        }

        public (string, string, string, string, string) GetDataInTableDb(string current, bool isReverse)
        {
            var data = context.TwinBlockSwitchs
                .AsNoTracking()
                .Where(s => s.Current == current && s.IsReverse == isReverse)
                .Select(s => new {
                    newArticle = s.Article,
                    newDirectMountingHandle = s.DirectMountingHandle.Article,
                    newDoorHandle = s.DoorHandle.Article,
                    newStock = s.Stock.Article,
                    newAdditionalPole = s.AdditionalPole.Article
                }).FirstOrDefault();
            return (data?.newArticle, data?.newDirectMountingHandle, data?.newDoorHandle, data?.newStock, data?.newAdditionalPole);
        }

        public byte[] GetBlobPictureDb(string current, bool isReverse)
        {
            return context.TwinBlockSwitchs
                .AsNoTracking()
                .Where(s => s.Current == current && s.IsReverse == isReverse)
                .Select(s => s.Picture
                ).FirstOrDefault();
        }
    }
}