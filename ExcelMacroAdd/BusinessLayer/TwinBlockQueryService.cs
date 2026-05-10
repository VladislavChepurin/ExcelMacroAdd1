using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using System;
using System.Data.Entity;
using System.Linq;

namespace ExcelMacroAdd.BusinessLayer
{
    public sealed class TwinBlockQueryService : ITwinBlockService
    {
        private readonly IUnitOfWorkFactory unitOfWorkFactory;

        public TwinBlockQueryService(IUnitOfWorkFactory unitOfWorkFactory)
        {
            this.unitOfWorkFactory = unitOfWorkFactory ?? throw new ArgumentNullException(nameof(unitOfWorkFactory));
        }

        public string[] GetComboBox1Items()
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return unitOfWork.Context.TwinBlockSwitchs
                    .AsNoTracking()
                    .Select(p => p.Current)
                    .ToHashSet()
                    .ToArray();
            }
        }

        public (string, string, string, string, string) GetDataInTableDb(string current, bool isReverse)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                var data = unitOfWork.Context.TwinBlockSwitchs
                    .AsNoTracking()
                    .Where(s => s.Current == current && s.IsReverse == isReverse)
                    .Select(s => new
                    {
                        Article = s.Article,
                        DirectMountingHandle = s.DirectMountingHandle.Article,
                        DoorHandle = s.DoorHandle.Article,
                        Stock = s.Stock.Article,
                        AdditionalPole = s.AdditionalPole.Article
                    })
                    .FirstOrDefault();

                return (data?.Article, data?.DirectMountingHandle, data?.DoorHandle, data?.Stock, data?.AdditionalPole);
            }
        }

        public byte[] GetBlobPictureDb(string current, bool isReverse)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return unitOfWork.Context.TwinBlockSwitchs
                    .AsNoTracking()
                    .Where(s => s.Current == current && s.IsReverse == isReverse)
                    .Select(s => s.Picture)
                    .FirstOrDefault();
            }
        }
    }
}
