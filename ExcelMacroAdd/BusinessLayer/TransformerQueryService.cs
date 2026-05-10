using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using ExcelMacroAdd.UserVariables;
using System;
using System.Data.Entity;
using System.Linq;

namespace ExcelMacroAdd.BusinessLayer
{
    public sealed class TransformerQueryService : ITransformerService
    {
        private readonly IUnitOfWorkFactory unitOfWorkFactory;

        public TransformerQueryService(IUnitOfWorkFactory unitOfWorkFactory)
        {
            this.unitOfWorkFactory = unitOfWorkFactory ?? throw new ArgumentNullException(nameof(unitOfWorkFactory));
        }

        public string[] GetComboBox2Items(string current)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return unitOfWork.Context.Transformers
                    .AsNoTracking()
                    .Where(p => p.Current == current)
                    .Select(p => p.TransformerAttribute.Bus)
                    .ToHashSet()
                    .ToArray();
            }
        }

        public string[] GetComboBox3Items(string current, string bus)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return unitOfWork.Context.Transformers
                    .AsNoTracking()
                    .Where(p => p.Current == current && p.TransformerAttribute.Bus == bus)
                    .Select(p => p.Accuracy)
                    .ToHashSet()
                    .ToArray();
            }
        }

        public string[] GetComboBox4Items(string current, string bus, string accuracy)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return unitOfWork.Context.Transformers
                    .AsNoTracking()
                    .Where(p => p.Current == current && p.TransformerAttribute.Bus == bus && p.Accuracy == accuracy)
                    .Select(p => p.Power)
                    .ToHashSet()
                    .ToArray();
            }
        }

        public StructTransformer GetArticle(string current, string bus, string accuracy, string power)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                var transformer = unitOfWork.Context.Transformers
                    .AsNoTracking()
                    .Where(t => t.Current == current
                                && t.TransformerAttribute.Bus == bus
                                && t.Accuracy == accuracy
                                && t.Power == power)
                    .Select(t => new
                    {
                        IekTti = t.Iek,
                        EkfTte = t.Ekf,
                        KeazTtk = t.Keaz,
                        TdmTtn = t.Tdm,
                        IekTop = t.IekTopTpsh,
                        DekTop = t.DekraftTopTpsh
                    })
                    .FirstOrDefault();

                return new StructTransformer
                {
                    IekTti = transformer?.IekTti,
                    EkfTte = transformer?.EkfTte,
                    KeazTtk = transformer?.KeazTtk,
                    TdmTtn = transformer?.TdmTtn,
                    IekTop = transformer?.IekTop,
                    DekTop = transformer?.DekTop
                };
            }
        }

        public byte[] GetBlobPictureDb(string current, string bus, string accuracy, string power)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                return unitOfWork.Context.Transformers
                    .AsNoTracking()
                    .Where(t => t.Current == current
                                && t.TransformerAttribute.Bus == bus
                                && t.Accuracy == accuracy
                                && t.Power == power)
                    .Select(s => s.TransformerAttribute.Picture)
                    .FirstOrDefault();
            }
        }
    }
}
