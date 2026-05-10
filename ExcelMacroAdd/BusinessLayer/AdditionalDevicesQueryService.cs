using ExcelMacroAdd.BusinessLayer.Interfaces;
using ExcelMacroAdd.DataLayer.UnitOfWork;
using ExcelMacroAdd.Models;
using System;
using System.Data.Entity;
using System.Linq;

namespace ExcelMacroAdd.BusinessLayer
{
    public sealed class AdditionalDevicesQueryService : IAdditionalDevicesService
    {
        private readonly IUnitOfWorkFactory unitOfWorkFactory;

        public AdditionalDevicesQueryService(IUnitOfWorkFactory unitOfWorkFactory)
        {
            this.unitOfWorkFactory = unitOfWorkFactory ?? throw new ArgumentNullException(nameof(unitOfWorkFactory));
        }

        public AdditionalDevices GetEntityAdditionalCircuitBreaker(string articleNumber)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                var data = unitOfWork.Context.CircuitBreakers
                    .AsNoTracking()
                    .Where(s => s.ArticleNumber == articleNumber)
                    .Select(s => new
                    {
                        Vendor = s.ProductVendor.VendorName,
                        ShuntTrip24v = s.ShuntTrip24v.Article,
                        ShuntTrip48v = s.ShuntTrip48v.Article,
                        ShuntTrip230v = s.ShuntTrip230v.Article,
                        UndervoltageRelease = s.UndervoltageRelease.Article,
                        SignalContact = s.SignalContact.Article,
                        AuxiliaryContact = s.AuxiliaryContact.Article,
                        SignalOrAuxiliaryContact = s.SignalOrAuxiliaryContact.Article
                    })
                    .FirstOrDefault();

                return new AdditionalDevices(
                    data?.Vendor,
                    data?.ShuntTrip24v,
                    data?.ShuntTrip48v,
                    data?.ShuntTrip230v,
                    data?.UndervoltageRelease,
                    data?.SignalContact,
                    data?.AuxiliaryContact,
                    data?.SignalOrAuxiliaryContact);
            }
        }

        public AdditionalDevices GetEntityAdditionalSwitch(string articleNumber)
        {
            using (var unitOfWork = unitOfWorkFactory.Create())
            {
                var data = unitOfWork.Context.Switches
                    .AsNoTracking()
                    .Where(s => s.ArticleNumber == articleNumber)
                    .Select(s => new
                    {
                        Vendor = s.ProductVendor.VendorName,
                        ShuntTrip24v = s.ShuntTrip24v.Article,
                        ShuntTrip48v = s.ShuntTrip48v.Article,
                        ShuntTrip230v = s.ShuntTrip230v.Article,
                        UndervoltageRelease = s.UndervoltageRelease.Article,
                        SignalContact = s.SignalContact.Article,
                        AuxiliaryContact = s.AuxiliaryContact.Article,
                        SignalOrAuxiliaryContact = s.SignalOrAuxiliaryContact.Article
                    })
                    .FirstOrDefault();

                return new AdditionalDevices(
                    data?.Vendor,
                    data?.ShuntTrip24v,
                    data?.ShuntTrip48v,
                    data?.ShuntTrip230v,
                    data?.UndervoltageRelease,
                    data?.SignalContact,
                    data?.AuxiliaryContact,
                    data?.SignalOrAuxiliaryContact);
            }
        }
    }
}
