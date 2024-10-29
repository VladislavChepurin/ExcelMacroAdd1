using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Models;
using System.Linq;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessAdditionalModularDevices
    {
        private readonly AppContext context;
        public AccessAdditionalModularDevices(AppContext context)
        {
            this.context = context;
        }

        public AdditionalDevices GetEntityAdditionalCircuitBreaker(string articleNumber)
        {
            var data = context.CircuitBreakers
              .AsNoTracking()
              .Where(s => s.ArticleNumber == articleNumber)
              .Select(s => new {
                  newVendor = s.ProductVendor.VendorName,
                  newShuntTrip24v = s.ShuntTrip24v.Article,
                  newShuntTrip48v = s.ShuntTrip48v.Article,
                  newShuntTrip230v = s.ShuntTrip230v.Article,
                  newUndervoltageRelease = s.UndervoltageRelease.Article,
                  newSignalContact = s.SignalContact.Article,
                  newAuxiliaryContact = s.AuxiliaryContact.Article,
                  newSignalOrAuxiliaryContact = s.SignalOrAuxiliaryContact.Article
              }).FirstOrDefault();

            return (new AdditionalDevices(data?.newVendor, data?.newShuntTrip24v, data?.newShuntTrip48v, data?.newShuntTrip230v,
                data?.newUndervoltageRelease, data?.newSignalContact, data?.newAuxiliaryContact, data?.newSignalOrAuxiliaryContact));
        }

        public AdditionalDevices GetEntityAdditionalSwitch(string articleNumber)
        {            
            var data = context.Switches
              .AsNoTracking()
              .Where(s => s.ArticleNumber == articleNumber)
              .Select(s => new {
                  newVendor = s.ProductVendor.VendorName,
                  newShuntTrip24v = s.ShuntTrip24v.Article,
                  newShuntTrip48v = s.ShuntTrip48v.Article,
                  newShuntTrip230v = s.ShuntTrip230v.Article,
                  newUndervoltageRelease = s.UndervoltageRelease.Article,
                  newSignalContact = s.SignalContact.Article,
                  newAuxiliaryContact = s.AuxiliaryContact.Article,
                  newSignalOrAuxiliaryContact = s.SignalOrAuxiliaryContact.Article
              }).FirstOrDefault();

            return (new AdditionalDevices(data?.newVendor, data?.newShuntTrip24v, data?.newShuntTrip48v, data?.newShuntTrip230v,
                data?.newUndervoltageRelease, data?.newSignalContact, data?.newAuxiliaryContact, data?.newSignalOrAuxiliaryContact));
        }
    }}
