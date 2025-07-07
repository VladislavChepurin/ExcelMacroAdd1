using ExcelMacroAdd.BisinnesLayer.Interfaces;
using Microsoft.Extensions.Caching.Memory;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.BisinnesLayer
{
    public class AccessData : ISelectionSwitchData, ISelectionCircuitBreakerData, IJournalData, ISelectionTransformerData, ISelectionTwinBlockData, ITermoCalcData, IAdditionalModularDevicesData, INotPriceComponent
    {
        public AccessCircuitBreaker AccessCircuitBreaker { get; set; }
        public AccessSwitch AccessSwitch { get; set; }
        public AccessAdditionalModularDevices AccessAdditionalModularDevices { get; set; }
        public AccessJournalNku AccessJournalNku { get; set; }
        public AccessTransformer AccessTransformer { get; set; }
        public AccessTwinBlock AccessTwinBlock { get; set; }
        public AccessTermoCalc AccessTermoCalc { get; set; }
        public AccessNotPriceComponent AccessNotPriceComponent { get; set; }

        public AccessData(AppContext context, IMemoryCache memoryCache)
        {
            AccessCircuitBreaker = new AccessCircuitBreaker(context, memoryCache);
            AccessSwitch = new AccessSwitch(context, memoryCache);
            AccessAdditionalModularDevices = new AccessAdditionalModularDevices(context);
            AccessJournalNku = new AccessJournalNku(context, memoryCache);
            AccessTransformer = new AccessTransformer(context);
            AccessTwinBlock = new AccessTwinBlock(context);
            AccessTermoCalc = new AccessTermoCalc(context);
            AccessNotPriceComponent = new AccessNotPriceComponent(context, memoryCache);
        }
    }
}
