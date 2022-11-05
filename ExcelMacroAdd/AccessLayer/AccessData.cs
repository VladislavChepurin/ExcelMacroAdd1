using ExcelMacroAdd.AccessLayer.Interfaces;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.AccessLayer
{
    public class AccessData: ISelectionCircuitBreakerData, IJournalData, ISelectionTransformerData, ISelectionTwinBlockData
    {
        public AccessCircuitBreaker AccessCircuitBreaker { get; set; }
        public AccessJournalNku AccessJournalNku { get; set; }
        public AccessTransformer AccessTransformer { get; set; }
        public AccessTwinBlock AccessTwinBlock { get; set; }

        public AccessData(AppContext context)
        {
            AccessCircuitBreaker = new AccessCircuitBreaker(context);
            AccessJournalNku = new AccessJournalNku(context);
            AccessTransformer = new AccessTransformer(context);
            AccessTwinBlock = new AccessTwinBlock(context);
        }
    }
}
