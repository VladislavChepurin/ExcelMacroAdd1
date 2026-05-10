using ExcelMacroAdd.BusinessLayer.Models;
using ExcelMacroAdd.DataLayer.Entity;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BusinessLayer.Interfaces
{
    public interface INotPriceComponentsService
    {
        Task<IList<NotPriceComponent>> GetAllRecordsAsync();

        Task<bool> RecordExistsAsync(string article);

        Task<bool> VendorExistsAsync(string vendorName);

        Task<NotPriceComponent> AddRecordAsync(NotPriceComponentSaveRequest request, bool createVendorIfMissing);

        Task<NotPriceComponent> UpdateRecordAsync(NotPriceComponentSaveRequest request, bool createVendorIfMissing);

        Task<NotPriceComponent> SetRecordStateAsync(string article, int? status);

        Task<bool> DeleteRecordAsync(int id);
    }
}
