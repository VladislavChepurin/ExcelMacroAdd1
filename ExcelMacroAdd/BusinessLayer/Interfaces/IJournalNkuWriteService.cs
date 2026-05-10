using ExcelMacroAdd.BusinessLayer.Models;
using System.Threading.Tasks;

namespace ExcelMacroAdd.BusinessLayer.Interfaces
{
    public interface IJournalNkuWriteService
    {
        Task<JournalNkuWriteResult> AddBoxAsync(JournalNkuWriteRequest request);

        Task<JournalNkuWriteResult> UpdateBoxAsync(JournalNkuWriteRequest request);
    }
}
