using ExcelMacroAdd.Models;

namespace ExcelMacroAdd.BusinessLayer.Interfaces
{
    public interface IAdditionalDevicesService
    {
        AdditionalDevices GetEntityAdditionalCircuitBreaker(string articleNumber);

        AdditionalDevices GetEntityAdditionalSwitch(string articleNumber);
    }
}
