namespace ExcelMacroAdd.Interfaces
{
    public interface IResourcesForm2
    {
        object[] CircuitBreakerCurrent { get; set; }
        object[] CircuitBreakerCurve { get; set; }
        object[] MaxCircuitBreakerCurrent { get; set; }
        object[] AmountOfPolesCircuitBreaker { get; set; }
        object[] CircuitBreakerVendor { get; set; }
        object[] LoadSwitchCurrent { get; set; }
        object[] AmountOfPolesLoadSwitch { get; set; }
        object[] LoadSwitchVendor { get; set; }
    }
}
