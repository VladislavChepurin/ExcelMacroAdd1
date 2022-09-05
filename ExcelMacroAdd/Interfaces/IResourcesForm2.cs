namespace ExcelMacroAdd.Interfaces
{
    public interface IResourcesForm2
    {
        string[] CircuitBreakerCurrent { get; set; }
        string[] CircuitBreakerCurve { get; set; }
        string[] MaxCircuitBreakerCurrent { get; set; }
        string[] AmountOfPolesCircuitBreaker { get; set; }
        string[] CircuitBreakerVendor { get; set; }
        string[] LoadSwitchCurrent { get; set; }
        string[] AmountOfPolesLoadSwitch { get; set; }
        string[] LoadSwitchVendor { get; set; }
    }
}
