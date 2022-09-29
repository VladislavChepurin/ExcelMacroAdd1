using ExcelMacroAdd.Interfaces;
using System;

namespace ExcelMacroAdd.Serializable.Entity
{
    [Serializable]
    public class ResourcesForm2 : IResourcesForm2
    {
        public object[] CircuitBreakerCurrent { get; set;}
        public object[] CircuitBreakerCurve { get; set; }
        public object[] MaxCircuitBreakerCurrent { get; set; }
        public object[] AmountOfPolesCircuitBreaker { get; set; }
        public object[] CircuitBreakerVendor { get; set; }
        public object[] LoadSwitchCurrent { get; set; }
        public object[] AmountOfPolesLoadSwitch { get; set; }
        public object[] LoadSwitchVendor { get; set; }

        public ResourcesForm2(string[] circuitBreakerCurrent, string[] circuitBreakerCurve, string[] maxCircuitBreakerCurrent, string[] amountOfPolesCircuitBreaker,
                                    string[] circuitBreakerVendor, string[] loadSwitchCurrent, string[] amountOfPolesLoadSwitch, string[] loadSwitchVendor)
        {
            CircuitBreakerCurrent = circuitBreakerCurrent;
            CircuitBreakerCurve = circuitBreakerCurve;
            MaxCircuitBreakerCurrent = maxCircuitBreakerCurrent;
            AmountOfPolesCircuitBreaker = amountOfPolesCircuitBreaker;
            CircuitBreakerVendor = circuitBreakerVendor;
            LoadSwitchCurrent = loadSwitchCurrent;
            AmountOfPolesLoadSwitch = amountOfPolesLoadSwitch;
            LoadSwitchVendor = loadSwitchVendor;
        }
    }
}
