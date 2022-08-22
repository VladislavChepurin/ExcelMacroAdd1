using System;

namespace ExcelMacroAdd.Serializable
{
    [Serializable]
    public class StringResourcesForm2
    {
        public string[] CircuitBreakerCurrent { get; set;}
        public string[] CircuitBreakerCurve { get; set; }
        public string[] MaxCircuitBreakerCurrent { get; set; }
        public string[] AmountOfPolesCircuitBreaker { get; set; }
        public string[] CircuitBreakerVendor { get; set; }
        public string[] LoadSwitchCurrent { get; set; }
        public string[] AmountOfPolesLoadSwitch { get; set; }
        public string[] LoadSwitchVendor { get; set; }

        public StringResourcesForm2(string[] circuitBreakerCurrent, string[] circuitBreakerCurve, string[] maxCircuitBreakerCurrent, string[] amountOfPolesCircuitBreaker,
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
