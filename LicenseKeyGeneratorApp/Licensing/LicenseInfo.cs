using System;

namespace LicenseKeyGeneratorApp.Licensing
{
    public class LicenseInfo
    {
        public string? Product { get; set; }
        public string? LicenseOwner { get; set; }
        public string? Organization { get; set; }
        public DateTime ValidFrom { get; set; }
        public DateTime ValidTo { get; set; }
        public string? Signature { get; set; }

        /// <summary>
        /// Формирует строку для подписи.
        /// Формат строго фиксирован и должен совпадать с ExcelMacroAdd:
        /// Product|LicenseOwner|Organization|ValidFrom:yyyy-MM-dd|ValidTo:yyyy-MM-dd
        /// </summary>
        public string GetSignableString()
        {
            return string.Join("|",
                Product,
                LicenseOwner,
                Organization,
                "ValidFrom:" + ValidFrom.ToString("yyyy-MM-dd"),
                "ValidTo:" + ValidTo.ToString("yyyy-MM-dd"));
        }
    }
}
