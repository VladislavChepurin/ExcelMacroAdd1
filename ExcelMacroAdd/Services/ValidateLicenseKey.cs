using ExcelMacroAdd.Services.Interfaces;
using ExcelMacroAdd.Services.Licensing;
using Newtonsoft.Json;
using System;
using System.IO;

namespace ExcelMacroAdd.Services
{
    internal class ValidateLicenseKey : IValidateLicenseKey
    {
        private const string ExpectedProduct = "ExcelMacroAdd";
        private readonly string _licenseFilePath;

        public ValidateLicenseKey()
            : this(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "license.json"))
        {
        }

        public ValidateLicenseKey(string licenseFilePath)
        {
            _licenseFilePath = licenseFilePath;
        }

        public bool ValidateKey()
        {
            try
            {
                // 1. Проверяем наличие файла
                if (!File.Exists(_licenseFilePath))
                {
                    Logger.Log("Файл лицензии не найден: " + _licenseFilePath, Logger.LogLevel.Warning);
                    return false;
                }

                // 2. Читаем и десериализуем JSON
                string json = File.ReadAllText(_licenseFilePath);
                var license = JsonConvert.DeserializeObject<LicenseInfo>(json);

                if (license == null)
                {
                    Logger.Log("Не удалось прочитать файл лицензии", Logger.LogLevel.Warning);
                    return false;
                }

                // 3. Проверяем продукт
                if (!string.Equals(license.Product, ExpectedProduct, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Log("Неверный продукт в лицензии: " + license.Product, Logger.LogLevel.Warning);
                    return false;
                }

                // 4. Проверяем даты
                var today = DateTime.Today;

                if (today < license.ValidFrom.Date)
                {
                    Logger.Log("Лицензия ещё не действует. ValidFrom: " + license.ValidFrom.ToString("yyyy-MM-dd"), Logger.LogLevel.Warning);
                    return false;
                }

                if (today > license.ValidTo.Date)
                {
                    Logger.Log("Срок лицензии истёк. ValidTo: " + license.ValidTo.ToString("yyyy-MM-dd"), Logger.LogLevel.Warning);
                    return false;
                }

                // 5. Проверяем подпись
                if (string.IsNullOrWhiteSpace(license.Signature))
                {
                    Logger.Log("Подпись лицензии пуста", Logger.LogLevel.Warning);
                    return false;
                }

                string signableString = license.GetSignableString();
                bool signatureValid = LicenseSignatureService.VerifySignature(signableString, license.Signature);

                if (!signatureValid)
                {
                    Logger.Log("Подпись лицензии недействительна", Logger.LogLevel.Warning);
                    return false;
                }

                return true;
            }
            catch (JsonException ex)
            {
                Logger.LogException(ex, "Ошибка формата license.json");
                return false;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Ошибка проверки лицензии");
                return false;
            }
        }
    }
}
