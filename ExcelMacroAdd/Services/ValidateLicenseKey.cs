using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelMacroAdd.Services
{
    internal class ValidateLicenseKey: IValidateLicenseKey
    {
        private readonly byte[] SecretKey = Encoding.UTF8.GetBytes("MySuperSecretKey@12345");
        private readonly string key;
        private readonly string username;   
                
        public ValidateLicenseKey(string key)
        {
            this.key = key;
            username = Environment.UserName;
        }

        private string GenerateKey(string username)
        {
            int currentYear = DateTime.Now.Year;
            string rawData = $"{username.ToLower()}|{currentYear}";

            using (var hmac = new HMACSHA256(SecretKey))
            {
                byte[] hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(rawData));
                string base64Key = Convert.ToBase64String(hash);

                // Заменяем специальные символы для URL-безопасности
                string safeKey = base64Key.Replace('+', '-').Replace('/', '_');

                // Укорачиваем ключ и форматируем
                return FormatKey(safeKey.Substring(0, 24)); // Берем первые 24 символа
            }
        }

        private string FormatKey(string key)
        {
            // Разбиваем ключ на группы по 6 символов
            return Regex.Replace(key, "(.{6})", "$1-").TrimEnd('-');
        }

        public bool ValidateKey()
        {
            // Генерируем ключ с текущими данными и сравниваем
            string generatedKey = GenerateKey(username);
            return generatedKey.Equals(key);
        }
    }
}
