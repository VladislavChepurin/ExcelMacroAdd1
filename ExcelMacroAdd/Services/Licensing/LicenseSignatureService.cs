using System;
using System.Security.Cryptography;
using System.Text;

namespace ExcelMacroAdd.Services.Licensing
{
    /// <summary>
    /// Проверяет RSA-подпись лицензии с использованием публичного ключа.
    /// Приватный ключ находится только в LicenseKeyGeneratorApp.
    /// </summary>
    internal static class LicenseSignatureService
    {
        // Публичный RSA-ключ (2048 бит). Это НЕ секрет — он только проверяет подпись.
        private const string PublicKeyXml =
@"<RSAKeyValue>
  <Modulus>lU4f7dK3Vn8zW0K6d7m20+cK5KMe/RoAm0fKjgolBzNLKmTy4JXZpqDzmelAaM6RKNSzDxQksPeNvN2Ypzs8tEf3WuCtNL4XHOpAahrpi8s+lBKfdDGYaN/GrULtMfbcJPw6tyzpxSU+CqN9GwPeWOQ/62lXszKJLA+y8wm6/iC+bxcQR7aFbjgz74RalOG30OS0JP8NnUMwuZMvsPUZLm7s44biNsgjezJ8RdY/83XuS17kMrPVgTxIYtzhjvb+w7bt+kL86LP5rtCBB7L0+4UgB+qHmK3ITQ2ihvsm5V5A9gjhyTLQpnwT4VA0kWQTPAJWEA2wyfr2WXtxSaRCGw==</Modulus>
  <Exponent>AQAB</Exponent>
</RSAKeyValue>";

        /// <summary>
        /// Проверяет RSA-подпись строки.
        /// </summary>
        /// <param name="data">Подписываемая строка (без поля Signature)</param>
        /// <param name="signatureBase64">Подпись в формате Base64</param>
        /// <returns>true — подпись валидна</returns>
        public static bool VerifySignature(string data, string signatureBase64)
        {
            if (string.IsNullOrEmpty(data) || string.IsNullOrEmpty(signatureBase64))
                return false;

            try
            {
                byte[] dataBytes = Encoding.UTF8.GetBytes(data);
                byte[] signatureBytes = Convert.FromBase64String(signatureBase64);

                using (var rsa = new RSACryptoServiceProvider())
                {
                    rsa.FromXmlString(PublicKeyXml);
                    return rsa.VerifyData(dataBytes, CryptoConfig.MapNameToOID("SHA256"), signatureBytes);
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
