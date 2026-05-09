using System;
using System.Security.Cryptography;
using System.Text;

namespace LicenseKeyGeneratorApp.Licensing
{
    /// <summary>
    /// Подписывает лицензию приватным RSA-ключом.
    /// Приватный ключ хранится ТОЛЬКО здесь, в генераторе.
    /// </summary>
    internal static class LicenseSignatureService
    {
        // Приватный RSA-ключ (2048 бит). Хранить только в генераторе!
        private const string PrivateKeyXml =
@"<RSAKeyValue>
  <Modulus>lU4f7dK3Vn8zW0K6d7m20+cK5KMe/RoAm0fKjgolBzNLKmTy4JXZpqDzmelAaM6RKNSzDxQksPeNvN2Ypzs8tEf3WuCtNL4XHOpAahrpi8s+lBKfdDGYaN/GrULtMfbcJPw6tyzpxSU+CqN9GwPeWOQ/62lXszKJLA+y8wm6/iC+bxcQR7aFbjgz74RalOG30OS0JP8NnUMwuZMvsPUZLm7s44biNsgjezJ8RdY/83XuS17kMrPVgTxIYtzhjvb+w7bt+kL86LP5rtCBB7L0+4UgB+qHmK3ITQ2ihvsm5V5A9gjhyTLQpnwT4VA0kWQTPAJWEA2wyfr2WXtxSaRCGw==</Modulus>
  <Exponent>AQAB</Exponent>
  <P>yfNqJaJs7ofYnhYd1zV50YrApHfHwH1EsJkw9ZIHO8dOW3wZznC9GMD+nYP2uSp0bzjsuCgdBWU3Ncv5s720Tb0jDP3qUoc/1a+n2AFYh6rJnsa+mOrpvesm3pWQPVm9Do9qKPE6/+uHxP0f+onpsP/KPixJs7RXXXrgRaOyVrM=</P>
  <Q>vUO2W4eWSM0ZeRQil7q5hjI7kyNRn7jBCJpP8fN1Y5EJ7nbhb9+bQOKz5BZZf0JIxRBhPV6ax1kOPZjzh5lLHHmeRx1YyQBE5ODmhEE0B4x9c16K2Wa0Ftd4WT3Jg/8g3peGUbgXys9q0EI4G3tSjVjpDYiP/VR8qvgi9dgBWvk=</Q>
  <DP>ipYea8ExG+fhgWsQA1XRSTj8xmDklXXho4cdEAisKhu17BYX55F6UvhuQk4DDELUMFdSK3Zro/43ixV1QCGZEBgRa6L8ILJr3gpzFkqmJEPRpMIinfHngctTmz/sAg4JLWrBoWMZ5/IL8+T5AweNdUez1EK0OTwzEBV4vpF9mv0=</DP>
  <DQ>h2ziGVBFesY3SencbtFPWvSqqDgHedBLX4p7Vdcs0hfAEX/DA7fucVlF+xj65RJa25dC3RTKj4XrqKu+5fIMSs3DMYOQOhMVOOisSUoWnqgqQ9kMZU8V4ZpAJSsO/IIb1Op7VBH0BEyyU15uo0t04GsUJ3jl/xDrO7Ld4Se0oJE=</DQ>
  <InverseQ>kWK8iBWce3msCspqGRvuXUAGDqTnn+KOD2fdb9g/0vk+GxwEoQBy1vIhQVikldhe5/bzu6D4z0rL9CTPgY+IMv2YkqvlkJ/8ludSkLhHju8K+ggObAYGUIu7gtk2ltlOsPQFggqGkSC0IbUEELckJgPLLw8i8Pkfry8ah0Fpum8=</InverseQ>
  <D>Bhaj/a5AlLHmNbwAX55+oqCC1LUEL/0N9kcUrvsh7Gu+jnGEZ/0kXYOlu2qEGmIGEFywGpbPMjo+GOwObA9h19Yxc47C8WopBiBVVR5Y1L8Kg75Iq1PUa75oWytmAcoXyxhQCqU1uTjeEU/+a4oaWJSiOKbYkTGn31iaiwekDPXX6hvPVT45GpW92Zt98f68Y0NxRKxNiG68v4ZiVAd7LGxK47VGD0ESwYtx/Omfx/Ijf68SFIhYER9Aq0RJCnI5qA/h4PKdTbnAAKtg1XtN3hVLuqin9fhkDYqYJlAYFIUtJb5Qgj+/7g1y0QcNs9hx+X0BJpv1u3orbGENjKLZOQ==</D>
</RSAKeyValue>";

        /// <summary>
        /// Подписывает строку приватным RSA-ключом.
        /// </summary>
        /// <param name="data">Подписываемая строка</param>
        /// <returns>Подпись в формате Base64</returns>
        public static string SignData(string data)
        {
            byte[] dataBytes = Encoding.UTF8.GetBytes(data);

            using (var rsa = new RSACryptoServiceProvider())
            {
                rsa.FromXmlString(PrivateKeyXml);
                byte[] signature = rsa.SignData(dataBytes, CryptoConfig.MapNameToOID("SHA256"));
                return Convert.ToBase64String(signature);
            }
        }
    }
}
