using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace LicenseKeyGeneratorApp
{
    public partial class LicenceForm : Form
    {
        // Секретный ключ (должен храниться в безопасном месте)
        private static readonly byte[] SecretKey = Encoding.UTF8.GetBytes("MySuperSecretKey@12345");

        public LicenceForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int currentYear = DateTime.Now.Year;
            string rawData = $"{textBoxUserName.Text.ToLower()}|{currentYear}";

            using (var hmac = new HMACSHA256(SecretKey))
            {
                byte[] hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(rawData));
                string base64Key = Convert.ToBase64String(hash);

                // Заменяем специальные символы для URL-безопасности
                string safeKey = base64Key.Replace('+', '-').Replace('/', '_');

                // Укорачиваем ключ и форматируем
                textBoxKey.Text = FormatKey(safeKey[..24]); // Берем первые 24 символа
            }
        }

        private static string FormatKey(string key)
        {
            // Разбиваем ключ на группы по 6 символов
            return Regex.Replace(key, "(.{6})", "$1-").TrimEnd('-');
        }   
    }
}
