using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace LicenseKeyGeneratorApp
{
    public partial class LicenceForm : Form
    {
        // ��������� ���� (������ ��������� � ���������� �����)
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

                // �������� ����������� ������� ��� URL-������������
                string safeKey = base64Key.Replace('+', '-').Replace('/', '_');

                // ����������� ���� � �����������
                textBoxKey.Text = FormatKey(safeKey[..24]); // ����� ������ 24 �������
            }
        }

        private static string FormatKey(string key)
        {
            // ��������� ���� �� ������ �� 6 ��������
            return Regex.Replace(key, "(.{6})", "$1-").TrimEnd('-');
        }   
    }
}
