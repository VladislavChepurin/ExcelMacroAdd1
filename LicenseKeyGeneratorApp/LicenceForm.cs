using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using LicenseKeyGeneratorApp.Licensing;

namespace LicenseKeyGeneratorApp
{
    public partial class LicenceForm : Form
    {
        public LicenceForm()
        {
            InitializeComponent();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            // Валидация полей
            if (string.IsNullOrWhiteSpace(txtProduct.Text))
            {
                MessageBox.Show("Укажите название продукта", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtOwner.Text))
            {
                MessageBox.Show("Укажите владельца лицензии", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtOrganization.Text))
            {
                MessageBox.Show("Укажите организацию", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dtpValidFrom.Value.Date > dtpValidTo.Value.Date)
            {
                MessageBox.Show("Дата начала не может быть позже даты окончания", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 1. Собираем LicenseInfo
                var license = new LicenseInfo
                {
                    Product = txtProduct.Text.Trim(),
                    LicenseOwner = txtOwner.Text.Trim(),
                    Organization = txtOrganization.Text.Trim(),
                    ValidFrom = dtpValidFrom.Value.Date,
                    ValidTo = dtpValidTo.Value.Date
                };

                // 2. Формируем подписываемую строку
                string signableString = license.GetSignableString();

                // 3. Подписываем приватным ключом
                license.Signature = LicenseSignatureService.SignData(signableString);

                // 4. Показываем подписываемую строку для контроля
                txtSignableString.Text = signableString;
                txtSignature.Text = license.Signature;

                // 5. Сохраняем в файл
                using (var dialog = new SaveFileDialog())
                {
                    dialog.FileName = "license.json";
                    dialog.Filter = "JSON файлы (*.json)|*.json|Все файлы (*.*)|*.*";
                    dialog.DefaultExt = "json";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        var options = new JsonSerializerOptions
                        {
                            WriteIndented = true,
                            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                        };

                        string json = JsonSerializer.Serialize(license, options);
                        File.WriteAllText(dialog.FileName, json);

                        MessageBox.Show(
                            $"Лицензия сохранена:\n{dialog.FileName}\n\nДействует с {license.ValidFrom:dd.MM.yyyy} по {license.ValidTo:dd.MM.yyyy}",
                            "Готово",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка генерации лицензии:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
