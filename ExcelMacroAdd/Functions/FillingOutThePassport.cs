using ExcelMacroAdd.Forms;
using ExcelMacroAdd.Servises;
using System;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    internal class FillingOutThePassport : AbstractFunctions
    {
        private Lazy<DBConect> dBConect;

        public FillingOutThePassport(Lazy<DBConect> dBConect)
        {
            this.dBConect = dBConect;
        }

        public override void Start()
        {
            dBConect.Value.OpenDB();
            // Проверка по имени книги
            if (application.ActiveWorkbook.Name == dBConect.Value.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2))
            {
                new Thread(() =>
                {
                    Form1 fs = new Form1(dBConect);
                    fs.ShowDialog();
                    Thread.Sleep(100);
                }).Start();
            }
            else
            {
                MessageBox.Show(
                "Программа работает только в файле " + dBConect.Value.RequestDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2) + "\n Пожайлуста откройте целевую книгу и запустите программу.",
                "Ошибка вызова",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }
            dBConect.Value.CloseDB();
        }
    }
}
