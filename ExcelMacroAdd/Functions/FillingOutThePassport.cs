using ExcelMacroAdd.Forms;
using ExcelMacroAdd.Interfaces;
using System.Threading;

namespace ExcelMacroAdd.Functions
{
    internal class FillingOutThePassport : AbstractFunctions
    {
        private IDBConect dBConect;

        public FillingOutThePassport(IDBConect dBConect)
        {
            this.dBConect = dBConect;
        }

        protected internal override void Start()
        {
            dBConect?.OpenDB();
            if (application.ActiveWorkbook.Name != dBConect?.ReadOnlyOneNoteDB("SELECT * FROM settings WHERE set_name = 'sJornal';", 2)) // Проверка по имени книги
            {
                MessageWarning("Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                            "Имя книги не совпадает с целевой");
                dBConect?.CloseDB();
                return;
            }
            new Thread(() =>
            {
                Form1 fs = new Form1(dBConect);
                fs.ShowDialog();
                Thread.Sleep(100);
            }).Start();

            dBConect.CloseDB();
        }
    }
}
