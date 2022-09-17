using ExcelMacroAdd.Forms;
using ExcelMacroAdd.Interfaces;
using System.Threading;

namespace ExcelMacroAdd.Functions
{
    internal class FillingOutThePassport : AbstractFunctions
    {
        private readonly IResources resources;

        public FillingOutThePassport(IResources resources)
        {
            this.resources = resources;
        }

        public sealed override void Start()
        {
            if (application.ActiveWorkbook.Name != resources.NameFileJornal) // Проверка по имени книги
            {
                MessageWarning("Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                            "Имя книги не совпадает с целевой");
                return;
            }
            new Thread(() =>
            {
                Form1 fs = new Form1(resources);
                fs.ShowDialog();
                Thread.Sleep(100);
            }).Start();
        }
    }
}
