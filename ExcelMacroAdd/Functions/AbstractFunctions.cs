using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    public abstract class AbstractFunctions
    {
        internal const int NumberProjectColumn = 1;
        internal const int TitleProjectColumn = 2;
        internal const int NumberItemColumn = 3;
        internal const int TitleMRTColumn = 4;
        internal const int ModelColumn = 5;
        internal const int QuantityJournalColumn = 6;
        internal const int TitleLVSwitchgearColumn = 7;
        internal const int DesignationOfLVSwitchgearColumn = 8;
        internal const int TechnicalSpecificationsColumn = 9;
        internal const int VoltageColumn = 10;
        internal const int CurrentColumn = 11;
        internal const int IPRatingColumn = 12;
        internal const int ClimaticCategoryColumn = 13;
        internal const int MassColumn = 14;
        internal const int EnclosureHeightColumn = 15;
        internal const int EnclosureWidthColumn = 16;
        internal const int EnclosureDepthColumn = 17;
        internal const int ManufacturingDataColumn = 21;
        internal const int SerialNumberColumn = 22;
        internal const int CabinetArticleColumn = 27;
        internal const int ApparatusMountingColumn = 28;
        internal const int EarthingSystemColumn = 29;
        internal const int CabinetMaterialTypeColumn = 30;
        internal const int MountingTypeColumn = 31;

        /// <summary>
        /// COM-объекты Excel получаются как свойства, а не как поля-инициализаторы.
        /// 
        /// Раньше Worksheet, Cell, WorkBook захватывались в момент создания объекта (new CorrectDb(...)).
        /// Если между созданием и вызовом Start() пользователь переключал лист или менял выделение —
        /// код работал с устаревшими ссылками.
        /// 
        /// Для ViewModels (NotPriceComponentsViewModel, FillingOutPassportViewModel и др.) это
        /// особенно критично: VM создаётся при открытии формы и живёт, пока форма открыта.
        /// Worksheet/Cell должны отражать актуальное состояние на момент каждого действия.
        ///
        /// Application — единственный стабильный объект, он не меняется за сессию.
        /// Worksheet, Cell, WorkBook — каждый раз берутся актуальные.
        /// </summary>
        protected readonly Microsoft.Office.Interop.Excel.Application Application = Globals.ThisAddIn.GetApplication();
        protected Worksheet Worksheet => Globals.ThisAddIn.GetActiveWorksheet();
        protected Range Cell => Globals.ThisAddIn.GetActiveCell();
        protected Workbook WorkBook => Globals.ThisAddIn.GetActiveWorkBook();

        public virtual void Start() { }
        public virtual Task StartAsync() => Task.CompletedTask;

        protected internal void MessageInformation(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        protected internal void MessageWarning(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        protected internal void MessageError(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}