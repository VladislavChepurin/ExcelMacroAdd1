using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Services.Interfaces;
using ExcelMacroAdd.UserVariables;
using Microsoft.Office.Interop.Excel;
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Label = System.Windows.Forms.Label;
using TextBox = System.Windows.Forms.TextBox;

namespace ExcelMacroAdd.Forms
{
    internal partial class Settings : Form
    {
        private enum RowsToArray
        {
            IekLine,
            EkfLine,
            DkcLine,
            KeazLine,
            DekraftLine,
            TdmLine,
            AbbLine,
            SchneiderLine,
            ChintLine
        }

        private sealed class VendorSettingsRow
        {
            public VendorSettingsRow(
                string vendorName,
                TextBox formula1TextBox,
                TextBox formula2TextBox,
                TextBox formula3TextBox,
                TextBox discountTextBox,
                Label dateLabel)
            {
                VendorName = vendorName;
                Formula1TextBox = formula1TextBox;
                Formula2TextBox = formula2TextBox;
                Formula3TextBox = formula3TextBox;
                DiscountTextBox = discountTextBox;
                DateLabel = dateLabel;
            }

            public string VendorName { get; }

            private TextBox Formula1TextBox { get; }
            private TextBox Formula2TextBox { get; }
            private TextBox Formula3TextBox { get; }
            private TextBox DiscountTextBox { get; }
            private Label DateLabel { get; }

            public string Formula1
            {
                get => Formula1TextBox.Text ?? string.Empty;
                set => Formula1TextBox.Text = value ?? string.Empty;
            }

            public string Formula2
            {
                get => Formula2TextBox.Text ?? string.Empty;
                set => Formula2TextBox.Text = value ?? string.Empty;
            }

            public string Formula3
            {
                get => Formula3TextBox.Text ?? string.Empty;
                set => Formula3TextBox.Text = value ?? string.Empty;
            }

            public string Discount
            {
                get => DiscountTextBox.Text ?? string.Empty;
                set => DiscountTextBox.Text = value ?? string.Empty;
            }

            public string Date
            {
                get => DateLabel.Text ?? string.Empty;
                set => DateLabel.Text = value ?? string.Empty;
            }

            public void Apply(Vendor vendor)
            {
                Formula1 = vendor.Formula_1;
                Formula2 = vendor.Formula_2;
                Formula3 = vendor.Formula_3;
                Discount = vendor.Discount.ToString();
                Date = vendor.Date;
            }
        }

        private static readonly CultureInfo RussianCulture = CultureInfo.GetCultureInfo("ru-RU");

        private readonly IDataInXml dataInXml;
        private readonly VendorSettingsRow[] vendorRows;

        internal Settings(IDataInXml dataInXml)
        {
            InitializeComponent();
            this.dataInXml = dataInXml;
            vendorRows = CreateVendorRows();
        }

        #region KeyPress

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e) => HandleDiscountKeyPress(e);

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e) => HandleDiscountKeyPress(e);

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e) => HandleDiscountKeyPress(e);

        private void textBox16_KeyPress(object sender, KeyPressEventArgs e) => HandleDiscountKeyPress(e);

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e) => HandleDiscountKeyPress(e);

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e) => HandleDiscountKeyPress(e);

        private void textBox28_KeyPress(object sender, KeyPressEventArgs e) => HandleDiscountKeyPress(e);

        private void textBox32_KeyPress(object sender, KeyPressEventArgs e) => HandleDiscountKeyPress(e);

        private void textBox36_KeyPress(object sender, KeyPressEventArgs e) => HandleDiscountKeyPress(e);

        #endregion

        private void Settings_Load(object sender, EventArgs e)
        {
            try
            {
                // Загружаем в форму файл Settings.xml
                foreach (Vendor vendor in dataInXml.ReadFileXml())
                {
                    VendorSettingsRow vendorRow = FindVendorRow(vendor.VendorAttribute)
                        ?? throw new NullReferenceException("Не коректное значение в классе Form3");

                    vendorRow.Apply(vendor);
                }
            }
            catch (NullReferenceException)
            {
                MessageBox.Show(
                $@"Внимание! Возникла ошибка в файле Settings.xml,{Environment.NewLine} файл будет восстановлен автоматически.",
                @"Ошибка файла Settings.xml",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private void ReadExcelFunc(RowsToArray vendorLine)
        {
            Worksheet worksheet = null;
            Range activeCell = null;
            Range formula1Cell = null;
            Range formula2Cell = null;
            Range salesCell = null;
            Range formula3Cell = null;

            try
            {
                worksheet = Globals.ThisAddIn.GetActiveWorksheet();
                activeCell = Globals.ThisAddIn.GetActiveCell();

                VendorSettingsRow vendorRow = GetVendorRow(vendorLine);
                int currentRow = activeCell.Row;

                // Read Cells "B_" if value not empty then continue our work
                formula1Cell = (Range)worksheet.Cells[currentRow, 2];
                string formula1 = formula1Cell.FormulaLocal as string;
                if (!string.IsNullOrEmpty(formula1))
                {
                    vendorRow.Formula1 = VprFormulaReplace(formula1, currentRow);
                }

                // Read Cells "D_" if value not empty then continue our work
                formula2Cell = (Range)worksheet.Cells[currentRow, 4];
                string formula2 = formula2Cell.FormulaLocal as string;
                if (!string.IsNullOrEmpty(formula2))
                {
                    vendorRow.Formula2 = VprFormulaReplace(formula2, currentRow);
                }

                // Read Cells "F_"
                salesCell = (Range)worksheet.Cells[currentRow, 6];
                object sales = salesCell.Value2;
                if (sales is double)
                {
                    vendorRow.Discount = sales.ToString();
                }

                // Read Cells "G_" if value not empty then continue our work
                formula3Cell = (Range)worksheet.Cells[currentRow, 7];
                string formula3 = formula3Cell.FormulaLocal as string;
                if (!string.IsNullOrEmpty(formula3))
                {
                    vendorRow.Formula3 = VprFormulaReplace(formula3, currentRow);
                }
            }
            finally
            {
                ReleaseComObject(formula3Cell);
                ReleaseComObject(salesCell);
                ReleaseComObject(formula2Cell);
                ReleaseComObject(formula1Cell);
                ReleaseComObject(activeCell);
                ReleaseComObject(worksheet);
            }
        }

        /// <summary>
        /// Фнкцция замены для ВПР при считывании
        /// </summary>
        /// <param name="mReplase"></param>
        /// <param name="rows"></param>
        /// <returns></returns>
        public static string VprFormulaReplace(string mReplase, int rows)
        {
            return mReplase.Replace("=ВПР(A" + rows.ToString(), "=ВПР(A{0}");
        }

        /// <summary>
        /// Общий метод записи настроек вендора в XML.
        /// После записи сбрасывает кэш пересчитанных листов,
        /// чтобы следующая вставка формул выполнила полный пересчёт
        /// (формулы ссылаются на прайс, который мог измениться).
        /// </summary>
        private void WriteVendorSettings(RowsToArray vendorLine)
        {
            VendorSettingsRow vendorRow = GetVendorRow(vendorLine);
            string localDateText = DateTime.Now.ToString(RussianCulture);

            dataInXml.WriteXml(
                vendorRow.VendorName,
                vendorRow.Formula1,
                vendorRow.Formula2,
                vendorRow.Formula3,
                vendorRow.Discount,
                localDateText);

            vendorRow.Date = localDateText;

            // Формулы ВПР изменились → кэш пересчёта больше не валиден.
            // Следующий ExcelPerformanceScope выполнит полный Calculate().
            ExcelPerformanceScope.InvalidateCache();
        }

        private static void HandleDiscountKeyPress(KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private VendorSettingsRow[] CreateVendorRows()
        {
            return new[]
            {
                new VendorSettingsRow("IEK", textBox1, textBox2, textBox3, textBox4, label33),
                new VendorSettingsRow("EKF", textBox5, textBox6, textBox7, textBox8, label34),
                new VendorSettingsRow("DKC", textBox9, textBox10, textBox11, textBox12, label35),
                new VendorSettingsRow("KEAZ", textBox13, textBox14, textBox15, textBox16, label36),
                new VendorSettingsRow("DEKraft", textBox17, textBox18, textBox19, textBox20, label37),
                new VendorSettingsRow("TDM", textBox21, textBox22, textBox23, textBox24, label38),
                new VendorSettingsRow("ABB", textBox25, textBox26, textBox27, textBox28, label39),
                new VendorSettingsRow("Schneider", textBox29, textBox30, textBox31, textBox32, label40),
                new VendorSettingsRow("Chint", textBox33, textBox34, textBox35, textBox36, label41)
            };
        }

        private VendorSettingsRow GetVendorRow(RowsToArray vendorLine)
        {
            return vendorRows[(int)vendorLine];
        }

        private VendorSettingsRow FindVendorRow(string vendorName)
        {
            foreach (VendorSettingsRow vendorRow in vendorRows)
            {
                if (string.Equals(vendorRow.VendorName, vendorName, StringComparison.OrdinalIgnoreCase))
                {
                    return vendorRow;
                }
            }

            return null;
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }

        #region Write buttons — сохранение настроек вендоров в XML

        /// <summary>
        /// Write IEK settings to xml
        /// </summary>
        private void button2_Click(object sender, EventArgs e) => WriteVendorSettings(RowsToArray.IekLine);

        /// <summary>
        /// Write EKF settings to xml
        /// </summary>
        private void button4_Click(object sender, EventArgs e) => WriteVendorSettings(RowsToArray.EkfLine);

        /// <summary>
        /// Write DKC settings to xml
        /// </summary>
        private void button6_Click(object sender, EventArgs e) => WriteVendorSettings(RowsToArray.DkcLine);

        /// <summary>
        /// Write KEAZ settings to xml
        /// </summary>
        private void button8_Click(object sender, EventArgs e) => WriteVendorSettings(RowsToArray.KeazLine);

        /// <summary>
        /// Write DEKraft settings to xml
        /// </summary>
        private void button10_Click(object sender, EventArgs e) => WriteVendorSettings(RowsToArray.DekraftLine);

        /// <summary>
        /// Write TDM settings to xml
        /// </summary>
        private void button12_Click(object sender, EventArgs e) => WriteVendorSettings(RowsToArray.TdmLine);

        /// <summary>
        /// Write ABB settings to xml
        /// </summary>
        private void button14_Click(object sender, EventArgs e) => WriteVendorSettings(RowsToArray.AbbLine);

        /// <summary>
        /// Write Schneider settings to xml
        /// </summary>
        private void button16_Click(object sender, EventArgs e) => WriteVendorSettings(RowsToArray.SchneiderLine);

        /// <summary>
        /// Write Chint settings to xml
        /// </summary>
        private void button18_Click(object sender, EventArgs e) => WriteVendorSettings(RowsToArray.ChintLine);

        #endregion

        #region Read buttons — считывание формул с листа Excel

        /// <summary>
        /// Read IEK formula in ExcelSheets
        /// </summary>
        private void button1_Click(object sender, EventArgs e) => ReadExcelFunc(RowsToArray.IekLine);

        /// <summary>
        /// Read EKF formula in ExcelSheets
        /// </summary>
        private void button3_Click(object sender, EventArgs e) => ReadExcelFunc(RowsToArray.EkfLine);

        /// <summary>
        /// Read DKC formula in ExcelSheets
        /// </summary>
        private void button5_Click(object sender, EventArgs e) => ReadExcelFunc(RowsToArray.DkcLine);

        /// <summary>
        /// Read KEAZ formula in ExcelSheets
        /// </summary>
        private void button7_Click(object sender, EventArgs e) => ReadExcelFunc(RowsToArray.KeazLine);

        /// <summary>
        /// Read DEKraft formula in ExcelSheets
        /// </summary>
        private void button9_Click(object sender, EventArgs e) => ReadExcelFunc(RowsToArray.DekraftLine);

        /// <summary>
        /// Read TDM formula in ExcelSheets
        /// </summary>
        private void button11_Click(object sender, EventArgs e) => ReadExcelFunc(RowsToArray.TdmLine);

        /// <summary>
        /// Read ABB formula in ExcelSheets
        /// </summary>
        private void button13_Click(object sender, EventArgs e) => ReadExcelFunc(RowsToArray.AbbLine);

        /// <summary>
        /// Read Schneider formula in ExcelSheets
        /// </summary>
        private void button15_Click(object sender, EventArgs e) => ReadExcelFunc(RowsToArray.SchneiderLine);

        /// <summary>
        /// Read Chint formula in ExcelSheets
        /// </summary>
        private void button17_Click(object sender, EventArgs e) => ReadExcelFunc(RowsToArray.ChintLine);

        #endregion
    }
}
