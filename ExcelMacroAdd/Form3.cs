using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Label = System.Windows.Forms.Label;
using TextBox = System.Windows.Forms.TextBox;

namespace ExcelMacroAdd
{
    enum RowsToArray
    {
        IekLine,
        EkfLine,
        DkcLine,
        KeazLine,
        DekraftLine,
        TdmLine,
        AbbLine,
        SchneiderLine
    }

    public partial class Form3 : Form
    {
        // Folders AppData content Settings.xml
        readonly string file = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\AddIns\ExcelMacroAdd\Settings.xml";

        public Form3()
        {
            InitializeComponent();
        }

        #region KeyPress

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox16_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox28_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox32_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }


        #endregion


        private Label[] RetupnLabelArray()
        {
            Label[] labels = new Label[] { label33, label34, label35, label36, label37, label38, label39, label40 };
            return labels;
        }

        private TextBox[,] ReturnTextBoxArray()
        {
            TextBox[,] textBoxes = new TextBox[8, 4]
            {
                {
                    textBox1, textBox2, textBox3, textBox4      //IEK
                },
                {
                    textBox5, textBox6, textBox7, textBox8      //EKF
                },
                {
                    textBox9, textBox10, textBox11, textBox12   //DKC
                },
                {
                    textBox13, textBox14, textBox15, textBox16  //KEAZ
                },
                {
                    textBox17, textBox18, textBox19, textBox20  //DEKraft
                },
                {
                   textBox21, textBox22, textBox23, textBox24   //TDM
                },
                {
                   textBox25, textBox26, textBox27, textBox28   //ABB
                },
                {
                   textBox29, textBox30, textBox31, textBox32   //Schneider
                }
            };
            return textBoxes;
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            try
            { // преписать!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                XDocument xdoc = XDocument.Load(file);
                // получаем корневой узел
                XElement dataRow = xdoc.Element("MetaSettings");
                if (dataRow != null)
                {
                    foreach (XElement vendor in dataRow.Elements("Vendor"))
                    {
                        switch ((vendor.Attribute("vendor")).Value)
                        {
                            case "IEK":
                                textBox1.Text = vendor.Element("Formula_1").Value;
                                textBox2.Text = vendor.Element("Formula_2").Value;
                                textBox3.Text = vendor.Element("Formula_3").Value;
                                textBox4.Text = vendor.Element("Discont").Value;
                                label33.Text = vendor.Element("Date").Value;
                                break;

                            case "EKF":
                                textBox5.Text = vendor.Element("Formula_1").Value;
                                textBox6.Text = vendor.Element("Formula_2").Value;
                                textBox7.Text = vendor.Element("Formula_3").Value;
                                textBox8.Text = vendor.Element("Discont").Value;
                                label34.Text = vendor.Element("Date").Value;
                                break;

                            case "DKC":
                                textBox9.Text = vendor.Element("Formula_1").Value;
                                textBox10.Text = vendor.Element("Formula_2").Value;
                                textBox11.Text = vendor.Element("Formula_3").Value;
                                textBox12.Text = vendor.Element("Discont").Value;
                                label35.Text = vendor.Element("Date").Value;
                                break;

                            case "KEAZ":
                                textBox13.Text = vendor.Element("Formula_1").Value;
                                textBox14.Text = vendor.Element("Formula_2").Value;
                                textBox15.Text = vendor.Element("Formula_3").Value;
                                textBox16.Text = vendor.Element("Discont").Value;
                                label36.Text = vendor.Element("Date").Value;
                                break;

                            case "DEKraft":
                                textBox17.Text = vendor.Element("Formula_1").Value;
                                textBox18.Text = vendor.Element("Formula_2").Value;
                                textBox19.Text = vendor.Element("Formula_3").Value;
                                textBox20.Text = vendor.Element("Discont").Value;
                                label37.Text = vendor.Element("Date").Value;
                                break;

                            case "TDM":
                                textBox21.Text = vendor.Element("Formula_1").Value;
                                textBox22.Text = vendor.Element("Formula_2").Value;
                                textBox23.Text = vendor.Element("Formula_3").Value;
                                textBox24.Text = vendor.Element("Discont").Value;
                                label38.Text = vendor.Element("Date").Value;
                                break;

                            case "ABB":
                                textBox25.Text = vendor.Element("Formula_1").Value;
                                textBox26.Text = vendor.Element("Formula_2").Value;
                                textBox27.Text = vendor.Element("Formula_3").Value;
                                textBox28.Text = vendor.Element("Discont").Value;
                                label39.Text = vendor.Element("Date").Value;
                                break;

                            case "Schneider":
                                textBox29.Text = vendor.Element("Formula_1").Value;
                                textBox30.Text = vendor.Element("Formula_2").Value;
                                textBox31.Text = vendor.Element("Formula_3").Value;
                                textBox32.Text = vendor.Element("Discont").Value;
                                label40.Text = vendor.Element("Date").Value;
                                break;
                        }
                    }
                }
            }

            catch (NullReferenceException)
            {
                //Здесь написать код починки файла Settings.xml
            }

            catch (FileNotFoundException) // Востановление файла при его удалении
            {
                DataInXml dataInXml = new DataInXml();
                dataInXml.XmlFileCreate();
            }
        }
        /// <summary>
        /// Функция записи в Xml, первый параметр - вендор в настройках, второй - номер строки в двумерном массиве TextBox[,]
        /// </summary>
        /// <param name="vendor"></param>
        /// <param name="rowsArray"></param>
        private void WiterXmlFunc(string vendor, int rowsArray)
        {
            TextBox[,] textBoxes = ReturnTextBoxArray();
            Label[] labels = RetupnLabelArray();
            XDocument xdoc = XDocument.Load(file);
            var index = xdoc.Element("MetaSettings")?.Elements("Vendor").FirstOrDefault(p => p.Attribute("vendor")?.Value == vendor);
            if (index != null)
            {
                // Записываем первую формулу
                var formula_1 = index.Element("Formula_1");
                if (formula_1 != null) formula_1.Value = textBoxes[rowsArray, 0].Text;
                // Записываем вторую формулу
                var formula_2 = index.Element("Formula_2");
                if (formula_2 != null) formula_2.Value = textBoxes[rowsArray, 1].Text;
                // Записываем третью формулу
                var formula_3 = index.Element("Formula_3");
                if (formula_3 != null) formula_3.Value = textBoxes[rowsArray, 2].Text;
                // Записываем скидку
                var discont = index.Element("Discont");
                if (discont != null) discont.Value = textBoxes[rowsArray, 3].Text;
                // Записываем дату и время
                DateTime localDate = DateTime.Now;
                var date = index.Element("Date");
                if (date != null) date.Value = localDate.ToString();
                // Записываем в форму дату обновления
                labels[rowsArray].Text = localDate.ToString();

                // Сохраняем документ
                xdoc.Save(file);
            }
        }

        private void ReadExcelFunc(int rowsArray)
        {
            Excel.Application application = Globals.ThisAddIn.GetApplication();
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Range cell = Globals.ThisAddIn.GetActiveCell();

            TextBox[,] textBoxes = ReturnTextBoxArray();

            int firstRow = cell.Row;

            // Read Cells "B_" if value not empty then continue our work
            string formula_1 = worksheet.Cells[firstRow, 2]?.FormulaLocal;                    
            if (formula_1 != String.Empty) 
            {
                textBoxes[rowsArray, 0].Text = Replace.VprFormulaReplace(formula_1, firstRow);
            }
            // Read Cells "D_" if value not empty then continue our work
            string formula_2 = worksheet.Cells[firstRow, 4]?.FormulaLocal;
            if (formula_2 != String.Empty)
            {
                textBoxes[rowsArray, 1].Text = Replace.VprFormulaReplace(formula_2, firstRow);
            }
            // Read Cells "G_" if value not empty then continue our work
            string formula_3 = worksheet.Cells[firstRow, 7]?.FormulaLocal;
            if (formula_3 != String.Empty)
            {
                textBoxes[rowsArray, 2].Text = Replace.VprFormulaReplace(formula_3, firstRow);
            }
        }


        /// <summary>
        /// Write IEK settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            WiterXmlFunc("IEK", (int)RowsToArray.IekLine);
        }

        /// <summary>
        /// Write EKF settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            WiterXmlFunc("EKF", (int)RowsToArray.EkfLine);
        }

        /// <summary>
        /// Write DKC settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            WiterXmlFunc("DKC", (int)RowsToArray.DkcLine);
        }

        /// <summary>
        /// Write KEAZ settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            WiterXmlFunc("KEAZ", (int)RowsToArray.KeazLine);
        }

        /// <summary>
        /// Write DEKraft settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button10_Click(object sender, EventArgs e)
        {
            WiterXmlFunc("DEKraft", (int)RowsToArray.DekraftLine);
        }

        /// <summary>
        /// Write TDM settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button12_Click(object sender, EventArgs e)
        {
            WiterXmlFunc("TDM", (int)RowsToArray.TdmLine);
        }

        /// <summary>
        /// Write ABB settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button14_Click(object sender, EventArgs e)
        {
            WiterXmlFunc("ABB", (int)RowsToArray.AbbLine);
        }

        /// <summary>
        /// Write Schneider settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button16_Click(object sender, EventArgs e)
        {
            WiterXmlFunc("Schneider", (int)RowsToArray.SchneiderLine);
        }
        /// <summary>
        /// Read IEK formula in ExcelSheets
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            ReadExcelFunc((int)RowsToArray.IekLine);
        }
        /// <summary>
        /// Read EKF formula in ExcelSheets
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            ReadExcelFunc((int)RowsToArray.EkfLine);
        }
        /// <summary>
        /// Read DKC formula in ExcelSheets
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            ReadExcelFunc((int)RowsToArray.DkcLine);
        }
        /// <summary>
        /// Read KEAZ formula in ExcelSheets
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            ReadExcelFunc((int)RowsToArray.KeazLine);
        }
        /// <summary>
        /// Read DEKraft formula in ExcelSheets
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            ReadExcelFunc((int)RowsToArray.DekraftLine);
        }
        /// <summary>
        /// Read TDM formula in ExcelSheets
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button11_Click(object sender, EventArgs e)
        {
            ReadExcelFunc((int)RowsToArray.TdmLine);
        }
        /// <summary>
        /// Read ABB formula in ExcelSheets
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button13_Click(object sender, EventArgs e)
        {
            ReadExcelFunc((int)RowsToArray.AbbLine);
        }
        /// <summary>
        /// Read Schneider formula in ExcelSheets
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button15_Click(object sender, EventArgs e)
        {
            ReadExcelFunc((int)RowsToArray.SchneiderLine);
        }
    }
}
