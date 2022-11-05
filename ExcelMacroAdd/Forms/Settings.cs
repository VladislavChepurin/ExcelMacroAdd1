using ExcelMacroAdd.Interfaces;
using ExcelMacroAdd.UserVariables;
using Microsoft.Office.Interop.Excel;
using System;
using System.Globalization;
using System.Windows.Forms;
using Label = System.Windows.Forms.Label;
using TextBox = System.Windows.Forms.TextBox;

namespace ExcelMacroAdd.Forms
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

    internal partial class Settings : Form
    {    
        private readonly IDataInXml dataInXml;
        internal Settings(IDataInXml dataInXml)
        {
            InitializeComponent();
            this.dataInXml = dataInXml;
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

        private Label[] ReturnLabelArray()
        {
            Label[] labels = new Label[] { label33, label34, label35, label36, label37, label38, label39, label40 };
            return labels;
        }

        private TextBox[,] ReturnTextBoxArray()
        {
            TextBox[,] textBoxes =
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
        private void Settings_Load(object sender, EventArgs e)
        {
            try
            { // Загружаем в форму файл Settings.xml
                foreach (Vendor vendor in dataInXml.ReadFileXml())
                {
                    switch (vendor.VendorAttribute)
                    {
                        case "IEK":
                            textBox1.Text = vendor.Formula_1;
                            textBox2.Text = vendor.Formula_2;
                            textBox3.Text = vendor.Formula_3;
                            textBox4.Text = vendor.Discount.ToString();
                            label33.Text = vendor.Date;
                            break;
                        case "EKF":
                            textBox5.Text = vendor.Formula_1;
                            textBox6.Text = vendor.Formula_2;
                            textBox7.Text = vendor.Formula_3;
                            textBox8.Text = vendor.Discount.ToString();
                            label34.Text = vendor.Date;
                            break;
                        case "DKC":
                            textBox9.Text = vendor.Formula_1;
                            textBox10.Text = vendor.Formula_2;
                            textBox11.Text = vendor.Formula_3;
                            textBox12.Text = vendor.Discount.ToString();
                            label35.Text = vendor.Date;
                            break;
                        case "KEAZ":
                            textBox13.Text = vendor.Formula_1;
                            textBox14.Text = vendor.Formula_2;
                            textBox15.Text = vendor.Formula_3;
                            textBox16.Text = vendor.Discount.ToString();
                            label36.Text = vendor.Date;
                            break;
                        case "DEKraft":
                            textBox17.Text = vendor.Formula_1;
                            textBox18.Text = vendor.Formula_2;
                            textBox19.Text = vendor.Formula_3;
                            textBox20.Text = vendor.Discount.ToString();
                            label37.Text = vendor.Date;
                            break;
                        case "TDM":
                            textBox21.Text = vendor.Formula_1;
                            textBox22.Text = vendor.Formula_2;
                            textBox23.Text = vendor.Formula_3;
                            textBox24.Text = vendor.Discount.ToString();
                            label38.Text = vendor.Date;
                            break;
                        case "ABB":
                            textBox25.Text = vendor.Formula_1;
                            textBox26.Text = vendor.Formula_2;
                            textBox27.Text = vendor.Formula_3;
                            textBox28.Text = vendor.Discount.ToString();
                            label39.Text = vendor.Date;
                            break;
                        case "Schneider":
                            textBox29.Text = vendor.Formula_1;
                            textBox30.Text = vendor.Formula_2;
                            textBox31.Text = vendor.Formula_3;
                            textBox32.Text = vendor.Discount.ToString();
                            label40.Text = vendor.Date;
                            break;
                        default:
                            throw new NullReferenceException("Не коректное значение в классе Form3");
                    }
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
        private void ReadExcelFunc(int rowsArray)
        {
            Worksheet worksheet = Globals.ThisAddIn.GetActiveWorksheet();
            Range cell = Globals.ThisAddIn.GetActiveCell();

            TextBox[,] textBoxes = ReturnTextBoxArray();

            int firstRow = cell.Row;

            // Read Cells "B_" if value not empty then continue our work
            string formula1 = worksheet.Cells[firstRow, 2]?.FormulaLocal;                    
            if (formula1 != String.Empty) 
            {
                textBoxes[rowsArray, 0].Text = VprFormulaReplace(formula1, firstRow);
            }
            // Read Cells "D_" if value not empty then continue our work
            string formula2 = worksheet.Cells[firstRow, 4]?.FormulaLocal;
            if (formula2 != String.Empty)
            {
                textBoxes[rowsArray, 1].Text = VprFormulaReplace(formula2, firstRow);
            }
            // Read Cells "G_" if value not empty then continue our work
            string formula3 = worksheet.Cells[firstRow, 7]?.FormulaLocal;
            if (formula3 != String.Empty)
            {
                textBoxes[rowsArray, 2].Text = VprFormulaReplace(formula3, firstRow);
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
        /// Write IEK settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            int line = (int)RowsToArray.IekLine;

            TextBox[,] textBoxes = ReturnTextBoxArray();
            Label[] labels = ReturnLabelArray();

            DateTime localDate = DateTime.Now;
            dataInXml.WriteXml("IEK", textBoxes[line, 0].Text ?? String.Empty,
                                            textBoxes[line, 1].Text ?? String.Empty,
                                            textBoxes[line, 2].Text ?? String.Empty,
                                            textBoxes[line, 3].Text ?? String.Empty,
                                            localDate.ToString(new CultureInfo("ru-RU")));    
            labels[line].Text = localDate.ToString(new CultureInfo("ru-RU"));
        }

        /// <summary>
        /// Write EKF settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            int line = (int)RowsToArray.EkfLine;

            TextBox[,] textBoxes = ReturnTextBoxArray();
            Label[] labels = ReturnLabelArray();

            DateTime localDate = DateTime.Now;
            dataInXml.WriteXml("EKF", textBoxes[line, 0].Text ?? String.Empty,
                                            textBoxes[line, 1].Text ?? String.Empty,
                                            textBoxes[line, 2].Text ?? String.Empty,
                                            textBoxes[line, 3].Text ?? String.Empty,
                                            localDate.ToString(new CultureInfo("ru-RU")));
            labels[line].Text = localDate.ToString(new CultureInfo("ru-RU"));
        }

        /// <summary>
        /// Write DKC settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            int line = (int)RowsToArray.DkcLine;

            TextBox[,] textBoxes = ReturnTextBoxArray();
            Label[] labels = ReturnLabelArray();

            DateTime localDate = DateTime.Now;
            dataInXml.WriteXml("DKC", textBoxes[line, 0].Text ?? String.Empty,
                                            textBoxes[line, 1].Text ?? String.Empty,
                                            textBoxes[line, 2].Text ?? String.Empty,
                                            textBoxes[line, 3].Text ?? String.Empty,
                                            localDate.ToString(new CultureInfo("ru-RU")));
            labels[line].Text = localDate.ToString(new CultureInfo("ru-RU"));
        }

        /// <summary>
        /// Write KEAZ settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            int line = (int)RowsToArray.KeazLine;

            TextBox[,] textBoxes = ReturnTextBoxArray();
            Label[] labels = ReturnLabelArray();

            DateTime localDate = DateTime.Now;
            dataInXml.WriteXml("KEAZ", textBoxes[line, 0].Text ?? String.Empty,
                                            textBoxes[line, 1].Text ?? String.Empty,
                                            textBoxes[line, 2].Text ?? String.Empty,
                                            textBoxes[line, 3].Text ?? String.Empty,
                                            localDate.ToString(new CultureInfo("ru-RU")));
            labels[line].Text = localDate.ToString(new CultureInfo("ru-RU"));
        }

        /// <summary>
        /// Write DEKraft settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button10_Click(object sender, EventArgs e)
        {
            int line = (int)RowsToArray.DekraftLine;

            TextBox[,] textBoxes = ReturnTextBoxArray();
            Label[] labels = ReturnLabelArray();

            DateTime localDate = DateTime.Now;
            dataInXml.WriteXml("DEKraft", textBoxes[line, 0].Text ?? String.Empty,
                                            textBoxes[line, 1].Text ?? String.Empty,
                                            textBoxes[line, 2].Text ?? String.Empty,
                                            textBoxes[line, 3].Text ?? String.Empty,
                                            localDate.ToString(new CultureInfo("ru-RU")));
            labels[line].Text = localDate.ToString(new CultureInfo("ru-RU"));
        }

        /// <summary>
        /// Write TDM settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button12_Click(object sender, EventArgs e)
        {
            int line = (int)RowsToArray.TdmLine;

            TextBox[,] textBoxes = ReturnTextBoxArray();
            Label[] labels = ReturnLabelArray();

            DateTime localDate = DateTime.Now;
            dataInXml.WriteXml("TDM", textBoxes[line, 0].Text ?? String.Empty,
                                            textBoxes[line, 1].Text ?? String.Empty,
                                            textBoxes[line, 2].Text ?? String.Empty,
                                            textBoxes[line, 3].Text ?? String.Empty,
                                            localDate.ToString(new CultureInfo("ru-RU")));
            labels[line].Text = localDate.ToString(new CultureInfo("ru-RU"));
        }

        /// <summary>
        /// Write ABB settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button14_Click(object sender, EventArgs e)
        {
            int line = (int)RowsToArray.AbbLine;

            TextBox[,] textBoxes = ReturnTextBoxArray();
            Label[] labels = ReturnLabelArray();

            DateTime localDate = DateTime.Now;
            dataInXml.WriteXml("ABB", textBoxes[line, 0].Text ?? String.Empty,
                                            textBoxes[line, 1].Text ?? String.Empty,
                                            textBoxes[line, 2].Text ?? String.Empty,
                                            textBoxes[line, 3].Text ?? String.Empty,
                                            localDate.ToString(new CultureInfo("ru-RU")));
            labels[line].Text = localDate.ToString(new CultureInfo("ru-RU"));
        }

        /// <summary>
        /// Write Schneider settings to xml
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button16_Click(object sender, EventArgs e)
        {
            int line = (int)RowsToArray.SchneiderLine;

            TextBox[,] textBoxes = ReturnTextBoxArray();
            Label[] labels = ReturnLabelArray();

            DateTime localDate = DateTime.Now;
            dataInXml.WriteXml("Schneider", textBoxes[line, 0].Text ?? String.Empty,
                                            textBoxes[line, 1].Text ?? String.Empty,
                                            textBoxes[line, 2].Text ?? String.Empty,
                                            textBoxes[line, 3].Text ?? String.Empty,
                                            localDate.ToString(new CultureInfo("ru-RU")));
            labels[line].Text = localDate.ToString(new CultureInfo("ru-RU"));
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
