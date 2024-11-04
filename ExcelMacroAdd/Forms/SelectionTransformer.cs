using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class SelectionTransformer : Form
    {
        private const byte StartTransformerCurrent = 0; // Начальный ток трансформации
        private readonly IDataInXml dataInXml;
        private readonly ISelectionTransformerData accessData;

        //Singelton
        private static SelectionTransformer instance;
        public static async Task getInstance(IDataInXml dataInXml, ISelectionTransformerData accessData, IFormSettings formSettings)
        {
            if (instance == null)
            {
                await Task.Run(() =>
                {
                    instance = new SelectionTransformer(dataInXml, accessData)
                    {
                        TopMost = formSettings.FormTopMost
                    };
                    instance.ShowDialog();
                });
            }
        }

        private void SelectionTransformer_FormClosed(object sender, FormClosedEventArgs e) =>
            instance = null;

        private SelectionTransformer(IDataInXml dataInXml, ISelectionTransformerData accessData)
        {
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            InitializeComponent();
        }

        private void SelectionTransformer_Load(object sender, EventArgs e)
        {
            comboBox1.Items.AddRange(new object[] {"5/5", "10/5", "15/5", "20/5", "25/5", "30/5", "40/5", "50/5", "60/5", "75/5", "80/5", "100/5", "120/5", "125/5", "150/5", "200/5", "250/5", "300/5",
                "400/5", "500/5", "600/5", "750/5", "800/5", "1000/5", "1200/5", "1250/5", "1500/5", "1600/5", "2000/5", "2250/5", "2500/5", "3000/5", "4000/5", "5000/5" });
            comboBox1.SelectedIndex = StartTransformerCurrent;

            button1.Click += (s, a) =>
            {
                CopyToClipboard(label11.Text);
            };

            button2.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "IEK", 0, label11.Text);
                writeExcel.Start();
            };

            button3.Click += (s, a) =>
            {
                CopyToClipboard(label12.Text);
            };

            button4.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "EKF", 0, label12.Text);
                writeExcel.Start();
            };

            button5.Click += (s, a) =>
            {
                CopyToClipboard(label13.Text);
            };

            button6.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "KEAZ", 0, label13.Text);
                writeExcel.Start();
            };

            button7.Click += (s, a) =>
            {
                CopyToClipboard(label14.Text);
            };

            button8.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "TDM", 0, label14.Text);
                writeExcel.Start();
            };

            button9.Click += (s, a) =>
            {
                CopyToClipboard(label15.Text);
            };

            button10.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "IEK", 0, label15.Text);
                writeExcel.Start();
            };

            button11.Click += (s, a) =>
            {
                CopyToClipboard(label16.Text);
            };

            button12.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "DEKraft", 0, label16.Text);
                writeExcel.Start();
            };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Pen pen = new Pen(Color.FromArgb(40, 0, 0, 100));
            e.Graphics.DrawLine(pen, 268, 102, 638, 102);
            e.Graphics.DrawLine(pen, 268, 143, 638, 143);
            e.Graphics.DrawLine(pen, 268, 184, 638, 184);
            e.Graphics.DrawLine(pen, 268, 225, 638, 225);
            e.Graphics.DrawLine(pen, 268, 266, 638, 266);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboboxBusRefresh();
            ComboboxAccuracyRefresh();
            ComboboxPowerRefresh();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboboxAccuracyRefresh();
            ComboboxPowerRefresh();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboboxPowerRefresh();
        }

        private void ComboboxBusRefresh()
        {
            try
            {
                comboBox2.Items.Clear();
                // ReSharper disable once CoVariantArrayConversion
                comboBox2.Items.AddRange(accessData.AccessTransformer.GetComboBox2Items(comboBox1.SelectedItem.ToString()));
                comboBox2.SelectedIndex = 0;
            }
            catch (NotSupportedException)
            {
                MessageBox.Show(Properties.Resources.invalidOperation);
            }
        }

        private void ComboboxAccuracyRefresh()
        {
            try
            {
                comboBox3.Items.Clear();
                // ReSharper disable once CoVariantArrayConversion
                comboBox3.Items.AddRange(accessData.AccessTransformer.GetComboBox3Items(comboBox1.SelectedItem.ToString(), comboBox2.SelectedItem.ToString()));
                comboBox3.SelectedIndex = 0;
            }
            catch (NotSupportedException)
            {
                MessageBox.Show(Properties.Resources.invalidOperation);
            }
        }

        private void ComboboxPowerRefresh()
        {
            try
            {
                comboBox4.Items.Clear();
                // ReSharper disable once CoVariantArrayConversion
                comboBox4.Items.AddRange(accessData.AccessTransformer.GetComboBox4Items(comboBox1.SelectedItem.ToString(), comboBox2.SelectedItem.ToString(), comboBox3.SelectedItem.ToString()));
                comboBox4.SelectedIndex = 0;
            }
            catch (NotSupportedException)
            {
                MessageBox.Show(Properties.Resources.invalidOperation);
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            var transformerRow =
               accessData.AccessTransformer.GetArticle(
                   comboBox1.SelectedItem.ToString(),
                   comboBox2.SelectedItem.ToString(),
                   comboBox3.SelectedItem.ToString(),
                   comboBox4.SelectedItem.ToString());

            if (string.IsNullOrEmpty(transformerRow.IekTti))
            {
                button1.Enabled = false;
                button2.Enabled = false;
                label11.Text = Properties.Resources.absent;
            }
            else
            {
                button1.Enabled = true;
                button2.Enabled = true;
                label11.Text = transformerRow.IekTti;
            }

            if (string.IsNullOrEmpty(transformerRow.EkfTte))
            {
                button3.Enabled = false;
                button4.Enabled = false;
                label12.Text = Properties.Resources.absent;
            }
            else
            {
                button3.Enabled = true;
                button4.Enabled = true;
                label12.Text = transformerRow.EkfTte;
            }

            if (string.IsNullOrEmpty(transformerRow.KeazTtk))
            {
                button5.Enabled = false;
                button6.Enabled = false;
                label13.Text = Properties.Resources.absent;
            }
            else
            {
                button5.Enabled = true;
                button6.Enabled = true;
                label13.Text = transformerRow.KeazTtk;
            }

            if (string.IsNullOrEmpty(transformerRow.TdmTtn))
            {
                button7.Enabled = false;
                button8.Enabled = false;
                label14.Text = Properties.Resources.absent;
            }
            else
            {
                button7.Enabled = true;
                button8.Enabled = true;
                label14.Text = transformerRow.TdmTtn;
            }

            if (string.IsNullOrEmpty(transformerRow.IekTop))
            {
                button9.Enabled = false;
                button10.Enabled = false;
                label15.Text = Properties.Resources.absent;
            }
            else
            {
                button9.Enabled = true;
                button10.Enabled = true;
                label15.Text = transformerRow.IekTop;
            }

            if (string.IsNullOrEmpty(transformerRow.DekTop))
            {
                button11.Enabled = false;
                button12.Enabled = false;
                label16.Text = Properties.Resources.absent;
            }
            else
            {
                button11.Enabled = true;
                button12.Enabled = true;
                label16.Text = transformerRow.DekTop;
            }
            //Обновление картинки из базы
            pictureBox1.Image = ByteArrayToImage(accessData.AccessTransformer.GetBlobPictureDb(
                comboBox1.SelectedItem.ToString(),
                comboBox2.SelectedItem.ToString(),
                comboBox3.SelectedItem.ToString(),
                comboBox4.SelectedItem.ToString()));
        }

        private void CopyToClipboard(string text)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, text);
            var thread = new Thread(() => Clipboard.SetDataObject(data, true));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }
        private Image ByteArrayToImage(byte[] byteArrayIn)
        {
            using (var ms = new MemoryStream(byteArrayIn))
            {
                var returnImage = Image.FromStream(ms);
                return returnImage;
            }
        }
    }
}
