using ExcelMacroAdd.AccessLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Interfaces;
using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using AppContext = ExcelMacroAdd.DataLayer.Entity.AppContext;

namespace ExcelMacroAdd.Forms
{
    public partial class Form4 : Form
    {
        private const byte StartTransformerCurrent = 0; // Начальный ток трансформации
        private readonly IDataInXml dataInXml;
        private readonly IForm4Data accessData;

        private readonly IResourcesForm4 resources;
        public Form4(IResourcesForm4 resources, IDataInXml dataInXml, IForm4Data accessData)
        {
            this.resources = resources;
            this.dataInXml = dataInXml;
            this.accessData = accessData;
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            comboBox1.Items.AddRange(resources.TransformerCurrent);
            comboBox1.SelectedIndex = StartTransformerCurrent;

            button1.Click += (s, a) =>
            {
                CopyToClipboard(label11.Text);
            };

            button2.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "Iek", 0, label11.Text);
                writeExcel.Start();
            };

            button3.Click += (s, a) =>
            {
                CopyToClipboard(label12.Text);
            };

            button4.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "Ekf", 0, label12.Text);
                writeExcel.Start();
            };

            button5.Click += (s, a) =>
            {
                CopyToClipboard(label13.Text);
            };

            button6.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "Keaz", 0, label13.Text);
                writeExcel.Start();
            };

            button7.Click += (s, a) =>
            {
                CopyToClipboard(label14.Text);
            };

            button8.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "Tdm", 0, label14.Text);
                writeExcel.Start();
            };

            button9.Click += (s, a) =>
            {
                CopyToClipboard(label15.Text);
            };

            button10.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "Iek", 0, label15.Text);
                writeExcel.Start();
            };

            button11.Click += (s, a) =>
            {
                CopyToClipboard(label16.Text);
            };

            button12.Click += (s, a) =>
            {
                var writeExcel = new WriteExcel(dataInXml, "Dekraft", 0, label16.Text);
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
            switch (comboBox2.SelectedItem.ToString())
            {
                case "Встроенная":
                    pictureBox1.Image = Properties.Resources.ttk_a;
                    break;
                case "30мм":
                    pictureBox1.Image = Properties.Resources.ttk_30;
                    break;
                case "40мм":
                    pictureBox1.Image = Properties.Resources.ttk_40;
                    break;
                case "60мм":
                    pictureBox1.Image = Properties.Resources.ttk_60;
                    break;
                case "85мм":
                    pictureBox1.Image = Properties.Resources.ttk_85;
                    break;
                case "100мм":
                    pictureBox1.Image = Properties.Resources.ttk_100;
                    break;
                case "125мм":
                    pictureBox1.Image = Properties.Resources.ttk_125;
                    break;
            }
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
                comboBox2.Items.AddRange(accessData.GetComboBox2Items(comboBox1.SelectedItem.ToString()));
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
                comboBox3.Items.AddRange(accessData.GetComboBox3Items(comboBox1.SelectedItem.ToString(), comboBox2.SelectedItem.ToString()));
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
                comboBox4.Items.AddRange(accessData.GetComboBox4Items(comboBox1.SelectedItem.ToString(), comboBox2.SelectedItem.ToString(), comboBox3.SelectedItem.ToString()));
                comboBox4.SelectedIndex = 0;
            }
            catch (NotSupportedException)
            {
                MessageBox.Show(Properties.Resources.invalidOperation);
            }
        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            PutDataDbTransformer();
        }

        private void PutDataDbTransformer()
        {
            using (AppContext db = new AppContext())
            {
                var transformerRow = db.Transformers
                    .Where(t=> t.Current == comboBox1.SelectedItem.ToString() 
                               && t.Bus == comboBox2.SelectedItem.ToString()
                               && t.Accuracy == comboBox3.SelectedItem.ToString()
                               && t.Power == comboBox4.SelectedItem.ToString())
                    .Select(  t => new { IekTti = t.Iek, EkfTte = t.Ekf, KeazTtk = t.Keaz, TdmTtn = t.Tdm, IekTop = t.IekTopTpsh, DekTop = t.DekraftTopTpsh })
                    .FirstOrDefault();

                if (string.IsNullOrEmpty(transformerRow?.IekTti))
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
                if (string.IsNullOrEmpty(transformerRow?.EkfTte))
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
                if (string.IsNullOrEmpty(transformerRow?.KeazTtk))
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
                if (string.IsNullOrEmpty(transformerRow?.TdmTtn))
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
                if (string.IsNullOrEmpty(transformerRow?.IekTop))
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

                if (string.IsNullOrEmpty(transformerRow?.DekTop))
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
            }
        }

        private void CopyToClipboard(string text)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, text);
            var thread = new Thread(() => Clipboard.SetDataObject(data, true));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }
    }
}
