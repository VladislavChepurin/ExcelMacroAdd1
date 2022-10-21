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

        private readonly IResourcesForm4 resources;
        public Form4(IResourcesForm4 resources, IDataInXml dataInXml)
        {
            this.resources = resources;
            InitializeComponent();
            this.dataInXml = dataInXml;
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            comboBox1.Items.AddRange(resources.TransformerCurrent);
            comboBox1.SelectedIndex = StartTransformerCurrent;
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
                using (var db = new AppContext()) //Перенести в /AccessLayer/AccessData
                {
                    comboBox2.Items.Clear();
                    // ReSharper disable once CoVariantArrayConversion
                    comboBox2.Items.AddRange(db.Transformers
                        .Where(p => p.Current == comboBox1.SelectedItem.ToString())
                        .Select(p => p.Bus)
                        .ToHashSet()
                        .ToArray());
                    comboBox2.SelectedIndex = 0;
                }
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
                using (var db = new AppContext()) //Перенести в /AccessLayer/AccessData
                {
                    // ReSharper disable once CoVariantArrayConversion
                    comboBox3.Items.AddRange(db.Transformers
                        .Where(p => p.Current == comboBox1.SelectedItem.ToString() &&
                                    p.Bus == comboBox2.SelectedItem.ToString())
                        .Select(p => p.Accuracy)
                        .ToHashSet()
                        .ToArray());
                }
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
                using (var db = new AppContext()) //Перенести в /AccessLayer/AccessData
                {
                    // ReSharper disable once CoVariantArrayConversion
                    comboBox4.Items.AddRange(db.Transformers
                        .Where(p => p.Current == comboBox1.SelectedItem.ToString() &&
                                    p.Bus == comboBox2.SelectedItem.ToString() &&
                                    p.Accuracy == comboBox3.SelectedItem.ToString())
                        .Select(p => p.Power)
                        .ToHashSet()
                        .ToArray());
                }
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

        private void button1_Click(object sender, EventArgs e)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, label11.Text);
            var thread = new Thread(() => Clipboard.SetDataObject(data, true));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, label12.Text);
            var thread = new Thread(() => Clipboard.SetDataObject(data, true));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, label13.Text);
            var thread = new Thread(() => Clipboard.SetDataObject(data, true));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, label14.Text);
            var thread = new Thread(() => Clipboard.SetDataObject(data, true));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, label15.Text);
            var thread = new Thread(() => Clipboard.SetDataObject(data, true));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, label16.Text);
            var thread = new Thread(() => Clipboard.SetDataObject(data, true));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var writeExcel = new WriteExcel(dataInXml, "Iek", 0, label11.Text);
            writeExcel.Start();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var writeExcel = new WriteExcel(dataInXml, "Ekf", 0, label12.Text);
            writeExcel.Start();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var writeExcel = new WriteExcel(dataInXml, "Keaz", 0, label13.Text);
            writeExcel.Start();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var writeExcel = new WriteExcel(dataInXml, "Tdm", 0, label14.Text);
            writeExcel.Start();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            var writeExcel = new WriteExcel(dataInXml, "Iek", 0, label15.Text);
            writeExcel.Start();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            var writeExcel = new WriteExcel(dataInXml, "Dekraft", 0, label16.Text);
            writeExcel.Start();
        }
    }
}
