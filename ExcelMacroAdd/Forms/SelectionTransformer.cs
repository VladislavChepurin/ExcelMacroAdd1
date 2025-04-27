using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.Functions;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.Services.Interfaces;
using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMacroAdd.Forms
{
    public partial class SelectionTransformer : Form
    {
        private const byte StartTransformerCurrent = 0; // Начальный ток трансформации
        private readonly IDataInXml dataInXml;
        private readonly ISelectionTransformerData accessData;
        static readonly Mutex Mutex = new Mutex(false, "MutexSelectionTransformer_SingleInstance");
        private bool _mutexAcquired = false;
        private readonly string[] _transformerRatios = {"5/5", "10/5", "15/5", "20/5", "25/5", "30/5", "40/5", "50/5", "60/5", "75/5", "80/5", "100/5", "120/5", "125/5", "150/5", "200/5", "250/5", "300/5",
                "400/5", "500/5", "600/5", "750/5", "800/5", "1000/5", "1200/5", "1250/5", "1500/5", "1600/5", "2000/5", "2250/5", "2500/5", "3000/5", "4000/5", "5000/5" };

        private void SelectionTransformer_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_mutexAcquired)
            {
                Mutex.ReleaseMutex();
                _mutexAcquired = false;
            }
        }            

        public SelectionTransformer(IDataInXml dataInXml, ISelectionTransformerData accessData, IFormSettings formSettings)
        {
            InitializeComponent();
            try
            {
                _mutexAcquired = Mutex.WaitOne(TimeSpan.FromSeconds(1), false);
                if (!_mutexAcquired)
                {
                    Close();
                }
            }
            catch (AbandonedMutexException)
            {
                _mutexAcquired = true; // Мьютекс был оставлен, но теперь принадлежит текущему потоку
            }
            TopMost = formSettings.FormTopMost;
            this.dataInXml = dataInXml;
            this.accessData = accessData;            
        }

        private void SelectionTransformer_Load(object sender, EventArgs e)
        {
            cmbTransformerRatio.Items.AddRange(_transformerRatios);
            cmbTransformerRatio.SelectedIndex = StartTransformerCurrent;
            InitializeButtons();
        }

        private void InitializeButtons()
        {
            var buttonMap = new[]
            {
                new { ButtonCopy = btnCopyIekTti, ButtonWrite = btnWriteIekTti, Label = lblIekTti, Vendor = "IEK" },
                new { ButtonCopy = btnCopyEkfTte, ButtonWrite = btnWriteEkfTte, Label = lblEkfTte, Vendor = "EKF" },
                new { ButtonCopy = btnCopyKeazTtk, ButtonWrite = btnWriteKeazTtk, Label = lblKeazTtk, Vendor = "KEAZ" },
                new { ButtonCopy = btnCopyTdmTtn, ButtonWrite = btnWriteTdmTtn, Label = lblTdmTtn, Vendor = "TDM" },
                new { ButtonCopy = btnCopyIekTop, ButtonWrite = btnWriteIekTop, Label = lblIekTop, Vendor = "IEK" },
                new { ButtonCopy = btnCopyDekTop, ButtonWrite = btnWriteDekTop, Label = lblDekTop, Vendor = "DEKraft" }
            };

            foreach (var item in buttonMap)
            {
                item.ButtonCopy.Click += (s, e) => CopyToClipboard(item.Label.Text);
                item.ButtonWrite.Click += (s, e) => WriteToExcel(item.Vendor, item.Label.Text);
            }
        }

        private void WriteToExcel(string vendor, string article)
        {
            if (string.IsNullOrEmpty(article)) return;
            var writeExcel = new WriteExcel(dataInXml, vendor, 0, article);
            writeExcel.Start();
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

        private void cmbTransformerRatio_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboboxBusRefresh();
            ComboboxAccuracyRefresh();
            ComboboxPowerRefresh();
        }

        private void cmbTransformerExecution_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboboxAccuracyRefresh();
            ComboboxPowerRefresh();
        }

        private void cmbTransformerAccuracy_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboboxPowerRefresh();
        }

        private void ComboboxBusRefresh()
        {
            try
            {
                cmbTransformerExecution.Items.Clear();             
                cmbTransformerExecution.Items.AddRange(accessData.AccessTransformer.GetComboBox2Items(cmbTransformerRatio.SelectedItem.ToString()));
                cmbTransformerExecution.SelectedIndex = 0;
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
                cmbTransformerAccuracy.Items.Clear();
                cmbTransformerAccuracy.Items.AddRange(accessData.AccessTransformer.GetComboBox3Items(cmbTransformerRatio.SelectedItem.ToString(), cmbTransformerExecution.SelectedItem.ToString()));
                cmbTransformerAccuracy.SelectedIndex = 0;
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
                cmbTransformerPower.Items.Clear();                
                cmbTransformerPower.Items.AddRange(accessData.AccessTransformer.GetComboBox4Items(cmbTransformerRatio.SelectedItem.ToString(), cmbTransformerExecution.SelectedItem.ToString(), cmbTransformerAccuracy.SelectedItem.ToString()));
                cmbTransformerPower.SelectedIndex = 0;
            }
            catch (NotSupportedException)
            {
                MessageBox.Show(Properties.Resources.invalidOperation);
            }
        }

        private void cmbTransformerPower_SelectedIndexChanged(object sender, EventArgs e)
        {
            var transformerRow =
              accessData.AccessTransformer.GetArticle(
                  cmbTransformerRatio.SelectedItem.ToString(),
                  cmbTransformerExecution.SelectedItem.ToString(),
                  cmbTransformerAccuracy.SelectedItem.ToString(),
                  cmbTransformerPower.SelectedItem.ToString());

            if (string.IsNullOrEmpty(transformerRow.IekTti))
            {
                btnCopyIekTti.Enabled = false;
                btnWriteIekTti.Enabled = false;
                lblIekTti.Text = Properties.Resources.absent;
            }
            else
            {
                btnCopyIekTti.Enabled = true;
                btnWriteIekTti.Enabled = true;
                lblIekTti.Text = transformerRow.IekTti;
            }

            if (string.IsNullOrEmpty(transformerRow.EkfTte))
            {
                btnCopyEkfTte.Enabled = false;
                btnWriteEkfTte.Enabled = false;
                lblEkfTte.Text = Properties.Resources.absent;
            }
            else
            {
                btnCopyEkfTte.Enabled = true;
                btnWriteEkfTte.Enabled = true;
                lblEkfTte.Text = transformerRow.EkfTte;
            }

            if (string.IsNullOrEmpty(transformerRow.KeazTtk))
            {
                btnCopyKeazTtk.Enabled = false;
                btnWriteKeazTtk.Enabled = false;
                lblKeazTtk.Text = Properties.Resources.absent;
            }
            else
            {
                btnCopyKeazTtk.Enabled = true;
                btnWriteKeazTtk.Enabled = true;
                lblKeazTtk.Text = transformerRow.KeazTtk;
            }

            if (string.IsNullOrEmpty(transformerRow.TdmTtn))
            {
                btnCopyTdmTtn.Enabled = false;
                btnWriteTdmTtn.Enabled = false;
                lblTdmTtn.Text = Properties.Resources.absent;
            }
            else
            {
                btnCopyTdmTtn.Enabled = true;
                btnWriteTdmTtn.Enabled = true;
                lblTdmTtn.Text = transformerRow.TdmTtn;
            }

            if (string.IsNullOrEmpty(transformerRow.IekTop))
            {
                btnCopyIekTop.Enabled = false;
                btnWriteIekTop.Enabled = false;
                lblIekTop.Text = Properties.Resources.absent;
            }
            else
            {
                btnCopyIekTop.Enabled = true;
                btnWriteIekTop.Enabled = true;
                lblIekTop.Text = transformerRow.IekTop;
            }

            if (string.IsNullOrEmpty(transformerRow.DekTop))
            {
                btnCopyDekTop.Enabled = false;
                btnWriteDekTop.Enabled = false;
                lblDekTop.Text = Properties.Resources.absent;
            }
            else
            {
                btnCopyDekTop.Enabled = true;
                btnWriteDekTop.Enabled = true;
                lblDekTop.Text = transformerRow.DekTop;
            }
            //Обновление картинки из базы
            pictureBox1.Image = ByteArrayToImage(accessData.AccessTransformer.GetBlobPictureDb(
                cmbTransformerRatio.SelectedItem.ToString(),
                cmbTransformerExecution.SelectedItem.ToString(),
                cmbTransformerAccuracy.SelectedItem.ToString(),
                cmbTransformerPower.SelectedItem.ToString()));
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
            if (byteArrayIn == null || byteArrayIn.Length == 0)
                return Properties.Resources.DefaultImage;

            using (var ms = new MemoryStream(byteArrayIn))
            {
                return Image.FromStream(ms);
            }
        }       
    }
}
