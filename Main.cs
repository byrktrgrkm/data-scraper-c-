using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace BirimFiyatlar
{
    public partial class Main : Form
    {

        private VeriToplayici veriToplayici;

        public Main()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }


        private void Main_Load(object sender, EventArgs e)
        {
            veriToplayici = new VeriToplayici(this);
            btnStop.Visible = false;
            btnRefresh.Visible = false;
            chckTumu.Checked = true;

            var list = new List<KeyValuePair<string, string>>() {
                new KeyValuePair<string, string>("çsb", "Çevre ve Şehircilik Bakanlığı"),
                new KeyValuePair<string, string>("dsi", "Devlet Su İşleri Genel Müdürlüğü"),
                new KeyValuePair<string, string>("dlh", "Altyapı Yatırımları (DLH)"),
                new KeyValuePair<string, string>("ilbank", "İller Bankası"),
                new KeyValuePair<string, string>("kgm", "Karayolları Genel Müdürlüğü"),
                new KeyValuePair<string, string>("kültür", "Kültür Bakanlığı"),
                new KeyValuePair<string, string>("msb", "Milli Savunma Bakanlığı"),
                new KeyValuePair<string, string>("orman", "Orman ve Su İşleri Bakanlığı"),
                new KeyValuePair<string, string>("ptt", "PTT"),
                new KeyValuePair<string, string>("tedas", "TEDAŞ [Elektrik Proje Tesis]"),
                new KeyValuePair<string, string>("vakiflar", "Vakıflar Genel Müdürlüğü"),
                new KeyValuePair<string, string>("tcdd", "TCDD Genel Müdürlüğü"),
            };


            comboKitapsec.DataSource = list.ToList();
            comboKitapsec.DisplayMember = "Value";
            comboKitapsec.ValueMember = "Key";

            comboKitapsec.SelectedIndex = 0;

        }
        private void startAction()
        {
            btnVeriTopla.Visible = true;
            btnVeriTopla.Enabled = false;
            btnStop.Visible = true;
            btnRefresh.Visible = false;

            if (chckTumu.Checked)
            {
                veriToplayici.setLimit(0);
            }
            else
            {
                veriToplayici.setLimit((int)numLimit.Value);
            }
            veriToplayici.setBook(comboKitapsec.SelectedValue.ToString());
            veriToplayici.setPage((int)numSayfa.Value);
            veriToplayici.setPageSize((int)numhersayfa.Value);
            veriToplayici.Start();
        }
        private void btnVeriTopla_Click(object sender, EventArgs e)
        {
            this.startAction();
        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            this.Stoping();
            veriToplayici.Stop();


        }
        public void Stoping()
        {
            btnVeriTopla.Text = "Durduruluyor..";
            btnStop.Visible = false;
        }
        public void Stoped()
        {
            btnVeriTopla.Text = "Devam Et";
            btnVeriTopla.Enabled = true;
            btnRefresh.Visible = true;
        }

        public void setInfo(string text)
        {
            lblInfo.Text = "Bilgi : " + text;
        }

        public void setTotalLink(int value)
        {
            lblToplamLink.Text = "Toplam Link : " + value.ToString();
        }
        public void setSelectedLink(int value)
        {
            lblToplananLink.Text = "Toplanan Link : " + value.ToString();
        }
        public void setOrderNo(int val)
        {
            lblSiraNo.Text = "İşlem Yapılan Sıra No: " + val.ToString();
        }
        public void loadingBar(int value)
        {
            progressBar1.Value = value;
        }

        public void popup(string text)
        {
            MessageBox.Show(text);
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            DialogResult show = MessageBox.Show("Tablo sıfırlanacaktır! Onaylıyor musunuz?", "Dikkat", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if(show == DialogResult.Yes)
            {
                dataGridView1.Rows.Clear();
                veriToplayici.refresh();

                this.startAction();
            }

          
        }

        public void Completed()
        {
            btnVeriTopla.Visible = false;
            btnVeriTopla.Text = "Veri Topla";
            btnRefresh.Visible = true;
            btnRefresh.Enabled = true;
            btnStop.Visible = false;
            this.loadingBar(0);

            this.setInfo("Tamamlandı.");
            
        }

        public void addTable(string pozno, string tanim, string birim, string kurum, string fasikul, List<PriceItem> data)
        {
            try
            {
                var index = dataGridView1.Rows.Add();
                DataGridViewRow row = dataGridView1.Rows[index];

                row.Cells["pozno"].Value = pozno;
                row.Cells["tanimi"].Value = tanim;
                row.Cells["birimi"].Value = birim;
                row.Cells["kurum"].Value = kurum;
                row.Cells["fasikul"].Value = fasikul;

                if (data.Count > 0)
                {
                    data.ForEach(item =>
                    {
                        if (row.Cells["d" + item.Year] != null)
                        {
                            row.Cells["d" + item.Year].Value = item.UnitPrice;
                        }

                    });
                }
            }
            catch(Exception e)
            {
                setInfo($"{pozno} nolu sırada sorun oluştu. {e.Message}");
            }
            
        }

        private  void btnExcelAktar_Click(object sender, EventArgs e)
        {
            btnExcelAktar.Enabled = false;
            btnVeriTopla.Enabled = false;
            btnRefresh.Enabled = false;

            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Excel Files | *.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {

                Thread thread = new Thread(t =>
                {
                    bool islem_zamani_hesaplama = false;
                    double ortalama_islem_zamani = 0;
                    DateTime d = DateTime.Now;

                    string execPath =
                        Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
                    Int32 unixTimestamp = (int)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalSeconds;
                    string path = execPath + "//output_" + unixTimestamp + ".xlsx";

                    Excel.Application excel = new Excel.Application();
                


                    excel.Visible = false;
                    object Missing = Type.Missing;
                    Workbook workbook = excel.Workbooks.Add(Missing);
                    Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
               

                    int StartCol = 1;
                    int StartRow = 1;
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                        myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                    }
                    StartRow++;
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if(!islem_zamani_hesaplama)
                        {
                            d = DateTime.Now;
                        }

                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {

                            Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                            myRange.Select();

                        }
                        btnExcelAktar.Text = "Aktarılıyor" + new String('.', (i % 3 + 1));

                        int val = (int)( ((i + 1) / Convert.ToDouble( dataGridView1.Rows.Count)) * 100);

                        loadingBar(val);

                        if (!islem_zamani_hesaplama)
                        {
                            ortalama_islem_zamani = (DateTime.Now - d).TotalMilliseconds;
                            islem_zamani_hesaplama = true;
                        }

                        double ms = (ortalama_islem_zamani * (dataGridView1.Rows.Count - i) / 1000);
                        

                        if(ms > 60)
                        {
                            int dakika = ((int)(ms / 60));
                            string saniye = ( ms -  dakika * 60).ToString("0.00");
                            setInfo($"yaklaşık {dakika.ToString()} dakika {saniye} saniye kaldı.");
                        }
                        else
                        {
                            string text = ms.ToString("0.00");
                            setInfo($"yaklaşık {text} saniye kaldı.");
                        }

                        
                    }


                    workbook.SaveAs(path);
                    workbook.Close();

                    File.Copy(path.Replace("file:\\", ""), dialog.FileName);

                    btnVeriTopla.Enabled = true;
                    btnExcelAktar.Enabled = true;
                    btnRefresh.Enabled = true;
                    btnExcelAktar.Text = "Excel Aktar";
                    this.setInfo("Aktarım tamamlandı");

                })
                { IsBackground = true };


                thread.Start();
            }
            else
            {
                btnVeriTopla.Enabled = true;
                btnExcelAktar.Enabled = true;
                btnRefresh.Enabled = true;

            }
    

        }

        private void chckTumu_CheckedChanged(object sender, EventArgs e)
        {

            lblLimit.Visible = !chckTumu.Checked;
            numLimit.Visible = !chckTumu.Checked;
        }
    }
}
