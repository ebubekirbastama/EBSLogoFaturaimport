using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using EBSFaturaimport;

namespace XmlOkuma
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }
        OpenFileDialog op; Thread th;
        ebubekirbastama eckyazilim = new ebubekirbastama();
        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 12;
            dataGridView1.Columns[0 ].Name = "Sıra No";
            dataGridView1.Columns[1 ].Name = "Barkod";
            dataGridView1.Columns[2 ].Name = "Mal/Hizmet";
            dataGridView1.Columns[3 ].Name = "Miktar";
            dataGridView1.Columns[4 ].Name = "Birim";
            dataGridView1.Columns[5 ].Name = "Birim Fiyat";
            dataGridView1.Columns[6 ].Name = "İskonto Oranı";
            dataGridView1.Columns[7 ].Name = "İskonto Tutarı";
            dataGridView1.Columns[8 ].Name = "KDV Oranı";
            dataGridView1.Columns[9 ].Name = "KDV Tutarı";
            dataGridView1.Columns[10].Name = "Mal Hizmet Tutarı";
            dataGridView1.Columns[11].Name = "Firma Adı";
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            op = new OpenFileDialog();
            op.Multiselect = true;
            if (op.ShowDialog() == DialogResult.OK)
            {
                for (int i = 0; i < op.FileNames.Length; i++)
                {
                    eckyazilim.bsl(op.FileNames[i].ToString(), dataGridView1, "PROCEDUCER_CODE", "MASTER_DEF");//PROCEDUCER_CODE
                }
              
            }
        }
        async void pictureBox2_Click(object sender, EventArgs e)
        {
            await Excelaktar();
        }
        async Task Excelaktar()
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            excel.DisplayAlerts = false;
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
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();
                }
            }
            MessageBox.Show("Bütün Verilerin Hepsi Aktarıldı.");
        }
        private void label1_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.ebubekirbastama.com/2021/04/logo-e-fatura-xml-to-excel.html");
        }
    }
}
