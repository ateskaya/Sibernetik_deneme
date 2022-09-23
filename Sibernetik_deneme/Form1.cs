using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Data.OleDb;
using System.DirectoryServices;
namespace Sibernetik_deneme
{
    public partial class Form1 : Form
    {
        OleDbConnection baglanti2;
        OleDbConnection baglanti3;
        OleDbDataAdapter oda;
        OleDbDataAdapter oda2;
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        OpenFileDialog od = new OpenFileDialog();
        OpenFileDialog od2 = new OpenFileDialog();
        BackgroundWorker bw = new BackgroundWorker
        {
            WorkerReportsProgress = true,
            WorkerSupportsCancellation = true
        };
        BackgroundWorker bw2 = new BackgroundWorker
        {
            WorkerReportsProgress = true,
            WorkerSupportsCancellation = true
        };

        public Form1()
        {

            InitializeComponent();
        }
        private void InsertExcelRecords()
        {
            try
            {
                OleDbConnection Econ = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + od.FileName + ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'");
                string Query = string.Format("Select * FROM [Sayfa1$]");
                OleDbCommand Ecom = new OleDbCommand(Query, Econ);
                Econ.Open();
                DataSet ds = new DataSet();
                OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
                Econ.Close();
                oda.Fill(ds);
                DataTable Exceldt = ds.Tables[0];
                Exceldt.AcceptChanges();
                SqlConnection sqlConnection = new SqlConnection(@"server=(localdb)\MSSQLLocalDB;Initial Catalog=Sibernetik;Integrated Security=SSPI");
                SqlBulkCopy objbulk = new SqlBulkCopy(sqlConnection);
                objbulk.DestinationTableName = "[Toplam Fatura2]";
                objbulk.ColumnMappings.Add("Tarih", "Tarih");
                objbulk.ColumnMappings.Add("[Belge No]", "Belge_No");
                objbulk.ColumnMappings.Add("[Firma Kodu]", "Firma_Kodu");
                objbulk.ColumnMappings.Add("[Firma Ad�]", "Firma_Ad�");
                objbulk.ColumnMappings.Add("[Stok Kodu]", "Stok_Kodu");
                objbulk.ColumnMappings.Add("[Stok Ad�]", "Stok_Ad�");
                objbulk.ColumnMappings.Add("[Sat�r A��klama]", "Sat�r_A��klama");
                objbulk.ColumnMappings.Add("Barkod", "Barkod");
                objbulk.ColumnMappings.Add("[KDV Oran�]", "KDV_Oran�");
                objbulk.ColumnMappings.Add("[Miktar]", "Miktar");
                objbulk.ColumnMappings.Add("[Birim]", "Birim");
                objbulk.ColumnMappings.Add("[Birim Fiyat� (TL)]", "Birim_Fiyat�_TL");
                objbulk.ColumnMappings.Add("[KDV Hari� Toplam (TL)]", "KDV_Hari�_Toplam_TL");
                objbulk.ColumnMappings.Add("[KDV Tutar� (TL)]", "KDV_Tutar�_TL");
                objbulk.ColumnMappings.Add("[KDV Dahil Toplam (TL)]", "KDV_Dahil_Toplam_TL");
                objbulk.ColumnMappings.Add("[Hareket Kuru]", "Hareket_Kuru");
                objbulk.ColumnMappings.Add("[D�viz Birim Fiyat]", "D�viz_Birim_Fiyat");
                objbulk.ColumnMappings.Add("[D�viz Tutar]", "D�viz_Tutar");
                //objbulk.ColumnMappings.Add("Vergi_Dairesi", "Vergi_Dairesi");
                //objbulk.ColumnMappings.Add("Vergi_No_TC_Kimlik", "Vergi_No_TC_Kimlik");
                sqlConnection.Open();
                objbulk.WriteToServer(Exceldt);
                sqlConnection.Close();
                MessageBox.Show("Data has been Imported successfully.", "Imported", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Data has not been Imported due to :{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void InsertExcelRecords2()
        {
            try
            {
                OleDbConnection Econ1 = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + od2.FileName + ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'");
                string Query1 = string.Format("Select * FROM [Sayfa1$]");
                OleDbCommand Ecom1 = new OleDbCommand(Query1, Econ1);
                Econ1.Open();
                DataSet ds1 = new DataSet();
                OleDbDataAdapter oda1 = new OleDbDataAdapter(Query1, Econ1);
                Econ1.Close();
                oda1.Fill(ds1);
                DataTable Exceldt = ds1.Tables[0];
                Exceldt.AcceptChanges();
                SqlConnection sqlConnection = new SqlConnection(@"server=(localdb)\MSSQLLocalDB;Initial Catalog=Sibernetik;Integrated Security=SSPI");
                SqlBulkCopy objbulk = new SqlBulkCopy(sqlConnection);
                objbulk.DestinationTableName = "[Y�klenilen KDV Listesi2]";
                objbulk.ColumnMappings.Add("[S�ra No]", "[S�ra No]");
                objbulk.ColumnMappings.Add("[Al�� Faturas�n�n Tarihi]", "[Al�� Faturas�n�n Tarihi]");
                objbulk.ColumnMappings.Add("[Al�� Faturas�n�n Serisi]", "[Al�� Faturas�n�n Serisi]");
                objbulk.ColumnMappings.Add("[Al�� Faturas�n�n S�ra No'su]", "[Al�� Faturas�n�n S�ra No'su]");
                objbulk.ColumnMappings.Add("[Belge No]", "[Belge No]");
                objbulk.ColumnMappings.Add("[Sat�c�n�n Ad� - Soyad� / �nvan�]", "[Sat�c�n�n Ad� - Soyad� / �nvan�]");
                objbulk.ColumnMappings.Add("[Sat�c�n�n Vergi Kimlik Numaras� / TC Kimlik Numaras�]", "[Sat�c�n�n Vergi Kimlik Numaras� / TC Kimlik Numaras�]");
                objbulk.ColumnMappings.Add("[Al�nan Mal ve/veya Hizmetin Cinsi]", "[Al�nan Mal ve/veya Hizmetin Cinsi]");
                objbulk.ColumnMappings.Add("[Al�nan Mal ve/veya Hizmetin Miktar�]", "[Al�nan Mal ve/veya Hizmetin Miktar�]");
                objbulk.ColumnMappings.Add("[Al�nan Mal ve/veya Hizmetin KDV Hari� Tutar�]", "[Al�nan Mal ve/veya Hizmetin KDV Hari� Tutar�]");
                objbulk.ColumnMappings.Add("[KDV'si]", "[KDV'si]");
                objbulk.ColumnMappings.Add("[B�nyeye Giren Mal ve/veya Hizmetin KDV'si]", "[B�nyeye Giren Mal ve/veya Hizmetin KDV'si]");
                objbulk.ColumnMappings.Add("[GGB Tescil No'su (Al�� �thalat �se)]", "[GGB Tescil No'su (Al�� �thalat �se)]");
                objbulk.ColumnMappings.Add("[Belgeye �li�kin �ade Hakk� Do�uran ��lem T�r�]", "[Belgeye �li�kin �ade Hakk� Do�uran ��lem T�r�]");
                objbulk.ColumnMappings.Add("[Y�klenim T�r�]", "[Y�klenim T�r�]");
                objbulk.ColumnMappings.Add("[Belgenin �ndirime Konu Edildi�i KDV D�nemi]", "[Belgenin �ndirime Konu Edildi�i KDV D�nemi]");
                objbulk.ColumnMappings.Add("[Belgenin Y�klenildi�i KDV D�nemi]", "[Belgenin Y�klenildi�i KDV D�nemi]");
                objbulk.ColumnMappings.Add("[Arac�n Plaka Numaras�]", "[Arac�n Plaka Numaras�]");
                objbulk.ColumnMappings.Add("[Arac�n �asi Numaras�]", "[Arac�n �asi Numaras�]");
                sqlConnection.Open();
                objbulk.WriteToServer(Exceldt);
                sqlConnection.Close();
                MessageBox.Show("Data has been Imported successfully.", "Imported", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Data has not been Imported due to :{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        void excelbaglan1()
        {
            baglanti2 = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + od2.FileName + ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'");
        }
        void excelbaglan2()
        {
            baglanti3 = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + od.FileName + ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'");
        }
        void veriTaban�naG�nder2()
        {
            excelbaglan1();
            baglanti2.Open();
            dt = new DataTable();
            oda2 = new OleDbDataAdapter("Select * from [Sayfa1$]", baglanti2);
            oda2.Fill(dt);
            dataGridView1.DataSource = dt;
            baglanti2.Close();

        }
        void veriTaban�naG�nder1()
        {
            excelbaglan2();
            baglanti3.Open();
            dt = new DataTable();
            oda = new OleDbDataAdapter("Select * from [Sayfa1$]", baglanti3);
            oda.Fill(dt2);
            dataGridView3.DataSource = dt2;
            baglanti3.Close();
        }
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
        }
        private void Form1_Load_1(object sender, EventArgs e)
        {
            excelbaglan1();
            excelbaglan2();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            od2.Filter = "Excell|*.xls;*.xlsx;";

            DialogResult dr2 = od2.ShowDialog();
            if (dr2 == DialogResult.Abort)
                return;
            if (dr2 == DialogResult.Cancel)
                return;
            tabload�.Text = od2.FileName.ToString();
            String text = tabload�.Text;
            veriTaban�naG�nder2();
            if (bw2.IsBusy)
            {
                return;
            }
            System.Diagnostics.Stopwatch sWatch = new System.Diagnostics.Stopwatch();
            bw2.DoWork += (bwSender, bwArg) =>
            {
                sWatch.Start();
                InsertExcelRecords2();
            };
            bw2.ProgressChanged += (bwSender, bwArg) =>
            {
            };
            bw2.RunWorkerCompleted += (bwSender, bwArg) =>
            {
                sWatch.Stop();
                tabload�.Text = "";
                button4.Enabled = true;
                bw2.Dispose();
            };
            button4.Enabled = false;
            button4.Visible = false;
            bw2.RunWorkerAsync();
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }
        private void tabload�2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
        }
        
        private void button5_Click(object sender, EventArgs e)
        {
            od.Filter = "Excell|*.xls;*.xlsx;";

            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
                return;
            if (dr == DialogResult.Cancel)
                return;
            tabload�2.Text = od.FileName.ToString();
            String text = tabload�2.Text;
            veriTaban�naG�nder1();
            if (bw.IsBusy)
            {
                return;
            }
            System.Diagnostics.Stopwatch sWatch = new System.Diagnostics.Stopwatch();
            bw.DoWork += (bwSender, bwArg) =>
            {
                sWatch.Start();
                InsertExcelRecords();
            };
            bw.ProgressChanged += (bwSender, bwArg) =>
            {
            };
            bw.RunWorkerCompleted += (bwSender, bwArg) =>
            {
                sWatch.Stop();
                tabload�2.Text = "";
                button5.Enabled = true;
                bw.Dispose();
            };
            button5.Enabled = false;
            button5.Visible = false; 
            bw.RunWorkerAsync();
        }
        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }
        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }
        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
        }
        private void label1_Click(object sender, EventArgs e)
        {
        }
    }
}