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
using ClosedXML.Excel;

namespace Sibernetik_deneme
{
    public partial class Form1 : Form
    {
        OleDbConnection baglanti2;
        OleDbConnection baglanti3;
        OleDbDataAdapter oda;
        OleDbDataAdapter oda2;
        SqlDataAdapter sda;
        SqlDataAdapter sda2;
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        DataTable dt4 = new DataTable();
        OpenFileDialog od = new OpenFileDialog();
        OpenFileDialog od2 = new OpenFileDialog();
        SqlConnection sqlConnection1;
        SqlConnection sqlConnection2;
        SqlConnection sqlConnection3;
        SqlConnection sqlConnection4;
        SqlCommand sqlCommand1;
        SqlCommand sqlCommand2;
        SqlCommand sqlCommand3;
        SqlCommand sqlCommand4;

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
                objbulk.DestinationTableName = "[Toplam_Fatura]";
                objbulk.ColumnMappings.Add("[Tarih]", "[Tarih]");
                objbulk.ColumnMappings.Add("[Belge No]", "[Belge_No]");
                objbulk.ColumnMappings.Add("[Firma Kodu]", "[Firma_Kodu]");
                objbulk.ColumnMappings.Add("[Firma Adý]", "[Firma_Adý]");
                objbulk.ColumnMappings.Add("[Vergi No / TC Kimlik No]", "[Vergi_No_TC_Kimlik_No]");
                objbulk.ColumnMappings.Add("[Stok Kodu]", "[Stok_Kodu]");
                objbulk.ColumnMappings.Add("[Stok Adý]", "[Stok_Adý]");
                objbulk.ColumnMappings.Add("[Depo Kodu]", "[Depo_Kodu]");
                objbulk.ColumnMappings.Add("[Satýr Açýklama]", "[Satýr_Açýklama]");
                objbulk.ColumnMappings.Add("[Barkod]", "[Barkod]");
                objbulk.ColumnMappings.Add("[KDV Oraný]", "[KDV_Oraný]");
                objbulk.ColumnMappings.Add("[Miktar]", "[Miktar]");
                objbulk.ColumnMappings.Add("[Birim]", "[Birim]");
                objbulk.ColumnMappings.Add("[Birim Fiyatý (TL)]", "[Birim_Fiyatý_TL]");
                objbulk.ColumnMappings.Add("[KDV Hariç Toplam (TL)]", "[KDV_Hariç_Toplam_TL]");
                objbulk.ColumnMappings.Add("[KDV Tutarý (TL)]", "[KDV_Tutarý_TL]");
                objbulk.ColumnMappings.Add("[KDV Dahil Toplam (TL)]", "[KDV_Dahil_Toplam_TL]");
                objbulk.ColumnMappings.Add("[Genel Toplam (TL)]", "[Genel_Toplam_TL]");
                objbulk.ColumnMappings.Add("[Hareket Kuru]", "[Hareket_Kuru]");
                objbulk.ColumnMappings.Add("[Döviz Birim Fiyat]", "[Döviz_Birim_Fiyat]");
                objbulk.ColumnMappings.Add("[Döviz Tutar]", "[Döviz_Tutar]");
                sqlConnection.Open();
                objbulk.WriteToServer(Exceldt);
                sqlConnection.Close();
                MessageBox.Show("Veri tabanýna eklendi.", "Bilgi Kutusu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Dosyanýn sayfa isminin 'Sayfa1' olduðundan ve seçtiðiniz dosyanýn içeriðinde Toplam Faturalarýn olduðundan emin olun!!!  "), "Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                MessageBox.Show("hata nedeni:" + ex);
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
                objbulk.DestinationTableName = "[Yüklenilen KDV Listesi2]";
                objbulk.ColumnMappings.Add("[Sýra No]", "[Sýra No]");
                objbulk.ColumnMappings.Add("[Alýþ Faturasýnýn Tarihi]", "[Alýþ Faturasýnýn Tarihi]");
                objbulk.ColumnMappings.Add("[Alýþ Faturasýnýn Serisi]", "[Alýþ Faturasýnýn Serisi]");
                objbulk.ColumnMappings.Add("[Alýþ Faturasýnýn Sýra No'su]", "[Alýþ Faturasýnýn Sýra No'su]");
                objbulk.ColumnMappings.Add("[Belge No]", "[Belge No]");
                objbulk.ColumnMappings.Add("[Satýcýnýn Adý - Soyadý / Ünvaný]", "[Satýcýnýn Adý - Soyadý / Ünvaný]");
                objbulk.ColumnMappings.Add("[Satýcýnýn Vergi Kimlik Numarasý / TC Kimlik Numarasý]", "[Satýcýnýn Vergi Kimlik Numarasý / TC Kimlik Numarasý]");
                objbulk.ColumnMappings.Add("[Alýnan Mal ve/veya Hizmetin Cinsi]", "[Alýnan Mal ve/veya Hizmetin Cinsi]");
                objbulk.ColumnMappings.Add("[Alýnan Mal ve/veya Hizmetin Miktarý]", "[Alýnan Mal ve/veya Hizmetin Miktarý]");
                objbulk.ColumnMappings.Add("[Alýnan Mal ve/veya Hizmetin KDV Hariç Tutarý]", "[Alýnan Mal ve/veya Hizmetin KDV Hariç Tutarý]");
                objbulk.ColumnMappings.Add("[KDV'si]", "[KDV'si]");
                objbulk.ColumnMappings.Add("[Bünyeye Giren Mal ve/veya Hizmetin KDV'si]", "[Bünyeye Giren Mal ve/veya Hizmetin KDV'si]");
                objbulk.ColumnMappings.Add("[GGB Tescil No'su (Alýþ Ýthalat Ýse)]", "[GGB Tescil No'su (Alýþ Ýthalat Ýse)]");
                objbulk.ColumnMappings.Add("[Belgeye Ýliþkin Ýade Hakký Doðuran Ýþlem Türü]", "[Belgeye Ýliþkin Ýade Hakký Doðuran Ýþlem Türü]");
                objbulk.ColumnMappings.Add("[Yüklenim Türü]", "[Yüklenim Türü]");
                objbulk.ColumnMappings.Add("[Belgenin Ýndirime Konu Edildiði KDV Dönemi]", "[Belgenin Ýndirime Konu Edildiði KDV Dönemi]");
                objbulk.ColumnMappings.Add("[Belgenin Yüklenildiði KDV Dönemi]", "[Belgenin Yüklenildiði KDV Dönemi]");
                objbulk.ColumnMappings.Add("[Aracýn Plaka Numarasý]", "[Aracýn Plaka Numarasý]");
                objbulk.ColumnMappings.Add("[Aracýn Þasi Numarasý]", "[Aracýn Þasi Numarasý]");
                sqlConnection.Open();
                objbulk.WriteToServer(Exceldt);
                sqlConnection.Close();
                MessageBox.Show("Veri tabanýna eklendi.", "Bilgi Kutusu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Dosyanýn sayfa isminin 'Sayfa1' olduðundan ve seçtiðiniz dosyanýn içeriðinde Yüklenilen KDV Listesi olduðundan emin olun!!! "), "Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                MessageBox.Show("hata nedeni:" + ex);
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
        void veriTabanýnaGönder2()
        {
            excelbaglan1();
            baglanti2.Open();
            dt = new DataTable();
            oda2 = new OleDbDataAdapter("Select * from [Sayfa1$]", baglanti2);
            oda2.Fill(dt);
            dataGridView1.DataSource = dt;
            baglanti2.Close();
        }
        void veriTabanýnaGönder1()
        {
            excelbaglan2();
            baglanti3.Open();
            dt2 = new DataTable();
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
            button4.Visible = false;
            button5.Visible = false; 
        }
        private void button4_Click(object sender, EventArgs e)
        {
            od2.Filter = "Excell|*.xls;*.xlsx;";
            DialogResult dr2 = od2.ShowDialog();
            if (dr2 == DialogResult.Abort)
                return;
            if (dr2 == DialogResult.Cancel)
                return;
            tabloadý.Text = od2.FileName.ToString();
            veriTabanýnaGönder2();
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
                tabloadý.Text = od2.SafeFileName;
                button4.Enabled = true;
                bw2.Dispose();
            };
            button4.Enabled = false;
            button4.Visible = false;
            bw2.RunWorkerAsync();
            button2.Visible= true;
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }
        private void tabloadý2_MouseDoubleClick(object sender, MouseEventArgs e)
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
            tabloadý2.Text = od.FileName.ToString();
            veriTabanýnaGönder1();
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
                tabloadý2.Text = od.SafeFileName;
                button5.Enabled = true;
                bw.Dispose();
            };
            button5.Enabled = false;
            button5.Visible = false; 
            bw.RunWorkerAsync();
            button1.Visible = true;
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
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                sqlConnection1 = new SqlConnection(@"server=(localdb)\MSSQLLocalDB;Initial Catalog=Sibernetik;Integrated Security=SSPI");
                sqlConnection1.Open();
                sqlCommand1 = new SqlCommand("delete from [Toplam_Fatura]", sqlConnection1);
                sqlCommand1.ExecuteNonQuery();
                MessageBox.Show("Veri tabanýndan silindi.");
                button5.Visible = true;
                sqlConnection1.Close();
                button1.Visible=false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Uygulamayý kapatýn ve tekrar baþlatýn daha sonra iþlemlerinizi yapabilirsiniz");
                MessageBox.Show("hata nedeni:"+ex);            
            }
            dt2.Rows.Clear();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                sqlConnection2 = new SqlConnection(@"server=(localdb)\MSSQLLocalDB;Initial Catalog=Sibernetik;Integrated Security=SSPI");
                sqlConnection2.Open();
                sqlCommand2 = new SqlCommand("delete from [Yüklenilen KDV Listesi2]", sqlConnection2);
                sqlCommand2.ExecuteNonQuery();
                MessageBox.Show("Veri tabanýndan silindi.");
                button4.Visible = true;
                sqlConnection2.Close();
                button2.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Uygulamayý kapatýn ve tekrar baþlatýn daha sonra iþlemlerinizi yapabilirsiniz");
                MessageBox.Show("hata nedeni:" + ex);
            }
            dt.Rows.Clear();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                sqlConnection3 = new SqlConnection(@"server=(localdb)\MSSQLLocalDB;Initial Catalog=Sibernetik;Integrated Security=SSPI");
                sqlConnection3.Open();
                
                sqlCommand3 = new SqlCommand("select * from Toplam_Fatura a where not exists" +
                    "(select * from[Yüklenilen KDV Listesi2] b where a.Belge_No = b.[Belge No]" +
                    " and a.Tarih = b.[Alýþ Faturasýnýn Tarihi] and a.Vergi_No_TC_Kimlik_No = b.[Satýcýnýn Vergi Kimlik Numarasý / TC Kimlik Numarasý]); ", sqlConnection3);
                sda = new SqlDataAdapter(sqlCommand3);
                sda.Fill(dt3);
                dataGridView4.DataSource = dt3;
                MessageBox.Show("Kod çalýþtý");
                sqlConnection3.Close();
                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel_Dosyasý|*.xlsx" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            using (XLWorkbook workbook = new XLWorkbook())
                            {
                                workbook.Worksheets.Add(dt3, "deneme");
                                workbook.SaveAs(sfd.FileName);
                            }
                            MessageBox.Show("Excel dosyasý oluþturuldu.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Excel Dosyasý Oluþturulamadý", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Uygulamayý kapatýn ve tekrar deneyin!!!");
                MessageBox.Show("hata nedeni:" + ex);
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                sqlConnection4 = new SqlConnection(@"server=(localdb)\MSSQLLocalDB;Initial Catalog=Sibernetik;Integrated Security=SSPI");
                sqlConnection4.Open();
                sqlCommand4 = new SqlCommand("select * from [Toplam_Fatura] a inner join [Yüklenilen KDV Listesi2] b on a.[Belge_No]=b.[Belge No];", sqlConnection4);
                sda2 = new SqlDataAdapter(sqlCommand4);
                sda2.Fill(dt4);
                dataGridView2.DataSource = dt4;
                MessageBox.Show("Kod çalýþtý");
                sqlConnection4.Close();
                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel_Dosyasý|*.xlsx" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            using (XLWorkbook workbook = new XLWorkbook())
                            {
                                workbook.Worksheets.Add(dt4, "deneme");
                                workbook.SaveAs(sfd.FileName);
                            }
                            MessageBox.Show("Excel dosyasý oluþturuldu.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Excel Dosyasý Oluþturulamadý", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Uygulamayý kapatýn ve tekrar deneyin!!!!");
                MessageBox.Show("hata nedeni:" + ex);
            }
        }
    }
}