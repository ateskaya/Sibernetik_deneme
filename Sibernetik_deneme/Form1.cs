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
using excel=Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Data.OleDb;
using System.DirectoryServices;

namespace Sibernetik_deneme
{
    public partial class Form1 : Form
    {
         SqlConnection baglanti;
         SqlDataAdapter da;
         SqlCommand komut;
         OleDbConnection baglanti2;
         OleDbConnection baglanti3;
         OleDbDataAdapter oda;
         OleDbDataAdapter oda2;
         DataTable dt = new DataTable();
         DataTable dt2 = new DataTable();
         DataTable dt3 = new DataTable();

        public Form1()
        {
        
            InitializeComponent();
        }

        



        void MusteriGetir()
        {
            baglanti = new SqlConnection(@"server=(localdb)\MSSQLLocalDB;Initial Catalog=Sibernetik;Integrated Security=SSPI");
            baglanti.Open();
            da = new SqlDataAdapter("Select *From vicual_deneme", baglanti);
            DataTable tablo = new DataTable();
            da.Fill(tablo);
            dataGridView1.DataSource = tablo;
            baglanti.Close();
        }
        void ServerBaglan()
        {
            baglanti = new SqlConnection(@"server=(localdb)\MSSQLLocalDB;Initial Catalog=Sibernetik;Integrated Security=SSPI");
            baglanti.Open();
            // da = new SqlDataAdapter("Select *From vicual_deneme", baglanti);
            // DataTable tablo = new DataTable();
            // da.Fill(tablo);
            // dataGridView1.DataSource = tablo;
            baglanti.Close();
        }
        void excelbaglan1()
        {
            baglanti2 = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + openFileDialog1.FileName + ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'");
            
        }
        void excelbaglan2()
        {
            baglanti3 = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + openFileDialog2.FileName + ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;'");
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
            dt = new DataTable();
            oda = new OleDbDataAdapter("Select * from [Sayfa1$]", baglanti3);
            oda.Fill(dt2);
            dataGridView3.DataSource = dt2;
            baglanti3.Close();

        }
        void veriTabanýnaGönder()
        {
                dt3.Rows.Clear();
                dt3.Columns.Clear();    
                DataRow dr = dt3.NewRow();
                baglanti2.Open();
                baglanti3.Open();
                dt3.Columns.Add("Belge No", typeof(string));
                dt3.Columns.Add("Firma Kodu", typeof(string));
                dt3.Columns.Add("Firma Adý", typeof(string));
                var result = from Sayfa1 in dt2.AsEnumerable()
                             join Sayfa2 in dt.AsEnumerable() on (string)Sayfa1["Belge No"] equals (string)Sayfa2["Alýþ Faturasýnýn Sýra No'su"]
                             select new
                             {
                                 Belgo_No = (string)Sayfa1["Belge No"],
                                 Firma_Kodu = (string)Sayfa1["Firma Kodu"],
                                 Firma_Adý = (string)Sayfa1["Firma Adý"]
                             };
                foreach (var item in result)
                {
                    dr["Belge No"] = item.Belgo_No;
                    dr["Firma Kodu"] = item.Firma_Kodu;
                    dr["Firma Adý"] = item.Firma_Adý;
                dt3.Rows.Add(dr.ItemArray);
                }
                dataGridView4.DataSource = dt3;
                baglanti2.Close();
                baglanti3.Close();
        }
        void openFileDialogBaglan1()
        {
            openFileDialog1.Title = "Lütfen Dosya Seçin";
            openFileDialog1.FileName = "Dosya seç";
            openFileDialog1.InitialDirectory = @"C:\Users\ensar\OneDrive\Masaüstü\sibernetik";
            openFileDialog1.Filter = "xlsx dosyalayarý(*.xlsx)|*.xlsx|xls dosyalarý(*.xls)|*.xls|Csv dosyalarý(*.csv)|*.csv";
        }
        void openFileDialogBaglan2()
        {
            openFileDialog2.Title = "Lütfen Dosya Seçin";
            openFileDialog2.FileName = "Dosya seç";
            openFileDialog2.InitialDirectory = @"C:\Users\ensar\OneDrive\Masaüstü\sibernetik";
            openFileDialog2.Filter = "xlsx dosyalayarý(*.xlsx)|*.xlsx|xls dosyalarý(*.xls)|*.xls|Csv dosyalarý(*.csv)|*.csv";
        }
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
        }
        private void Form1_Load_1(object sender, EventArgs e)
        {
            //MusteriGetir();
            //ServerBaglan();
            openFileDialogBaglan1();
            openFileDialogBaglan2();
            excelbaglan1();
            excelbaglan2();
        }
        

        private void button1_Click_1(object sender, EventArgs e)
        {
            veriTabanýnaGönder();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            veriTabanýnaGönder2();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            /*string sorgu = "Select * From vicual_deneme Where Miktar=@Miktar";
            komut = new SqlCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@Miktar", Convert.ToInt32(tabloadý.Text));
            baglanti.Open();
            komut.ExecuteNonQuery();
            baglanti.Close();
            MusteriGetir();*/
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            /*string sorgu = "Update musteri Set Belge_No=@Belge_No,KDV_Oraný=@KDV_Oraný,Birim=@Birim Where Miktar=@Miktar";
            komut = new SqlCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@Miktar", Convert.ToInt32(tabloadý.Text));
            komut.Parameters.AddWithValue("@Belge_No", textBox2.Text);
            komut.Parameters.AddWithValue("@KDV_Oraný", textBox3.Text);
            komut.Parameters.AddWithValue("@Birim", textBox4.Text);
            baglanti.Open();
            komut.ExecuteNonQuery();
            baglanti.Close();
            MusteriGetir();*/
        }

        private void dosyaismi_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                dosyaismi.Text = openFileDialog1.FileName;
                string tabloismi = openFileDialog1.SafeFileName;
                tabloadý.Text = tabloismi;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void tabloadý2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
        }

        private void dosyaismi2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                dosyaismi2.Text = openFileDialog2.FileName;
                string tabloismi = openFileDialog2.SafeFileName;
                tabloadý2.Text = tabloismi;
            }
        }

        private void dosyaismi2_TextChanged(object sender, EventArgs e)
        {
        }

        private void button5_Click(object sender, EventArgs e)
        {
            veriTabanýnaGönder1();
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void button8_Click(object sender, EventArgs e)
        {
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {

        }

    }
}