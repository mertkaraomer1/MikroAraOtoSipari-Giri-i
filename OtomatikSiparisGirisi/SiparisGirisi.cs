using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace OtomatikSiparisGirisi
{
    public partial class SiparisGirisi : Form
    {
        public SiparisGirisi()
        {
            InitializeComponent();
        }


        SqlConnection baglanti;
        DataSet ds;
        SqlDataAdapter da;

        private void SiparisGirisi_Load(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=MERTSANAL;Initial Catalog=MikroDB_V16_ERMEDAS;User ID=sa;Password=1234;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            dataGridView3.ColumnCount = 1;
            dataGridView3.Columns[0].Name = "sto_kod";

        }


        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            try
            {
                // Dosya seçme penceresi açmak için
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyasý |*.xlsx";
                file.ShowDialog();

                // seçtiðimiz excel'in tam yolu
                string tamYol = file.FileName;

                //Excel baðlantý adresi
                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                //baðlantý 
                OleDbConnection baglanti = new(baglantiAdresi);

                //tüm verileri seçmek için select sorgumuz. Sayfa1 olan kýsmý Excel'de hangi sayfayý açmak istiyosanýz orayý yazabilirsiniz.
                OleDbCommand komut = new OleDbCommand("Select * From [" + textBox1.Text + "$]", baglanti);

                //baðlantýyý açýyoruz.
                baglanti.Open();

                //Gelen verileri DataAdapter'e atýyoruz.
                OleDbDataAdapter da = new OleDbDataAdapter(komut);

                //Grid'imiz için bir DataTable oluþturuyoruz.
                DataTable data = new DataTable();

                //DataAdapter'da ki verileri data adýndaki DataTable'a dolduruyoruz.
                da.Fill(data);

                //DataGrid'imizin kaynaðýný oluþturduðumuz DataTable ile dolduruyoruz.
                dataGridView1.DataSource = data;

            }
            catch (Exception ex)
            {
                // Hata alýrsak ekrana bastýrýyoruz.
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (baglanti.State == ConnectionState.Open)
                    {
                        baglanti.Close();
                    }

                    var IKN = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value); //0 sütün
                    var Musteri_Adi = Convert.ToString("MNG KARGO"); //6 sütün
                    var Malzeme_Adi = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value).ToString(); //7 sütün
                    var Barkod = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value).ToString(); //9 sütün
                    var DMO_Urun_No = Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value).ToString(); //10 sütün
                    var Miktar = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value); //12 sütün
                    var Satin_Alma_Sip_No = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value);//14 sütün
                    var Tekif_Tutarý = Convert.ToDateTime(dataGridView1.Rows[i].Cells[7].Value).ToString("yyyy-MM-dd"); //16 sütün
                    var Sipariþ_Parti_No = Convert.ToDateTime(dataGridView1.Rows[i].Cells[7].Value).ToString("yyyy-MM-dd"); //22 sütün
                    var Teslim_Adresi = Convert.ToDateTime(dataGridView1.Rows[i].Cells[7].Value).ToString("yyyy-MM-dd"); //24 sütün

                    baglanti.Open();
                    SqlCommand komut = new SqlCommand("INSERT INTO SIPARISLER (Gonderilcek_firma,Kargo_Sirketi,Desi_KG,Fiyat,Adet,Depo,ÝL,Tarih) VALUES ('" + IKN + "' , '" + Musteri_Adi + "','" + Malzeme_Adi + "' , '" + Barkod + "' , '" + DMO_Urun_No + "', '" + Miktar + "','" + Satin_Alma_Sip_No + "','" + Tekif_Tutarý + "','" + Sipariþ_Parti_No + "','" + Teslim_Adresi + "')", baglanti);
                    komut.ExecuteNonQuery();
                }
                baglanti.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("HATA VAR!!!");
            }
            finally
            {
                MessageBox.Show("BAÞARI ÝLE KAYDEDÝLDÝ...");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            da = new SqlDataAdapter("Select sto_kod,bar_kodu from  stoklar left join BARKOD_TANIMLARI on  bar_stokkodu=sto_kod where bar_master=1", baglanti);
            ds = new DataSet();
            baglanti.Open();
            da.Fill(ds, "stoklar");
            dataGridView2.DataSource = ds.Tables["stoklar"];
            baglanti.Close();
        }
        string sto_kod;
        private void button4_Click(object sender, EventArgs e)
        {
            // DataGridView1 ve DataGridView2'deki verileri yükleyin
            // DataGridView1'deki "barkod" sütunu DataGridView2'deki "barkod" sütunuyla eþleþmelidir
            // DataGridView2'deki "stok kodu" sütunu DataGridView3'deki "stok kodu" sütunuyla eþleþmelidir

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string barcode = Convert.ToString(dataGridView1.Rows[i].Cells[8].Value);

                for (int j = 0; j < dataGridView2.Rows.Count; j++)
                {
                    if (barcode == Convert.ToString(dataGridView2.Rows[j].Cells[1].Value))
                    {
                        // Eþleþen bir barkod bulundu
                        // DataGridView2'deki stok kodunu alýn ve DataGridView3'e yazýn
                        string stockCode = Convert.ToString(dataGridView2.Rows[j].Cells[0].Value);

                        dataGridView3.Rows.Add(stockCode);
                        break;
                    }
                }

            }

        }
    }
}