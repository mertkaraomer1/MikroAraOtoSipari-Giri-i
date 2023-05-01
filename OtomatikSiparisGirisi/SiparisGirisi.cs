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
                // Dosya se�me penceresi a�mak i�in
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyas� |*.xlsx";
                file.ShowDialog();

                // se�ti�imiz excel'in tam yolu
                string tamYol = file.FileName;

                //Excel ba�lant� adresi
                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                //ba�lant� 
                OleDbConnection baglanti = new(baglantiAdresi);

                //t�m verileri se�mek i�in select sorgumuz. Sayfa1 olan k�sm� Excel'de hangi sayfay� a�mak istiyosan�z oray� yazabilirsiniz.
                OleDbCommand komut = new OleDbCommand("Select * From [" + textBox1.Text + "$]", baglanti);

                //ba�lant�y� a��yoruz.
                baglanti.Open();

                //Gelen verileri DataAdapter'e at�yoruz.
                OleDbDataAdapter da = new OleDbDataAdapter(komut);

                //Grid'imiz i�in bir DataTable olu�turuyoruz.
                DataTable data = new DataTable();

                //DataAdapter'da ki verileri data ad�ndaki DataTable'a dolduruyoruz.
                da.Fill(data);

                //DataGrid'imizin kayna��n� olu�turdu�umuz DataTable ile dolduruyoruz.
                dataGridView1.DataSource = data;

            }
            catch (Exception ex)
            {
                // Hata al�rsak ekrana bast�r�yoruz.
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

                    var IKN = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value); //0 s�t�n
                    var Musteri_Adi = Convert.ToString("MNG KARGO"); //6 s�t�n
                    var Malzeme_Adi = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value).ToString(); //7 s�t�n
                    var Barkod = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value).ToString(); //9 s�t�n
                    var DMO_Urun_No = Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value).ToString(); //10 s�t�n
                    var Miktar = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value); //12 s�t�n
                    var Satin_Alma_Sip_No = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value);//14 s�t�n
                    var Tekif_Tutar� = Convert.ToDateTime(dataGridView1.Rows[i].Cells[7].Value).ToString("yyyy-MM-dd"); //16 s�t�n
                    var Sipari�_Parti_No = Convert.ToDateTime(dataGridView1.Rows[i].Cells[7].Value).ToString("yyyy-MM-dd"); //22 s�t�n
                    var Teslim_Adresi = Convert.ToDateTime(dataGridView1.Rows[i].Cells[7].Value).ToString("yyyy-MM-dd"); //24 s�t�n

                    baglanti.Open();
                    SqlCommand komut = new SqlCommand("INSERT INTO SIPARISLER (Gonderilcek_firma,Kargo_Sirketi,Desi_KG,Fiyat,Adet,Depo,�L,Tarih) VALUES ('" + IKN + "' , '" + Musteri_Adi + "','" + Malzeme_Adi + "' , '" + Barkod + "' , '" + DMO_Urun_No + "', '" + Miktar + "','" + Satin_Alma_Sip_No + "','" + Tekif_Tutar� + "','" + Sipari�_Parti_No + "','" + Teslim_Adresi + "')", baglanti);
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
                MessageBox.Show("BA�ARI �LE KAYDED�LD�...");
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
            // DataGridView1 ve DataGridView2'deki verileri y�kleyin
            // DataGridView1'deki "barkod" s�tunu DataGridView2'deki "barkod" s�tunuyla e�le�melidir
            // DataGridView2'deki "stok kodu" s�tunu DataGridView3'deki "stok kodu" s�tunuyla e�le�melidir

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string barcode = Convert.ToString(dataGridView1.Rows[i].Cells[8].Value);

                for (int j = 0; j < dataGridView2.Rows.Count; j++)
                {
                    if (barcode == Convert.ToString(dataGridView2.Rows[j].Cells[1].Value))
                    {
                        // E�le�en bir barkod bulundu
                        // DataGridView2'deki stok kodunu al�n ve DataGridView3'e yaz�n
                        string stockCode = Convert.ToString(dataGridView2.Rows[j].Cells[0].Value);

                        dataGridView3.Rows.Add(stockCode);
                        break;
                    }
                }

            }

        }
    }
}