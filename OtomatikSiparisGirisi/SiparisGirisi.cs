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
            dataGridView3.ColumnCount =44;
            dataGridView3.Columns[0].Name = "GUÝD";
            dataGridView3.Columns[1].Name = "stok Kodu";
            dataGridView3.Columns[2].Name = "sip. Giriþ Tarihi";
            dataGridView3.Columns[3].Name = "sip. Güncelleme Tarihi";
            dataGridView3.Columns[4].Name = "1.SRM Merkezi";
            dataGridView3.Columns[5].Name = "Miktar";
            dataGridView3.Columns[6].Name = "Tutar";
            dataGridView3.Columns[7].Name = "Toplam Tutar";
            dataGridView3.Columns[8].Name = "Evrak No";
            dataGridView3.Columns[9].Name = "Evrak No Sýra";
            dataGridView3.Columns[10].Name = "Sipariþ Satýr No";
            dataGridView3.Columns[11].Name = "sip_belgeno";
            dataGridView3.Columns[12].Name = "Satýcý No";
            dataGridView3.Columns[13].Name = "Açýklama";
            dataGridView3.Columns[14].Name = "Sip_Field";
            dataGridView3.Columns[15].Name = "sip_DBCno";
            dataGridView3.Columns[16].Name = "sip_SpecRECno";
            dataGridView3.Columns[17].Name = "sip_iptal";
            dataGridView3.Columns[18].Name = "sip_hidden";
            dataGridView3.Columns[19].Name = "sip_kilitli";
            dataGridView3.Columns[20].Name = "sip_degisti";
            dataGridView3.Columns[21].Name = "sip_checksum";
            dataGridView3.Columns[22].Name = "sip_special1";
            dataGridView3.Columns[23].Name = "sip_special2";
            dataGridView3.Columns[24].Name = "sip_special3";
            dataGridView3.Columns[25].Name = "sip_firmano";
            dataGridView3.Columns[26].Name = "sip_subeno";
            dataGridView3.Columns[27].Name = "sip_Bitiþ Tarihi";
            dataGridView3.Columns[28].Name = "sip_Tarihi";
            dataGridView3.Columns[29].Name = "sip_tip";
            dataGridView3.Columns[30].Name = "sip_cins";
            dataGridView3.Columns[31].Name = "sip_belge_tarih";
            dataGridView3.Columns[32].Name = "sip_birim_pntr";
            dataGridView3.Columns[33].Name = "sip_teslim_miktar";
            dataGridView3.Columns[34].Name = "sip_iskonto_1";
            dataGridView3.Columns[35].Name = "sip_iskonto_2";
            dataGridView3.Columns[36].Name = "sip_iskonto_3";
            dataGridView3.Columns[37].Name = "sip_iskonto_4";
            dataGridView3.Columns[38].Name = "sip_iskonto_5";
            dataGridView3.Columns[39].Name = "sip_iskonto_6";
            dataGridView3.Columns[40].Name = "sip_masraf_1";
            dataGridView3.Columns[41].Name = "sip_masraf_2";
            dataGridView3.Columns[42].Name = "sip_masraf_3";
            dataGridView3.Columns[43].Name = "sip_masraf_4";
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
            int sip_masraf_1 = 0;
            int sip_masraf_2 = 0;
            int sip_masraf_3 = 0;
            int sip_masraf_4 = 0;
            int sip_iskonto_1 = 0;
            int sip_iskonto_2 = 0;
            int sip_iskonto_3 = 0;
            int sip_iskonto_4 = 0;
            int sip_iskonto_5 = 0;
            int sip_iskonto_6 = 0;
            int sip_birim_pntr = 1;
            string sip_belgeno = "";
            int sip_cins = 0;
            int sip_tip = 0;
            int sip_firmano = 0;
            int sip_subeno = 0;
            string sip_special1 = "";
            string sip_special2 = "";
            string sip_special3 = "";
            int sip_kilitli = 0;
            int sip_degisti = 0;
            int sip_checksum = 0;
            int sip_hidden = 0;
            int sip_iptal = 0;
            int sip_SpecRECno = 0;
            int sip_DBCno = 0;
            int sipField = 21;
            string SaticiNo = textBox3.Text.ToString();
            string SrmMrkz1 = "120.01.0754";
            string EvrakNo = textBox2.Text.ToString();
            // DataGridView1 ve DataGridView2'deki verileri yükleyin
            // DataGridView1'deki "barkod" sütunu DataGridView2'deki "barkod" sütunuyla eþleþmelidir
            // DataGridView2'deki "stok kodu" sütunu DataGridView3'deki "stok kodu" sütunuyla eþleþmelidir
            string previousValue = Convert.ToString(dataGridView1.Rows[0].Cells[3].Value); // sütundaki ilk hücrenin deðerini al

            int arttýr = 0;
            int Sýrano = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string currentValue = Convert.ToString(dataGridView1.Rows[i].Cells[3].Value); // sütundaki diðer hücrelerin deðerini al
                if (currentValue != previousValue) // deðerler farklýysa
                {
                    arttýr++;
                    Sýrano = 0;
                }
                else if (currentValue == previousValue)
                {
                    Sýrano++;
                }
                previousValue = currentValue; // bir sonraki hücre için önceki hücrenin deðerini sakla

                string barcode = Convert.ToString(dataGridView1.Rows[i].Cells[8].Value);
                for (int j = 0; j < dataGridView2.Rows.Count; j++)
                {
                    if (barcode == Convert.ToString(dataGridView2.Rows[j].Cells[1].Value))
                    {
                        // Eþleþen bir barkod bulundu
                        // DataGridView2'deki stok kodunu alýn ve DataGridView3'e yazýn
                        string stockCode = Convert.ToString(dataGridView2.Rows[j].Cells[0].Value);

                        int miktar = Convert.ToInt32(dataGridView1.Rows[i].Cells[11].Value);
                        int sip_teslim_miktar= Convert.ToInt32(dataGridView1.Rows[i].Cells[11].Value);
                        int tutar = Convert.ToInt32(dataGridView1.Rows[i].Cells[15].Value);
                        string aciklama = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                        int Toplamtutar = Convert.ToInt32(miktar * tutar);

                        dataGridView3.Rows.Add(Guid.NewGuid(), stockCode, DateTime.Now.ToString(), dateTimePicker1.Value.ToString(), SrmMrkz1, miktar, tutar, Toplamtutar, EvrakNo, arttýr, Sýrano,sip_belgeno, SaticiNo, aciklama, sipField, sip_DBCno, sip_SpecRECno, sip_iptal, sip_hidden, sip_kilitli, sip_degisti, sip_checksum, sip_special1, sip_special2, sip_special3, sip_firmano, sip_subeno,dateTimePicker2.Value.ToString(),DateTime.Now.ToString(),sip_tip,sip_cins, DateTime.Now.ToString(), sip_birim_pntr, sip_teslim_miktar,sip_iskonto_1,sip_iskonto_2,sip_iskonto_3,sip_iskonto_4,sip_iskonto_5,sip_iskonto_6,sip_masraf_1,sip_masraf_2,sip_masraf_3,sip_masraf_4);

                    }
                }
            }
        }


    }
}