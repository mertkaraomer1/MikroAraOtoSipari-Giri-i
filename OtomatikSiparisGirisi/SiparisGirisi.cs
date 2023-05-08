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
        //SqlDataAdapter daa;
        private void SiparisGirisi_Load(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=MERTSANAL;Initial Catalog=MikroDB_V16_ERMEDAS;User ID=sa;Password=1234;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            dataGridView3.ColumnCount = 128;
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
            dataGridView3.Columns[12].Name = "sip_satici_kod";
            dataGridView3.Columns[13].Name = "Açýklama";
            dataGridView3.Columns[14].Name = "sip_aciklama2";
            dataGridView3.Columns[15].Name = "Sip_Field";
            dataGridView3.Columns[16].Name = "sip_DBCno";
            dataGridView3.Columns[17].Name = "sip_SpecRECno";
            dataGridView3.Columns[18].Name = "sip_iptal";
            dataGridView3.Columns[19].Name = "sip_hidden";
            dataGridView3.Columns[20].Name = "sip_kilitli";
            dataGridView3.Columns[21].Name = "sip_degisti";
            dataGridView3.Columns[22].Name = "sip_checksum";
            dataGridView3.Columns[23].Name = "sip_special1";
            dataGridView3.Columns[24].Name = "sip_special2";
            dataGridView3.Columns[25].Name = "sip_special3";
            dataGridView3.Columns[26].Name = "sip_firmano";
            dataGridView3.Columns[27].Name = "sip_subeno";
            dataGridView3.Columns[28].Name = "sip_Bitiþ Tarihi";
            dataGridView3.Columns[29].Name = "sip_Tarihi";
            dataGridView3.Columns[30].Name = "sip_tip";
            dataGridView3.Columns[31].Name = "sip_cins";
            dataGridView3.Columns[32].Name = "sip_belge_tarih";
            dataGridView3.Columns[33].Name = "sip_birim_pntr";
            dataGridView3.Columns[34].Name = "sip_teslim_miktar";
            dataGridView3.Columns[35].Name = "sip_iskonto_1";
            dataGridView3.Columns[36].Name = "sip_iskonto_2";
            dataGridView3.Columns[37].Name = "sip_iskonto_3";
            dataGridView3.Columns[38].Name = "sip_iskonto_4";
            dataGridView3.Columns[39].Name = "sip_iskonto_5";
            dataGridView3.Columns[40].Name = "sip_iskonto_6";
            dataGridView3.Columns[41].Name = "sip_masraf_1";
            dataGridView3.Columns[42].Name = "sip_masraf_2";
            dataGridView3.Columns[43].Name = "sip_masraf_3";
            dataGridView3.Columns[44].Name = "sip_masraf_4";
            dataGridView3.Columns[45].Name = "sip_vergi_pntr";
            dataGridView3.Columns[46].Name = "sip_vergi";
            dataGridView3.Columns[47].Name = "sip_masvergi_pntr";
            dataGridView3.Columns[48].Name = "sip_masvergi";
            dataGridView3.Columns[49].Name = "sip_opno";
            dataGridView3.Columns[50].Name = "sip_depono";
            dataGridView3.Columns[51].Name = "sip_OnaylayanKulNo";
            dataGridView3.Columns[52].Name = "sip_vergisiz_fl";
            dataGridView3.Columns[53].Name = "sip_kapat_fl";
            dataGridView3.Columns[54].Name = "sip_promosyon_fl";
            dataGridView3.Columns[55].Name = "sip_cari_sormerk";
            dataGridView3.Columns[56].Name = "sip_stok_sormerk";
            dataGridView3.Columns[57].Name = "sip_cari_grupno";
            dataGridView3.Columns[58].Name = "sip_doviz_cinsi";
            dataGridView3.Columns[59].Name = "sip_doviz_kuru ";
            dataGridView3.Columns[60].Name = "sip_alt_doviz_kuru ";
            dataGridView3.Columns[61].Name = "sip_adresno";
            dataGridView3.Columns[62].Name = "sip_teslimturu";
            dataGridView3.Columns[63].Name = "sip_cagrilabilir_fl";
            dataGridView3.Columns[64].Name = "sip_prosip_uid";
            dataGridView3.Columns[65].Name = "sip_iskonto1";
            dataGridView3.Columns[66].Name = "sip_iskonto2";
            dataGridView3.Columns[67].Name = "sip_iskonto3";
            dataGridView3.Columns[68].Name = "sip_iskonto4";
            dataGridView3.Columns[69].Name = "sip_iskonto5";
            dataGridView3.Columns[70].Name = "sip_iskonto6 ";
            dataGridView3.Columns[71].Name = "sip_masraf1 ";
            dataGridView3.Columns[72].Name = "sip_masraf2";
            dataGridView3.Columns[73].Name = "sip_masraf3";
            dataGridView3.Columns[74].Name = "sip_masraf4";
            dataGridView3.Columns[75].Name = "sip_isk1";
            dataGridView3.Columns[76].Name = "sip_isk2";
            dataGridView3.Columns[77].Name = "sip_isk3";
            dataGridView3.Columns[78].Name = "sip_isk4";
            dataGridView3.Columns[79].Name = "sip_isk5";
            dataGridView3.Columns[80].Name = "sip_isk6";
            dataGridView3.Columns[81].Name = "sip_mas1";
            dataGridView3.Columns[82].Name = "sip_mas2";
            dataGridView3.Columns[83].Name = "sip_mas3";
            dataGridView3.Columns[84].Name = "sip_mas4";
            dataGridView3.Columns[85].Name = "sip_Exp_Imp_Kodu";
            dataGridView3.Columns[86].Name = " sip_kar_orani";
            dataGridView3.Columns[87].Name = " sip_durumu";
            dataGridView3.Columns[88].Name = "sip_stal_uid";
            dataGridView3.Columns[89].Name = "sip_planlananmiktar";
            dataGridView3.Columns[90].Name = "sip_teklif_uid";
            dataGridView3.Columns[91].Name = "sip_parti_kodu";
            dataGridView3.Columns[92].Name = "sip_lot_no";
            dataGridView3.Columns[93].Name = "sip_projekodu";
            dataGridView3.Columns[94].Name = "sip_fiyat_liste_no";
            dataGridView3.Columns[95].Name = "sip_Otv_Pntr";
            dataGridView3.Columns[96].Name = "sip_Otv_Vergi";
            dataGridView3.Columns[97].Name = "sip_otvtutari";
            dataGridView3.Columns[98].Name = "sip_OtvVergisiz_Fl";
            dataGridView3.Columns[99].Name = "sip_paket_kod";
            dataGridView3.Columns[100].Name = "sip_Rez_uid";
            dataGridView3.Columns[101].Name = "sip_harekettipi";
            dataGridView3.Columns[102].Name = "sip_yetkili_uid";
            dataGridView3.Columns[103].Name = "ip_kapatmanedenkod";
            dataGridView3.Columns[104].Name = " sip_gecerlilik_tarihi";
            dataGridView3.Columns[105].Name = "sip_onodeme_evrak_tip";
            dataGridView3.Columns[106].Name = "sip_onodeme_evrak_seri";
            dataGridView3.Columns[107].Name = "sip_onodeme_evrak_sira";
            dataGridView3.Columns[108].Name = "sip_rezervasyon_miktari";
            dataGridView3.Columns[109].Name = "sip_rezerveden_teslim_edilen";
            dataGridView3.Columns[110].Name = "sip_HareketGrupKodu1";
            dataGridView3.Columns[111].Name = "sip_HareketGrupKodu2";
            dataGridView3.Columns[112].Name = "sip_HareketGrupKodu3";
            dataGridView3.Columns[113].Name = " sip_Olcu1";
            dataGridView3.Columns[114].Name = "sip_Olcu2";
            dataGridView3.Columns[115].Name = "sip_Olcu3";
            dataGridView3.Columns[116].Name = "sip_Olcu4";
            dataGridView3.Columns[117].Name = "sip_Olcu5";
            dataGridView3.Columns[118].Name = "sip_FormulMiktarNo";
            dataGridView3.Columns[119].Name = "sip_FormulMiktar";
            dataGridView3.Columns[120].Name = "sip_satis_fiyat_doviz_cinsi";
            dataGridView3.Columns[121].Name = "sip_satis_fiyat_doviz_kuru";
            dataGridView3.Columns[122].Name = "sip_eticaret_kanal_kodu";
            dataGridView3.Columns[123].Name = "sip_Tevkifat_turu";
            dataGridView3.Columns[124].Name = "sip_otv_tevkifat_turu";
            dataGridView3.Columns[125].Name = "sip_otv_tevkifat_tutari";
            dataGridView3.Columns[126].Name = "sip_create_user";
            dataGridView3.Columns[127].Name = "sip_lastup_user";
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
            da = new SqlDataAdapter("Select som_kod,som_isim from  SORUMLULUK_MERKEZLERI ", baglanti);
            ds = new DataSet();
            baglanti.Open();
            da.Fill(ds, "stoklar");
            dataGridView2.DataSource = ds.Tables["stoklar"];
            da.Fill(ds, "SORUMLULUK_MERKEZLERI");
            dataGridView4.DataSource = ds.Tables["SORUMLULUK_MERKEZLERI"];
            baglanti.Close();
        }
        string sto_kod;
        private void button4_Click(object sender, EventArgs e)
        {
            string sip_Rez_uid = "00000000 - 0000 - 0000 - 0000 - 000000000000";
            int sip_harekettipi = 0;
            string sip_yetkili_uid = "00000000 - 0000 - 0000 - 0000 - 000000000000";
            string sip_kapatmanedenkod = "";
            DateTime sip_gecerlilik_tarihi = DateTime.Now;
            int sip_onodeme_evrak_tip = 0;
            string sip_onodeme_evrak_seri = "";
            int sip_onodeme_evrak_sira = 0;
            int sip_rezervasyon_miktari = 0;
            int sip_rezerveden_teslim_edilen = 0;
            string sip_HareketGrupKodu1 = "";
            string sip_HareketGrupKodu2 = "";
            string sip_HareketGrupKodu3 = "";
            int sip_Olcu1 = 0;
            int sip_Olcu2 = 0;
            int sip_Olcu3 = 0;
            int sip_Olcu4 = 0;
            int sip_Olcu5 = 0;
            int sip_FormulMiktarNo = 0;
            int sip_FormulMiktar = 0;
            int sip_satis_fiyat_doviz_cinsi = 0;
            int sip_satis_fiyat_doviz_kuru = 0;
            string sip_eticaret_kanal_kodu = "";
            int sip_Tevkifat_turu = 0;
            int sip_otv_tevkifat_turu = 0;
            int sip_otv_tevkifat_tutari = 0;
            string sip_projekodu = "";
            int sip_fiyat_liste_no = 0;
            int sip_Otv_Pntr = 0;
            int sip_Otv_Vergi = 0;
            int sip_otvtutari = 0;
            int sip_OtvVergisiz_Fl = 0;
            string sip_paket_kod = "";
            int sip_lot_no = 0;
            string sip_parti_kodu = "";
            string sip_stal_uid = "00000000 - 0000 - 0000 - 0000 - 000000000000";
            int sip_planlananmiktar = 0;
            string sip_teklif_uid = "00000000 - 0000 - 0000 - 0000 - 000000000000";
            string sip_Exp_Imp_Kodu = "";
            int sip_kar_orani = 0;
            int sip_durumu = 0;
            int sip_isk1 = 0;
            int sip_isk2 = 0;
            int sip_isk3 = 0;
            int sip_isk4 = 0;
            int sip_isk5 = 0;
            int sip_isk6 = 0;
            int sip_mas1 = 0;
            int sip_mas2 = 0;
            int sip_mas3 = 0;
            int sip_mas4 = 0;
            int sip_iskonto1 = 0;
            int sip_iskonto2 = 1;
            int sip_iskonto3 = 1;
            int sip_iskonto4 = 1;
            int sip_iskonto5 = 1;
            int sip_iskonto6 = 1;
            int sip_masraf1 = 0;
            int sip_masraf2 = 0;
            int sip_masraf3 = 0;
            int sip_masraf4 = 0;
            string sip_prosip_uid = "00000000 - 0000 - 0000 - 0000 - 000000000000";
            int sip_adresno = 1;
            string sip_teslimturu = "";
            int sip_cagrilabilir_fl = 0;
            string sip_cari_sormerk = "";
            string sip_stok_sormerk = "";
            int sip_cari_grupno = 0;
            int sip_doviz_cinsi = 0;
            int sip_doviz_kuru = 1;
            int sip_OnaylayanKulNo = 0;
            int sip_vergisiz_fl = 0;
            int sip_kapat_fl = 0;
            int sip_promosyon_fl = 0;
            int sip_depono = 4;
            string sip_aciklama2 = "";
            int sip_opno = -90;
            int sip_masvergi_pntr = 0;
            int sip_masvergi = 0;
            int sip_vergi_pntr = 3;
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
            double sip_alt_doviz_kuru = Convert.ToDouble(textBox4.Text);
            DateTime User_create_date = DateTime.Now;
            DateTime User_lastup_date = dateTimePicker1.Value;
            int sip_create_user = Convert.ToInt32(textBox2.Text);
            int sip_lastup_user = Convert.ToInt32(textBox5.Text);
            DateTime sip_tarih = DateTime.Now;
            DateTime sip_teslim_tarih = DateTime.Now;
            DateTime sip_belge_tarih = DateTime.Now;
            string sip_satici_kod = textBox3.Text.ToString();
            string SrmMrkz1 = "120.01.0754";

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
                        string aciklama = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                        int miktar = Convert.ToInt32(dataGridView1.Rows[i].Cells[11].Value);
                        int sip_teslim_miktar = Convert.ToInt32(dataGridView1.Rows[i].Cells[11].Value);
                        int tutar = Convert.ToInt32(dataGridView1.Rows[i].Cells[15].Value);
                        string sip_belgeno = Convert.ToString(dataGridView1.Rows[i].Cells[13].Value);
                        string sip_evrakno_seri = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                        int Toplamtutar = Convert.ToInt32(miktar * tutar);
                        double sip_vergi = Convert.ToDouble(Math.Round(Toplamtutar * 0.08, 2));

                        dataGridView3.Rows.Add(Guid.NewGuid(),
                                               stockCode,
                                               User_create_date,
                                               User_lastup_date,
                                               SrmMrkz1,
                                               miktar,
                                               tutar,
                                               Toplamtutar,
                                               sip_evrakno_seri,
                                               arttýr,
                                               Sýrano,
                                               sip_belgeno,
                                               sip_satici_kod,
                                               aciklama,
                                               sip_aciklama2,
                                               sipField,
                                               sip_DBCno,
                                               sip_SpecRECno,
                                               sip_iptal,
                                               sip_hidden,
                                               sip_kilitli,
                                               sip_degisti,
                                               sip_checksum,
                                               sip_special1,
                                               sip_special2,
                                               sip_special3,
                                               sip_firmano,
                                               sip_subeno,
                                               sip_tarih,
                                               sip_teslim_tarih,
                                               sip_tip,
                                               sip_cins,
                                               sip_belge_tarih,
                                               sip_birim_pntr,
                                               sip_teslim_miktar,
                                               sip_iskonto_1,
                                               sip_iskonto_2,
                                               sip_iskonto_3,
                                               sip_iskonto_4,
                                               sip_iskonto_5,
                                               sip_iskonto_6,
                                               sip_masraf_1,
                                               sip_masraf_2,
                                               sip_masraf_3,
                                               sip_masraf_4,
                                               sip_vergi_pntr,
                                               sip_vergi,
                                               sip_masvergi_pntr,
                                               sip_masvergi,
                                               sip_opno,
                                               sip_depono,
                                               sip_OnaylayanKulNo,
                                               sip_vergisiz_fl,
                                               sip_kapat_fl,
                                               sip_promosyon_fl,
                                               sip_cari_sormerk,
                                               sip_stok_sormerk,
                                               sip_cari_grupno,
                                               sip_doviz_cinsi,
                                               sip_doviz_kuru,
                                               sip_alt_doviz_kuru,
                                               sip_adresno,
                                               sip_teslimturu,
                                               sip_cagrilabilir_fl,
                                               sip_prosip_uid,
                                               sip_iskonto1,
                                               sip_iskonto2,
                                               sip_iskonto3,
                                               sip_iskonto4,
                                               sip_iskonto5,
                                               sip_iskonto6,
                                               sip_masraf1,
                                               sip_masraf2,
                                               sip_masraf3,
                                               sip_masraf4,
                                               sip_isk1,
                                               sip_isk2,
                                               sip_isk3,
                                               sip_isk4,
                                               sip_isk5,
                                               sip_isk6,
                                               sip_mas1,
                                               sip_mas2,
                                               sip_mas3,
                                               sip_mas4,
                                               sip_Exp_Imp_Kodu,
                                               sip_kar_orani,
                                               sip_durumu,
                                               sip_stal_uid,
                                               sip_planlananmiktar,
                                               sip_teklif_uid,
                                               sip_parti_kodu,
                                               sip_lot_no,
                                               sip_projekodu,
                                               sip_fiyat_liste_no,
                                               sip_Otv_Pntr,
                                               sip_Otv_Vergi,
                                               sip_otvtutari,
                                               sip_OtvVergisiz_Fl,
                                               sip_paket_kod,
                                               sip_Rez_uid,
                                               sip_harekettipi,
                                               sip_yetkili_uid,
                                               sip_kapatmanedenkod,
                                               sip_gecerlilik_tarihi,
                                               sip_onodeme_evrak_tip,
                                               sip_onodeme_evrak_seri,
                                               sip_onodeme_evrak_sira,
                                               sip_rezervasyon_miktari,
                                               sip_rezerveden_teslim_edilen,
                                               sip_HareketGrupKodu1,
                                               sip_HareketGrupKodu2,
                                               sip_HareketGrupKodu3,
                                               sip_Olcu1,
                                               sip_Olcu2,
                                               sip_Olcu3,
                                               sip_Olcu4,
                                               sip_Olcu5,
                                               sip_FormulMiktarNo,
                                               sip_FormulMiktar,
                                               sip_satis_fiyat_doviz_cinsi,
                                               sip_satis_fiyat_doviz_kuru,
                                               sip_eticaret_kanal_kodu,
                                               sip_Tevkifat_turu,
                                               sip_otv_tevkifat_turu,
                                               sip_otv_tevkifat_tutari,
                                               sip_create_user,
                                               sip_lastup_user);

                    }
                }
            }
        }
    }
}