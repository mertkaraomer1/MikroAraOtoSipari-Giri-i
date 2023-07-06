using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using DataTable = System.Data.DataTable;


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
        SqlDataAdapter daa;
        private void SiparisGirisi_Load(object sender, EventArgs e)
        {
            baglanti = new SqlConnection("Data Source=HPSERVER;Initial Catalog=MikroDB_V16_ERMED;User ID=SA;Password=1234;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            //baglanti = new SqlConnection("Data Source=HPSERVER;Initial Catalog=MikroDB_V16_ERMED;User ID=SA;Password=1234;Connect Timeout=30;Encrypt=False;Trust Server Certificate=True;Application Intent=ReadWrite;Multi Subnet Failover=False");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //dataGridView1.Rows.Clear();
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
                System.Data.DataTable data = new System.Data.DataTable();

                //DataAdapter'da ki verileri data adýndaki DataTable'a dolduruyoruz.
                da.Fill(data);

                //DataGrid'imizin kaynaðýný oluþturduðumuz DataTable ile dolduruyoruz.
                //dataGridView1.DataSource = data;

                customerList = DataTableConverter.ConvertTo<Customer>(data);
                MessageBox.Show("Müþteri Listesi Yüklendi");
            }
            catch (Exception ex)
            {
                // Hata alýrsak ekrana bastýrýyoruz.
                MessageBox.Show(ex.Message);
            }
        }
        List<Customer> customerList = new List<Customer>();
        List<Product> productList = new List<Product>();
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (baglanti.State == ConnectionState.Open)
                {
                    baglanti.Close();
                }
                baglanti.Open();
                bool first = true;
                foreach (var item in customerList)
                {

                    string sqlQuery = @"INSERT INTO [dbo].[SIPARISLER] 
                                    ([sip_Guid], [sip_DBCno], [sip_SpecRECno], [sip_iptal], [sip_fileid], [sip_hidden], [sip_kilitli], 
                                    [sip_degisti], [sip_checksum], [sip_create_user], [sip_create_date], [sip_lastup_user], [sip_lastup_date], 
                                    [sip_special1], [sip_special2], [sip_special3], [sip_firmano], [sip_subeno], [sip_tarih], [sip_teslim_tarih], 
                                    [sip_tip], [sip_cins], [sip_evrakno_seri], [sip_evrakno_sira], [sip_satirno], [sip_belgeno], [sip_belge_tarih], 
                                    [sip_satici_kod], [sip_musteri_kod], [sip_stok_kod], [sip_b_fiyat], [sip_miktar], [sip_birim_pntr], 
                                    [sip_teslim_miktar], [sip_tutar], [sip_iskonto_1], [sip_iskonto_2], [sip_iskonto_3], [sip_iskonto_4], 
                                    [sip_iskonto_5], [sip_iskonto_6], [sip_masraf_1], [sip_masraf_2], [sip_masraf_3], [sip_masraf_4], 
                                    [sip_vergi_pntr], [sip_vergi], [sip_masvergi_pntr], [sip_masvergi], [sip_opno], [sip_aciklama], 
                                    [sip_aciklama2], [sip_depono], [sip_OnaylayanKulNo], [sip_vergisiz_fl], [sip_kapat_fl], [sip_promosyon_fl], 
                                    [sip_cari_sormerk], [sip_stok_sormerk], [sip_cari_grupno], [sip_doviz_cinsi], [sip_doviz_kuru], 
                                    [sip_alt_doviz_kuru], [sip_adresno], [sip_teslimturu], [sip_cagrilabilir_fl], [sip_prosip_uid], 
                                    [sip_iskonto1], [sip_iskonto2], [sip_iskonto3], [sip_iskonto4], [sip_iskonto5], [sip_iskonto6], 
                                    [sip_masraf1], [sip_masraf2], [sip_masraf3], [sip_masraf4], [sip_isk1], [sip_isk2], [sip_isk3], 
                                    [sip_isk4], [sip_isk5], [sip_isk6], [sip_mas1], [sip_mas2], [sip_mas3], [sip_mas4], [sip_Exp_Imp_Kodu], 
                                    [sip_kar_orani], [sip_durumu], [sip_stal_uid], [sip_planlananmiktar], [sip_teklif_uid], [sip_parti_kodu], 
                                    [sip_lot_no], [sip_projekodu], [sip_fiyat_liste_no], [sip_Otv_Pntr], [sip_Otv_Vergi], [sip_otvtutari], 
                                    [sip_OtvVergisiz_Fl], [sip_paket_kod], [sip_Rez_uid], [sip_harekettipi], [sip_yetkili_uid], 
                                    [sip_kapatmanedenkod], [sip_gecerlilik_tarihi], [sip_onodeme_evrak_tip], [sip_onodeme_evrak_seri], 
                                    [sip_onodeme_evrak_sira], [sip_rezervasyon_miktari], [sip_rezerveden_teslim_edilen], 
                                    [sip_HareketGrupKodu1], [sip_HareketGrupKodu2], [sip_HareketGrupKodu3], [sip_Olcu1], [sip_Olcu2], 
                                    [sip_Olcu3], [sip_Olcu4], [sip_Olcu5], [sip_FormulMiktarNo], [sip_FormulMiktar], 
                                    [sip_satis_fiyat_doviz_cinsi], [sip_satis_fiyat_doviz_kuru], [sip_eticaret_kanal_kodu], 
                                    [sip_Tevkifat_turu], [sip_otv_tevkifat_turu], [sip_otv_tevkifat_tutari])
                                    VALUES 
                                    (@sip_Guid, @sip_DBCno, @sip_SpecRECno, @sip_iptal, @sip_fileid, @sip_hidden, @sip_kilitli, 
                                    @sip_degisti, @sip_checksum, @sip_create_user, @sip_create_date, @sip_lastup_user, @sip_lastup_date, 
                                    @sip_special1, @sip_special2, @sip_special3, @sip_firmano, @sip_subeno, @sip_tarih, @sip_teslim_tarih, 
                                    @sip_tip, @sip_cins, @sip_evrakno_seri, @sip_evrakno_sira, @sip_satirno, @sip_belgeno, @sip_belge_tarih, 
                                    @sip_satici_kod, @sip_musteri_kod, @sip_stok_kod, @sip_b_fiyat, @sip_miktar, @sip_birim_pntr, 
                                    @sip_teslim_miktar, @sip_tutar, @sip_iskonto_1, @sip_iskonto_2, @sip_iskonto_3, @sip_iskonto_4, 
                                    @sip_iskonto_5, @sip_iskonto_6, @sip_masraf_1, @sip_masraf_2, @sip_masraf_3, @sip_masraf_4, 
                                    @sip_vergi_pntr, @sip_vergi, @sip_masvergi_pntr, @sip_masvergi, @sip_opno, @sip_aciklama, 
                                    @sip_aciklama2, @sip_depono, @sip_OnaylayanKulNo, @sip_vergisiz_fl, @sip_kapat_fl, @sip_promosyon_fl, 
                                    @sip_cari_sormerk, @sip_stok_sormerk, @sip_cari_grupno, @sip_doviz_cinsi, @sip_doviz_kuru, 
                                    @sip_alt_doviz_kuru, @sip_adresno, @sip_teslimturu, @sip_cagrilabilir_fl, @sip_prosip_uid, 
                                    @sip_iskonto1, @sip_iskonto2, @sip_iskonto3, @sip_iskonto4, @sip_iskonto5, @sip_iskonto6, 
                                    @sip_masraf1, @sip_masraf2, @sip_masraf3, @sip_masraf4, @sip_isk1, @sip_isk2, @sip_isk3, 
                                    @sip_isk4, @sip_isk5, @sip_isk6, @sip_mas1, @sip_mas2, @sip_mas3, @sip_mas4, @sip_Exp_Imp_Kodu, 
                                    @sip_kar_orani, @sip_durumu, @sip_stal_uid, @sip_planlananmiktar, @sip_teklif_uid, @sip_parti_kodu, 
                                    @sip_lot_no, @sip_projekodu, @sip_fiyat_liste_no, @sip_Otv_Pntr, @sip_Otv_Vergi, @sip_otvtutari, 
                                    @sip_OtvVergisiz_Fl, @sip_paket_kod, @sip_Rez_uid, @sip_harekettipi, @sip_yetkili_uid, 
                                    @sip_kapatmanedenkod, @sip_gecerlilik_tarihi, @sip_onodeme_evrak_tip, @sip_onodeme_evrak_seri, 
                                    @sip_onodeme_evrak_sira, @sip_rezervasyon_miktari, @sip_rezerveden_teslim_edilen, 
                                    @sip_HareketGrupKodu1, @sip_HareketGrupKodu2, @sip_HareketGrupKodu3, @sip_Olcu1, @sip_Olcu2, 
                                    @sip_Olcu3, @sip_Olcu4, @sip_Olcu5, @sip_FormulMiktarNo, @sip_FormulMiktar, 
                                    @sip_satis_fiyat_doviz_cinsi, @sip_satis_fiyat_doviz_kuru, @sip_eticaret_kanal_kodu, 
                                    @sip_Tevkifat_turu, @sip_otv_tevkifat_turu, @sip_otv_tevkifat_tutari)";

                    using (SqlCommand command = new SqlCommand(sqlQuery, baglanti))
                    {
                        // Set parameter values from the SiparislerData object
                        command.Parameters.AddWithValue("@sip_Guid", item.sip_Guid);
                        command.Parameters.AddWithValue("@sip_DBCno", 0);
                        command.Parameters.AddWithValue("@sip_SpecRECno", 0);
                        command.Parameters.AddWithValue("@sip_iptal", false);
                        command.Parameters.AddWithValue("@sip_fileid", 21);
                        command.Parameters.AddWithValue("@sip_hidden", false);
                        command.Parameters.AddWithValue("@sip_kilitli", false);
                        command.Parameters.AddWithValue("@sip_degisti", false);
                        command.Parameters.AddWithValue("@sip_checksum", 0);
                        command.Parameters.AddWithValue("@sip_create_user", textBox2.Text);
                        command.Parameters.AddWithValue("@sip_create_date", DateTime.Now);
                        command.Parameters.AddWithValue("@sip_lastup_user", textBox5.Text);
                        command.Parameters.AddWithValue("@sip_lastup_date", DateTime.Now);
                        command.Parameters.AddWithValue("@sip_special1", string.Empty);
                        command.Parameters.AddWithValue("@sip_special2", string.Empty);
                        command.Parameters.AddWithValue("@sip_special3", string.Empty);
                        command.Parameters.AddWithValue("@sip_firmano", 0);
                        command.Parameters.AddWithValue("@sip_subeno", 0);
                        command.Parameters.AddWithValue("@sip_tarih", DateTime.Now.Date);
                        command.Parameters.AddWithValue("@sip_teslim_tarih", DateTime.Now.Date);
                        command.Parameters.AddWithValue("@sip_tip", 0);
                        command.Parameters.AddWithValue("@sip_cins", 0);
                        command.Parameters.AddWithValue("@sip_evrakno_seri", item.IKN);
                        command.Parameters.AddWithValue("@sip_evrakno_sira", item.sip_evrakno_sira);
                        command.Parameters.AddWithValue("@sip_satirno", item.sip_satirno);
                        command.Parameters.AddWithValue("@sip_belgeno", item.SiparisNo);
                        command.Parameters.AddWithValue("@sip_belge_tarih", DateTime.Now.Date);
                        command.Parameters.AddWithValue("@sip_satici_kod", textBox3.Text);
                        command.Parameters.AddWithValue("@sip_musteri_kod", textBox9.Text);
                        command.Parameters.AddWithValue("@sip_stok_kod", item.StokKod);
                        command.Parameters.AddWithValue("@sip_b_fiyat", item.TeklifTutari.ToFloat());
                        command.Parameters.AddWithValue("@sip_miktar", item.Miktar.ToFloat());
                        command.Parameters.AddWithValue("@sip_birim_pntr", 1);
                        command.Parameters.AddWithValue("@sip_teslim_miktar", 0);
                        command.Parameters.AddWithValue("@sip_tutar", item.sip_tutar);
                        command.Parameters.AddWithValue("@sip_iskonto_1", 0);
                        command.Parameters.AddWithValue("@sip_iskonto_2", 0);
                        command.Parameters.AddWithValue("@sip_iskonto_3", 0);
                        command.Parameters.AddWithValue("@sip_iskonto_4", 0);
                        command.Parameters.AddWithValue("@sip_iskonto_5", 0);
                        command.Parameters.AddWithValue("@sip_iskonto_6", 0);
                        command.Parameters.AddWithValue("@sip_masraf_1", 0);
                        command.Parameters.AddWithValue("@sip_masraf_2", 0);
                        command.Parameters.AddWithValue("@sip_masraf_3", 0);
                        command.Parameters.AddWithValue("@sip_masraf_4", 0);
                        command.Parameters.AddWithValue("@sip_vergi_pntr", 3);
                        command.Parameters.AddWithValue("@sip_vergi", item.sip_vergi);
                        command.Parameters.AddWithValue("@sip_masvergi_pntr", 0);
                        command.Parameters.AddWithValue("@sip_masvergi", 0);
                        command.Parameters.AddWithValue("@sip_opno", textBox7.Text);
                        command.Parameters.AddWithValue("@sip_aciklama", "");
                        command.Parameters.AddWithValue("@sip_aciklama2", string.Empty);
                        command.Parameters.AddWithValue("@sip_depono", 4);
                        command.Parameters.AddWithValue("@sip_OnaylayanKulNo", 0);
                        command.Parameters.AddWithValue("@sip_vergisiz_fl", false);
                        command.Parameters.AddWithValue("@sip_kapat_fl", false);
                        command.Parameters.AddWithValue("@sip_promosyon_fl", false);
                        command.Parameters.AddWithValue("@sip_cari_sormerk", item.SorMekKod);
                        command.Parameters.AddWithValue("@sip_stok_sormerk", item.SorMekKod);
                        command.Parameters.AddWithValue("@sip_cari_grupno", 0);
                        command.Parameters.AddWithValue("@sip_doviz_cinsi", 0);
                        command.Parameters.AddWithValue("@sip_doviz_kuru", 1);
                        command.Parameters.AddWithValue("@sip_alt_doviz_kuru", textBox4.Text);
                        command.Parameters.AddWithValue("@sip_adresno", 1);
                        command.Parameters.AddWithValue("@sip_teslimturu", string.Empty);
                        command.Parameters.AddWithValue("@sip_cagrilabilir_fl", true);
                        command.Parameters.AddWithValue("@sip_prosip_uid", Guid.Empty);
                        command.Parameters.AddWithValue("@sip_iskonto1", 0);
                        command.Parameters.AddWithValue("@sip_iskonto2", 1);
                        command.Parameters.AddWithValue("@sip_iskonto3", 1);
                        command.Parameters.AddWithValue("@sip_iskonto4", 1);
                        command.Parameters.AddWithValue("@sip_iskonto5", 1);
                        command.Parameters.AddWithValue("@sip_iskonto6", 1);
                        command.Parameters.AddWithValue("@sip_masraf1", 1);
                        command.Parameters.AddWithValue("@sip_masraf2", 1);
                        command.Parameters.AddWithValue("@sip_masraf3", 1);
                        command.Parameters.AddWithValue("@sip_masraf4", 1);
                        command.Parameters.AddWithValue("@sip_isk1", false);
                        command.Parameters.AddWithValue("@sip_isk2", false);
                        command.Parameters.AddWithValue("@sip_isk3", false);
                        command.Parameters.AddWithValue("@sip_isk4", false);
                        command.Parameters.AddWithValue("@sip_isk5", false);
                        command.Parameters.AddWithValue("@sip_isk6", false);
                        command.Parameters.AddWithValue("@sip_mas1", false);
                        command.Parameters.AddWithValue("@sip_mas2", false);
                        command.Parameters.AddWithValue("@sip_mas3", false);
                        command.Parameters.AddWithValue("@sip_mas4", false);
                        command.Parameters.AddWithValue("@sip_Exp_Imp_Kodu", string.Empty);
                        command.Parameters.AddWithValue("@sip_kar_orani", 0);
                        command.Parameters.AddWithValue("@sip_durumu", 0);
                        command.Parameters.AddWithValue("@sip_stal_uid", Guid.Empty);
                        command.Parameters.AddWithValue("@sip_planlananmiktar", 0);
                        command.Parameters.AddWithValue("@sip_teklif_uid", Guid.Empty);
                        command.Parameters.AddWithValue("@sip_parti_kodu", string.Empty);
                        command.Parameters.AddWithValue("@sip_lot_no", 0);
                        command.Parameters.AddWithValue("@sip_projekodu", string.Empty);
                        command.Parameters.AddWithValue("@sip_fiyat_liste_no", 1);
                        command.Parameters.AddWithValue("@sip_Otv_Pntr", 0);
                        command.Parameters.AddWithValue("@sip_Otv_Vergi", 0);
                        command.Parameters.AddWithValue("@sip_otvtutari", 0);
                        command.Parameters.AddWithValue("@sip_OtvVergisiz_Fl", 0);
                        command.Parameters.AddWithValue("@sip_paket_kod", string.Empty);
                        command.Parameters.AddWithValue("@sip_Rez_uid", Guid.Empty);
                        command.Parameters.AddWithValue("@sip_harekettipi", 0);
                        command.Parameters.AddWithValue("@sip_yetkili_uid", Guid.Empty);
                        command.Parameters.AddWithValue("@sip_kapatmanedenkod", string.Empty);
                        command.Parameters.AddWithValue("@sip_gecerlilik_tarihi", "1899/12/30 00:00:00");
                        command.Parameters.AddWithValue("@sip_onodeme_evrak_tip", 0);
                        command.Parameters.AddWithValue("@sip_onodeme_evrak_seri", string.Empty);
                        command.Parameters.AddWithValue("@sip_onodeme_evrak_sira", 0);
                        command.Parameters.AddWithValue("@sip_rezervasyon_miktari", 0);
                        command.Parameters.AddWithValue("@sip_rezerveden_teslim_edilen", 0);
                        command.Parameters.AddWithValue("@sip_HareketGrupKodu1", string.Empty);
                        command.Parameters.AddWithValue("@sip_HareketGrupKodu2", string.Empty);
                        command.Parameters.AddWithValue("@sip_HareketGrupKodu3", string.Empty);
                        command.Parameters.AddWithValue("@sip_Olcu1", 0);
                        command.Parameters.AddWithValue("@sip_Olcu2", 0);
                        command.Parameters.AddWithValue("@sip_Olcu3", 0);
                        command.Parameters.AddWithValue("@sip_Olcu4", 0);
                        command.Parameters.AddWithValue("@sip_Olcu5", 0);
                        command.Parameters.AddWithValue("@sip_FormulMiktarNo", 0);
                        command.Parameters.AddWithValue("sip_FormulMiktar", 0);
                        command.Parameters.AddWithValue("sip_satis_fiyat_doviz_cinsi", 0);
                        command.Parameters.AddWithValue("sip_satis_fiyat_doviz_kuru", 0);
                        command.Parameters.AddWithValue("sip_eticaret_kanal_kodu", string.Empty);
                        command.Parameters.AddWithValue("sip_Tevkifat_turu", 0);
                        command.Parameters.AddWithValue("sip_otv_tevkifat_turu", 0);
                        command.Parameters.AddWithValue("sip_otv_tevkifat_tutari", 0);
                        //command.ExecuteNonQuery();
                        //baglanti.Close();
                        // Set the remaining parameters in a similar manner
                        string queryText = command.CommandText;

                        // Replace the parameter placeholder with the actual value
                        foreach (SqlParameter parameter in command.Parameters)
                        {
                            queryText = queryText.Replace(parameter.ParameterName, parameter.Value.ToString());
                        }

                        // Print or store the query text with the actual parameter value
                        Console.WriteLine(queryText);

                        command.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("BAÞARI ÝLE KAYDEDÝLDÝ...");
            }
            catch (Exception ex)
            {
                MessageBox.Show("HATA VAR!!!");
            }
            finally
            {
                baglanti.Close();
            }

        }


        string sto_kod;

        private void button4_Click(object sender, EventArgs e)
        {
            if (customerList.Count == 0)
            {
                MessageBox.Show("Müþteri listesini yükleyiniz.");
                return;
            }
            if (productList.Count == 0)
            {
                MessageBox.Show("Ürün stok listesini yükleyiniz.");
                return;
            }
            int counter = 0;
            int sayac = 1;
            string previousMusteriNo = null;
            int previousSip_evrakno_sira = 0;

            foreach (var item in customerList)
            {
                var musteri = productList.FirstOrDefault(x => x.MusteriNo == item.MusteriNo);
                if (musteri == null)
                    continue;
                item.SorMekKod = musteri.Kodu;
                if (item.MusteriNo != previousMusteriNo)
                {

                    counter++;
                    item.sip_evrakno_sira = counter;
                }
                else
                {

                    item.sip_evrakno_sira = counter;

                }
                if (item.sip_evrakno_sira == previousSip_evrakno_sira)
                {
                    item.sip_satirno = sayac;
                    sayac++;
                }
                else if (item.sip_evrakno_sira != previousSip_evrakno_sira)
                {
                    item.sip_satirno = 0; // Reset sip_satirno to 1 for a new sip_evrakno_sira
                    sayac = 1; // Reset sayac to 2 for a new sip_evrakno_sira
                }

                previousMusteriNo = item.MusteriNo;
                previousSip_evrakno_sira = item.sip_evrakno_sira;
                double miktarInt, tutarInt;

                if (double.TryParse(item.Miktar, out miktarInt) && double.TryParse(item.TeklifTutari, out tutarInt))
                {
                    item.sip_tutar = miktarInt * tutarInt;
                }
                else
                {
                    Console.WriteLine("Geçersiz miktar veya tutar.");
                }
                int index = item.IKN.IndexOf("/"); // '/' karakterinin indeksini bulur

                if (index != -1)
                {
                    string sonrasý = item.IKN.Substring(index + 1); // '/' karakterinden sonrasýný alýr
                    item.IKN = sonrasý + textBox8.Text;
                }

            }
            customerList = customerList.Where(x => !string.IsNullOrWhiteSpace(x.SorMekKod)).ToList();

            List<string> barkods = customerList.Select(x => x.Barkod).ToList();
            string combinedString = "'" + string.Join("','", barkods) + "'";
            da = new SqlDataAdapter("Select sto_kod,bar_kodu from  stoklar left join BARKOD_TANIMLARI on  bar_stokkodu=sto_kod where bar_kodu in (" + combinedString + ") ", baglanti);
            System.Data.DataTable dataTable = new System.Data.DataTable();
            baglanti.Open();
            da.Fill(dataTable); // object referansi degiskene bise atilmamissa gelir. = new yapilmamis yani. koda tek atmisim mk xd
            baglanti.Close();
            var stokList = DataTableConverter.ConvertTo<Stok>(dataTable);

            foreach (var item in customerList)
            {
                item.StokKod = stokList.FirstOrDefault(x => x.Barkod == item.Barkod)?.StokKodu ?? string.Empty;

            }
            dataGridView1.DataSource = customerList;
            // customerList datagridviewe
            return;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //dataGridView4.Rows.Clear();
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
                OleDbCommand komut = new OleDbCommand("Select * From [" + textBox6.Text + "$]", baglanti);

                //baðlantýyý açýyoruz.
                baglanti.Open();

                //Gelen verileri DataAdapter'e atýyoruz.
                OleDbDataAdapter da = new OleDbDataAdapter(komut);

                //Grid'imiz için bir DataTable oluþturuyoruz.
                System.Data.DataTable data = new System.Data.DataTable();

                //DataAdapter'da ki verileri data adýndaki DataTable'a dolduruyoruz.
                da.Fill(data);

                //DataGrid'imizin kaynaðýný oluþturduðumuz DataTable ile dolduruyoruz.
                // dataGridView4.DataSource = data;
                productList = DataTableConverter.ConvertTo<Product>(data);
                MessageBox.Show("Ürün Listesi Yüklendi");
            }
            catch (Exception ex)
            {
                // Hata alýrsak ekrana bastýrýyoruz.
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            // DataGridView'deki verileri bir DataTable'a kopyalayýn
            DataTable dt = new DataTable();

            // Sütunlarý ekle (tablo baþlýklarýný da dahil etmek için)
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                dt.Columns.Add(column.HeaderText, column.ValueType);
            }

            // Satýrlarý ekle
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataRow dataRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dataRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dataRow);
            }

            // Excel uygulamasýný baþlatýn
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            // Yeni bir Excel çalýþma kitabý oluþturun
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            // DataTable'ý Excel çalýþma sayfasýna aktarýn (tablo baþlýklarýný da dahil etmek için)
            int rowIndex = 1;

            // Baþlýklarý yaz
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                worksheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                worksheet.Cells[1, j + 1].Font.Bold = true;
            }

            // Verileri yaz
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowIndex++;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    worksheet.Cells[rowIndex, j + 1] = dt.Rows[i][j].ToString();
                }
            }

        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.Green;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.White;
        }

        private void button5_MouseEnter(object sender, EventArgs e)
        {
            button5.BackColor = Color.Green;
        }

        private void button5_MouseLeave(object sender, EventArgs e)
        {
            button5.BackColor = Color.White;
        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            button3.BackColor = Color.Green;
        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            button3.BackColor = Color.White;
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.BackColor = Color.Green;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.White;
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            button4.BackColor = Color.Green;
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            button4.BackColor = Color.White;
        }
    }

}