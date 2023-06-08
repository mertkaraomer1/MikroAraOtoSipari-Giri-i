using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OtomatikSiparisGirisi
{
    public class Customer
    {
        public Guid sip_Guid =>Guid.NewGuid();
        [Description("IKN")]
        public string IKN { get; set; }

        [Description("İl Adı")]
        public string IlAdi { get; set; }

        [Description("Müşteri No")]
        public string MusteriNo { get; set; }

        [Description("VKN")]
        public string Vkn { get; set; }

        [Description("GLN Kodu")]
        public string GlnKodu { get; set; }

        [Description("Müşteri Adı")]
        public string MusteriAdi { get; set; }

        [Description("Malzeme Adı")]
        public string MalzemeAdi { get; set; }

        [Description("Malzeme Kodu")]
        public string MalzemeKodu { get; set; }

        [Description("Barkod")]
        public string Barkod { get; set; }

        [Description("DMO Ürün Numarası")]
        public string DmoUrunNumarasi { get; set; }

        [Description("Ürün Adı")]
        public string UrunAdi { get; set; }

        [Description("Miktar")]
        public string Miktar { get; set; }

        [Description("Ölçü Birimi")]
        public string OlcuBirimi { get; set; }

        [Description("Satın Alma Sipariş No")]
        public string SiparisNo { get; set; }

        [Description("Alıma Esas Fiyat")]
        public string AlimaEsasFiyat { get; set; }

        [Description("Teklif Tutarı")]
        public string TeklifTutari { get; set; }

        [Description("Teklif Tutarı Yerli Mi")]
        public string TeklifTutariYerliMi { get; set; }

        [Description("E-İhale Tutarı")]
        public string EihaleTutari { get; set; }

        [Description("E-İhale Tutarı Yerli Mi")]
        public string EihaleTutariYerliMi { get; set; }

        [Description("E-İhale Üzerine Kalan Firma Ünvanı")]
        public string EihaleKalanFirmaUnvani { get; set; }

        [Description("E-İhale Durumu")]
        public string EihaleDurumu { get; set; }

        [Description("Sipariş Parti No")]
        public string SiparisPartiNo { get; set; }

        [Description("Kategori")]
        public string Kategori { get; set; }

        [Description("Teslimat Adresi")]
        public string TeslimatAdresi { get; set; }
        public string SorMekKod { get; set; }
        public string StokKod { get; set; }
        public int sip_evrakno_sira { get; set; }
        public int sip_satirno { get; set; }
        public double sip_tutar { get; set; }
        public double sip_vergi => sip_tutar * 0.08;
    }
}
