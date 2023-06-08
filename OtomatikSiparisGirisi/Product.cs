using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OtomatikSiparisGirisi
{
    public class Product
    {
        [Description("KODU")]
        public string Kodu { get; set; }

        [Description("ADI")]
        public string Adi { get; set; }

        [Description("müşteri no")]
        public string MusteriNo { get; set; }
    }
}
