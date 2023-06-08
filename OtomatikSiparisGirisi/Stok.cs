using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OtomatikSiparisGirisi
{
    public class Stok
    {
        [Description("sto_kod")]
        public string StokKodu { get; set; }

        [Description("bar_kodu")]
        public string Barkod { get; set; }
    }
}
