using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OtomatikSiparisGirisi
{
    public static class TypeHelper
    {
        public static float ToFloat(this string stringValue)
        {
            float floatValue =0f;
            if (float.TryParse(stringValue, out floatValue))
            {
            }
            return floatValue;
        }
    }
}
