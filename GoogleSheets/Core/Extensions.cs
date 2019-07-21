using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheets.Core
{
    public static class Extensions
    {
        public static string ToDataString(this ExtendedValue value)
        {
            // try all the properties :)
            if (!(value.BoolValue is null)) return value.BoolValue.Value.ToString();
            if (!(value.NumberValue is null)) return value.NumberValue.Value.ToString(CultureInfo.CurrentCulture);
            if (!(value.StringValue is null)) return value.StringValue;
            return null;
        }
    }
}
