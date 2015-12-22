using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter
{
    public static class DateHelper
    {
        public static DateTime? FromString(string date)
        {
            try
            {
                return DateTime.Parse(date);
            }
            catch(Exception e)
            {
                return null;
            }
        }

        public static string ToString(DateTime? date)
        {
            if (date != null)
                return date.Value.ToString();
            else
                return null;
        }
    }
}
