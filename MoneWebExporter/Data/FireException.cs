using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter.Data
{
    class FireException : Exception
    {
        public FireException(string message)
            : base(message)
        {

        }
    }
}
