using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APITipoCambio
{
    class ApiResponse
    {

        public DateTime fecha { get; set; }
        public string compra { get; set; }
        public string venta { get; set; }

    }

    public class TipoCambio
    {

        public DateTime fecha { get; set; }
        public string Divisa { get; set; }
        public double Valor { get; set; }

    }
}
