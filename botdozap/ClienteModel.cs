using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MandarMensagem
{
    internal class ClienteModel
    {
        [Description("Nome")]
        public string Cliente { get; set; }

        [Description("Numero")]
        public string Numero { get; set; }
    }
}