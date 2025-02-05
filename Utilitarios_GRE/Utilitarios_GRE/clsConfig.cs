using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utilitarios_GRE
{
    public class clsConfig
    {

        public Servidor ServidorSAP { get; set; }
        public SociedadBD[] Sociedades { get; set; }
        public int NRUCenviar { get; set; } = 10;
        public int IntentosConexion { get; set; } = 10;
        public int IntentosCaptcha { get; set; } = 10;
        public int TiempoEsperaSUNAT { get; set; } = 300;
        public bool SociosInactivos { get; set; } = false;
        public bool ActualizarDireccionEntrega { get; set; } = true;
        public bool ActualizarRazonSocial { get; set; } = false;
        public bool LogCompleto { get; set; } = true;
        public bool CargaMasiva { get; set; } = true;
        public bool ActualizarTC { get; set; } = true;
        public string MetodoActualizacion { get; set; }
        public string OrigenTC { get; set; }
        public string TipoTC { get; set; }
    }

    public class Servidor
    {
        public string DbUserName { get; set; }
        public string DbPassword { get; set; }
        public string DbType { get; set; }
        public string Server { get; set; }
        public string LicenseServer { get; set; }
    }

    public class SociedadBD
    {
        public string DbName { get; set; }
        public string SAPuser { get; set; }
        public string SAPpassword { get; set; }
    }


}
