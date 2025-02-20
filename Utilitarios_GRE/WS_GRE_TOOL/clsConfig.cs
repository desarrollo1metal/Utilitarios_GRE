using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WS_GRE_TOOL
{
    public class clsConfig
    {

        public Servidor ServidorSAP { get; set; }
        public SociedadBD[] Sociedades { get; set; }

        public int TiempoEspera { get; set; } = 300;

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


        public string Path1FirmaOrigen { get; set; }
        public string Path1FirmaError { get; set; }
        public string Path1ProcesadoFirma { get; set; }



        public string Path2FirmaOrigen { get; set; }
        public string Path2FirmaError { get; set; }
        public string Path2ProcesadoFirma { get; set; }



        public string Path1CdrZip { get; set; }
        public string Path1CdrProcesado { get; set; }
        public string Path1CdrError { get; set; }


        public string Path2CdrZip { get; set; }
        public string Path2CdrProcesado { get; set; }
        public string Path2CdrError { get; set; }


        public string Pathbackupzip { get; set; }

    }

}
