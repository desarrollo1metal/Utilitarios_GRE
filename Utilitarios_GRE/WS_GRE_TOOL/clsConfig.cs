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

        public string PathFirmaOrigen { get; set; }
        public string PathFirmaOrigenIp3 { get; set; }
        public string PathFirmaError { get; set; }
        public string PathProcesadoFirma { get; set; }



        public string PathCdrZip { get; set; }
        public string PathCdrZipIp3 { get; set; }
        public string PathCdrProcesado { get; set; }
        public string PathCdrError { get; set; }
        public string Pathbackupzip { get; set; }

    }

}
