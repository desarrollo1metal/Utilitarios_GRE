using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using System.Windows.Forms;
using System.Configuration;

namespace Utilitarios_GRE
{
    public partial class Form1 : Form
    {
        static Company oCompany = new Company();
        static bool AbortarEnError = false;
        static bool CerrarAlFinalizar = false;

        public Form1()
        {


            conectarBDApi();
            
            InitializeComponent();
        }

        public void conectarBDApi()
        { 
            
        }

        private void SetApplication()
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Iniciando conexión");
            InitializeCompany();
            if (oCompany.Connect() == 0)
                Console.WriteLine("Conectado a " + oCompany.CompanyName);
            else
            {
                string errorout = oCompany.GetLastErrorDescription().ToString();
                Console.WriteLine(oCompany.GetLastErrorDescription());
                Console.WriteLine("Presione enter para finalizar...");
                Console.ReadLine();
                Environment.Exit(0);
            }
        }


        static void InitializeCompany()
        {
            oCompany = new Company();

            switch (ConfigurationManager.AppSettings["SAPDbType"].ToUpper()) //Tipo de BD
            {
                case "2005": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2005; break;
                case "2008": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2008; break;
                case "2012": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2012; break;
                case "2014": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2014; break;
                case "2016": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2016; break;
                case "2017": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2017; break;
                case "2019": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2019; break;
                case "HANA": oCompany.DbServerType = BoDataServerTypes.dst_HANADB; break;
            }
            oCompany.DbUserName = ConfigurationManager.AppSettings["SAPDbUserName"];
            oCompany.DbPassword = ConfigurationManager.AppSettings["SAPDbPassword"];
            oCompany.Server = ConfigurationManager.AppSettings["SAPServer"];
            oCompany.LicenseServer = ConfigurationManager.AppSettings["SAPLicenseServer"]; 
            oCompany.CompanyDB = ConfigurationManager.AppSettings["SAPDB"];
            oCompany.UserName = ConfigurationManager.AppSettings["SAPUserName"];
            oCompany.Password = ConfigurationManager.AppSettings["SAPPassword"];
            oCompany.language = BoSuppLangs.ln_Spanish_La;
            oCompany.UseTrusted = false;
            oCompany.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;
        }

    }
}
