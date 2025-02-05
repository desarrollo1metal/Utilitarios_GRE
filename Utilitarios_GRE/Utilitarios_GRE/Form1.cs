using Serilog;
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
using System.Diagnostics;
using System.IO;

namespace Utilitarios_GRE
{
    public partial class Form1 : Form
    {
        static Company oCompany = new Company();
        static bool AbortarEnError = false;
        static bool CerrarAlFinalizar = false;
        public static clsConfig ConfiguracionGeneral;
        

        public Form1()
        {
            Log.Logger = new LoggerConfiguration()
           //.WriteTo.Console()  // Muestra logs en la consola
           .WriteTo.File("logs/app.log", rollingInterval: RollingInterval.Day) // Guarda logs en archivos diarios
           .CreateLogger();

            Log.Information("Aplicación iniciada.");
            
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
            //InitializeCompany();
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


        static void InitializeCompany(SociedadBD sociedad)
        {
            //oCompany = new Company();

            ////oCompany = new Company();
            //switch (ConfiguracionGeneral.ServidorSAP.DbType)
            //{
            //    case "2005": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2005; break;
            //    case "2008": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2008; break;
            //    case "2012": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2012; break;
            //    case "2014": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2014; break;
            //    case "2016": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2016; break;
            //    case "2017": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2017; break;
            //    case "2019": oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2019; break;

            //    case "HANA": oCompany.DbServerType = BoDataServerTypes.dst_HANADB; break;
            //}
            //oCompany.DbUserName = ConfiguracionGeneral.ServidorSAP.DbUserName;
            //oCompany.DbPassword = ConfiguracionGeneral.ServidorSAP.DbPassword;
            //oCompany.Server = ConfiguracionGeneral.ServidorSAP.Server;
            //oCompany.CompanyDB = sociedad.DbName;
            //oCompany.UserName = sociedad.SAPuser;
            //oCompany.Password = sociedad.SAPpassword;
            //oCompany.language = BoSuppLangs.ln_Spanish_La;
            //oCompany.LicenseServer = ConfiguracionGeneral.ServidorSAP.LicenseServer;
            //oCompany.UseTrusted = false;


        }

        static bool LeerConfig()
        {
            try
            {
                string ruta = AppDomain.CurrentDomain.BaseDirectory + "Config.xml";
                if (!File.Exists(ruta))
                {
                    Console.WriteLine("No se encontró el archivo de configuración, ejecute la utilidad SN_Server_Config");
                    Environment.Exit(0);
                }
                clsConfig config = new clsConfig();
                using (var stringReader = new StreamReader(ruta))
                {
                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(clsConfig));
                    config = serializer.Deserialize(stringReader) as clsConfig;
                };
                Conexion.ConfiguracionGeneral = config;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (!LeerConfig()) return;

            if (Conexion.ConfiguracionGeneral.Sociedades.Length == 0)
            {
                //logger.Error("No se encontraron BD en el archivo de configuración, por favor revise");
                return;
            }

            foreach (SociedadBD sociedad in Conexion.ConfiguracionGeneral.Sociedades)
            {
                string msj;
                Conexion.InitializeCompany(sociedad);
                //logger.Info("Procesando socios");
                Log.Information("Conectando a la BD " + sociedad.DbName);
                Conexion.oCompany.Connect();
                if (Conexion.oCompany.Connected == false)
                {
                    int rpta = 0;
                    Conexion.oCompany.GetLastError(out rpta, out msj);
                    //logger.Error(msj);
                }
                else
                {
                    Log.Information("Conectado satisfactoriamente ");
                    //logger.Info("Conectado satisfactoriamente");
                    Conexion.InicializarVarGlob();
                    ////Util.FileMGMT.CreateFolder(AppDomain.CurrentDomain.BaseDirectory + "tmp");
                    //Procesar procesar = new Procesar(sociedad.DbName);
                    ////// procesos por cada BD
                    ////logger.Info("Iniciando el Proceso");
                    //procesar.IniciarProceso();

                    //logger.Info("Finalizando el Proceso");

                    Conexion.oCompany.Disconnect();
                    //logger.Info("Desconectando de la BD " + sociedad.DbName);
                    //procesar.Dispose();
                    //procesar = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                Conexion.DestroyCompany();
            }



        }
    }
}
