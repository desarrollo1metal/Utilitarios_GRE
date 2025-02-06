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



            Log.Information("Iniciando conexión");
            //InitializeCompany();
            if (oCompany.Connect() == 0)
                Log.Information("Conectado a " + oCompany.CompanyName);
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
                    Log.Error("No se encontró el archivo de configuración Config.xml en la ruta " + ruta.ToString());
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
                Log.Error(ex, ex.ToString());
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (!LeerConfig()) return;

            if (Conexion.ConfiguracionGeneral.Sociedades.Length == 0)
            {
                Log.Error("No se encontraron BD en el archivo de configuración, por favor revise");
                return;
            }

            Log.Information("BD encontradas en archivo son " + Conexion.ConfiguracionGeneral.Sociedades.Length.ToString());

            foreach (SociedadBD sociedad in Conexion.ConfiguracionGeneral.Sociedades)
            {

                string msj;
                Conexion.InitializeCompany(sociedad);
                //logger.Info("Procesando socios");
                Log.Information("Conectando a la BD " + sociedad.DbName + " con DI API");
                Conexion.oCompany.Connect();
                if (Conexion.oCompany.Connected == false)
                {
                    int rpta = 0;
                    Conexion.oCompany.GetLastError(out rpta, out msj);
                    Log.Error(rpta.ToString() + " -- " + msj.ToString());
                }
                else
                {
                    Log.Information("Conectado satisfactoriamente a BD " + sociedad.DbName + " con DI API");
                    //logger.Info("Conectado satisfactoriamente");
                    Conexion.InicializarVarGlob();
                    ////Util.FileMGMT.CreateFolder(AppDomain.CurrentDomain.BaseDirectory + "tmp");

                    //Procesar procesar = new Procesar(sociedad.DbName);
                    procesarXML(sociedad.PathFirma, sociedad.PathProcesadoFirma);


                    ////// procesos por cada BD
                    ////logger.Info("Iniciando el Proceso");
                    //procesar.IniciarProceso();

                    //logger.Info("Finalizando el Proceso");

                    Conexion.oCompany.Disconnect();
                    Log.Information("Desconectando de la BD " + sociedad.DbName);
                    //procesar.Dispose();
                    //procesar = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                Conexion.DestroyCompany();
            }

        }


        public void procesarXML(string origenT, string destinoT)
        {


            string rutaOrigen = @origenT;
            string rutaDestino = @destinoT;

            try
            {
                // 📌 Asegurar que la carpeta destino existe
                if (!Directory.Exists(rutaDestino))
                {
                    Directory.CreateDirectory(rutaDestino);
                    Log.Information($"Carpeta creada: {rutaDestino}");
                }

                // 📌 Obtener todos los archivos XML en la carpeta origen
                string[] archivos = Directory.GetFiles(rutaOrigen, "*.xml");

                Log.Information($"Carpeta Origen: {rutaOrigen}");
                Log.Information($"Carpeta Destino: {rutaDestino}");

                foreach (string archivo in archivos)
                {
                    string nombreArchivo = Path.GetFileName(archivo); // Obtener solo el nombre
                    string destino = Path.Combine(rutaDestino, nombreArchivo); // Ruta completa en destino

                    // 📌 Verificar si el archivo ya existe en la carpeta destino
                    if (File.Exists(destino))
                    {
                        File.Delete(destino); // Eliminar el archivo para poder moverlo
                        Log.Information($"Archivo existente eliminado: {nombreArchivo}");
                        
                    }

                    // 📌 Mover el archivo
                    File.Move(archivo, destino);
                    Log.Information($"Archivo movido: {nombreArchivo}");
                    //Console.WriteLine($"Archivo movido: {nombreArchivo}");
                }

                Log.Information("✅ Proceso completado.");
            }
            catch (Exception ex)
            {
                Log.Error($"❌ Error: {ex.Message}");
            }


        }
    }
}
