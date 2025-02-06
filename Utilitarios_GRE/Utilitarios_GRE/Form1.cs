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
using Serilog.Core;
using System.Collections;

namespace Utilitarios_GRE
{
    public partial class Form1 : Form
    {
        //static Company oCompany = new Company();
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



            //Log.Information("Iniciando conexión");
            ////InitializeCompany();
            //if (oCompany.Connect() == 0)
            //    Log.Information("Conectado a " + oCompany.CompanyName);
            //else
            //{
            //    string errorout = oCompany.GetLastErrorDescription().ToString();
            //    Console.WriteLine(oCompany.GetLastErrorDescription());
            //    Console.WriteLine("Presione enter para finalizar...");
            //    Console.ReadLine();
            //    Environment.Exit(0);
            //}
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

                    Log.Information("Finalizando el Proceso");

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


                Log.Information("Inicio de mover xml ");

                foreach (string archivo in archivos)
                {

                    try
                    {
                        string nombreArchivo = Path.GetFileName(archivo); // Obtener solo el nombre
                        string destino = Path.Combine(rutaDestino, nombreArchivo); // Ruta completa en destino

                        //ini procesar en SAP , adjuntar
                        if (procesarXMLAdjuntoSAP(archivo,nombreArchivo, origenT, destinoT))
                        {
                            
                        }
                        //fin procesar en SAP , adjuntar



                    }
                    catch (Exception ex)
                    {

                        Log.Error(ex,ex.ToString());
                    }

                }

                Log.Information("✅ Proceso completado.");
            }
            catch (Exception ex)
            {
                Log.Error($"❌ Error: {ex.Message}");
            }


        }

        private bool procesarXMLAdjuntoSAP(string archivot,string namefile ,string oringT ,string destinot) {

            bool val = false;
            string name = namefile;
            string[] dato = namefile.Split('-');

            string ruc = dato[0];
            string tipoDoc = dato[1];
            string serieDoc = dato[2];
            string correlativo = dato[3];
            string docentry = string.Empty;

            Log.Information($"Ruc = { ruc } : tipoDoc = {tipoDoc} : serieDoc = {serieDoc} : correlativo = {correlativo}");

            if (dato.Length == 4)
            {

                //ini adjuntar SAP
                Recordset oRecordSet = null;
                oRecordSet = (Recordset)Conexion.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                Log.Information("Buscando socios");
                string query = string.Empty;

                query = $@"SELECT DocEntry FROM ODLN WHERE U_BPP_MDTD = '09' AND U_BPP_MDSD = '{serieDoc}' AND U_BPP_MDCD = '{correlativo.Replace(".xml","") }' AND CANCELED <> 'Y'";

                //logger.Debug(query);
                oRecordSet.DoQuery(query);

                if (oRecordSet.RecordCount > 0)
                {
                    
                    while (!oRecordSet.EoF)
                    {
                        docentry = (oRecordSet.Fields.Item("DocEntry").Value.ToString());
                        
                        oRecordSet.MoveNext();
                    }
                    Log.Information("Se Encontrado el documento en SAP con DocEntry = " + docentry);
                    val true;
                }
                else
                {
                    Log.Error("No se encontraron Entrada");
                    return val = false;
                }


                // 🔹 Obtener la Entrega de Mercancía (ODLN)
                Documents oDelivery = (Documents)Conexion.oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                if (!oDelivery.GetByKey(Convert.ToInt32(docentry.Trim() )))  //5924)) // 1234 = ID de la Entrega (DocEntry)
                {
                    Log.Information("❌ No se encontró la Entrega de Mercancía.");
                   
                }

                // 🔹 Asignar el adjunto a la entrega
                //oDelivery.AttachmentEntry = attachEntry;
                //oDelivery.UserFields.Fields.Item("U_GMI_V1XML").Value = @"C:\TI_MIMSA\FILE_DAEMON\SF\SFS_v1.6\sunat_archivos\sfs\FIRMA\20565975812-09-T001-0003452.xml";
                string t1 = @destinot + "\\" + namefile;
                oDelivery.UserFields.Fields.Item("U_GMI_V1XML").Value = t1;

                // 🔹 Actualizar la Entrega con el adjunto
                int resultadoEntrega = oDelivery.Update();
                if (resultadoEntrega == 0)
                {
                    Log.Information("✅ Adjunto XML asociado correctamente a la Entrega.");
                    val = true;
                }
                else
                {
                    Log.Error($"❌ Error al actualizar la Entrega: {Conexion.oCompany.GetLastErrorDescription()}");
                    val = false;
                }

                //fin fin SAP

                //ini
                destinot = destinot + "\\" + namefile;
                // 📌 Verificar si el archivo ya existe en la carpeta destino
                if (File.Exists(destinot))
                {
                    File.Delete(destinot); // Eliminar el archivo para poder moverlo
                    Log.Information($"Archivo existente eliminado: {namefile}");
                }

                // 📌 Mover el archivo
                File.Move(archivot, destinot);
                Log.Information($"Archivo movido: {namefile}");
                //Console.WriteLine($"Archivo movido: {nombreArchivo}");

                //fin 


            }
            else {

                Log.Error("Nombre del archivo incompleto");
                val = false;
            }

            return val;
        }
    }
}
