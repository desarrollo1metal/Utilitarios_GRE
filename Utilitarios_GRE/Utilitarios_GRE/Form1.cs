
using System.IO.Compression;
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
using Ionic.Zip;
using Serilog.Core;
using System.Collections;
using System.Runtime.InteropServices;
using System.Xml;



namespace Utilitarios_GRE
{
    public partial class Form1 : Form
    {
        //static Company oCompany = new Company();
        static bool AbortarEnError = false;
        static bool CerrarAlFinalizar = false;
        public static clsConfig ConfiguracionGeneral;

        public string udfxmlFile    = string.Empty;
        public string udfcdr        = string.Empty;
        public string udfpdfsunat   = string.Empty;


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

        }


        static void InitializeCompany(SociedadBD sociedad)
        {
            
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

            udfxmlFile  = "U_GMI_V1XML";
            udfcdr      = "U_GMI_V1CDR";
            udfpdfsunat = "U_GRE_V1PDFS";

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
                    Log.Information("-------------------------------------------------");
                    Log.Information("Inicio el proceso de la sociedad " + sociedad.DbName);
                    //logger.Info("Conectado satisfactoriamente");
                    Conexion.InicializarVarGlob();
                    

                    //INI XML de Documento firmado
                    if (sociedad.PathFirmaOrigen != string.Empty)
                    {
                        procesarXML(sociedad.PathFirmaOrigen, sociedad.PathProcesadoFirma ,sociedad.PathFirmaError );
                    }

                    if (sociedad.PathFirmaOrigenIp3 != string.Empty)
                    {
                        procesarXML(sociedad.PathFirmaOrigenIp3, sociedad.PathProcesadoFirma, sociedad.PathFirmaError);
                    }
                    //FIN XML de Documento firmado


                    //INI CDR de respuesta
                    if (sociedad.PathCdrZip != string.Empty)
                    {
                        procesarCdr(sociedad.PathCdrZip, sociedad.PathCdrProcesado, sociedad.PathCdrError);
                    }


                    //FIN CDR de respuesta



                    Log.Information("Finalizando el Proceso de la sociedad " + sociedad.DbName);

                    Conexion.oCompany.Disconnect();
                    Log.Information("Desconectando de la BD " + sociedad.DbName);

                    Log.Information("-------------------------------------------------");
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                Conexion.DestroyCompany();
            }

        }


        public void procesarXML(string origenT, string destinoT, string PathFirmaError )
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

                // 📌 Asegurar que la carpeta destino de errores existe
                if (!Directory.Exists(PathFirmaError))
                {
                    Directory.CreateDirectory(PathFirmaError);
                    Log.Information($"Carpeta error creada: {PathFirmaError}");
                }

                // 📌 Obtener todos los archivos XML en la carpeta origen
                string[] archivos = Directory.GetFiles(rutaOrigen, "*.xml");

                Log.Information($"Carpeta Origen: {rutaOrigen}");
                Log.Information($"Carpeta Destino: {rutaDestino}");


                Log.Information("Inicio de mover xml " + destinoT);

                foreach (string archivo in archivos)
                {
                    string nombreArchivo = Path.GetFileName(archivo); // Obtener solo el nombre

                    try
                    {
                        
                        string destino = Path.Combine(rutaDestino, nombreArchivo); // Ruta completa en destino

                        Log.Information("----------------------- Inicio " + nombreArchivo + "-----------------------");
                        //ini procesar en SAP , adjuntar
                        if (procesarXMLAdjuntoSAP(archivo, nombreArchivo, origenT, destinoT , PathFirmaError))
                        {

                            Log.Information($"Termino exito de proceso , GRE  {nombreArchivo}");
                        }
                        else
                        {
                            Log.Error($"Termino con error de proceso , GRE  {nombreArchivo}");
                        }
                        //fin procesar en SAP , adjuntar


                    }
                    catch (Exception ex)
                    {

                        Log.Error(ex, ex.ToString());
                    }
                    finally
                    {
                        Log.Information("----------------------- Fin " + nombreArchivo + "-----------------------");
                    }

                }

                //Log.Information("✅ Proceso completado  de BD " + );
            }
            catch (Exception ex)
            {
                Log.Error($"❌ Error: {ex.Message}");
            }


        }

        public void procesarCdr(string origenT, string destinoT, string PathFirmaError)
        {

            string rutaOrigen                   = @origenT;
            string rutaDestino                  = @destinoT;
            string fullFileNamecrd              = string.Empty;
            string DocumentDescriptionpdfSunat  = string.Empty;

            try
            {
                // 📌 Asegurar que la carpeta destino existe
                if (!Directory.Exists(rutaDestino))
                {
                    Directory.CreateDirectory(rutaDestino);
                    Log.Information($"Carpeta creada CDR : {rutaDestino}");
                }

                // 📌 Asegurar que la carpeta destino de errores existe
                if (!Directory.Exists(PathFirmaError))
                {
                    Directory.CreateDirectory(PathFirmaError);
                    Log.Information($"Carpeta error creada CDR : {PathFirmaError}");
                }

                // 📌 Obtener todos los archivos CDR.ZIP en la carpeta origen
                string[] archivos = Directory.GetFiles(rutaOrigen, "*.zip");

                Log.Information($"Carpeta Origen: {rutaOrigen}");
                Log.Information($"Carpeta Destino: {rutaDestino}");

                Log.Information("----------------------- Inicio -----------------------");
                Log.Information("Inicio de mover xml " + destinoT);


                string rutaA = @rutaOrigen; // Ruta origen
                string rutaB = @rutaDestino; // Ruta destino

                try
                {
                    // Obtener todos los archivos ZIP en la ruta A
                    string[] archivosZip = Directory.GetFiles(rutaA, "*.zip");

                    foreach (string archivo in archivosZip)
                    {
                        string nombreArchivo = Path.GetFileName(archivo);
                        string zipDestino = Path.Combine(rutaB, nombreArchivo);

                        if (File.Exists(zipDestino))
                        {
                            Log.Information("Eliminado la carpeta con nombre " + zipDestino.ToString());
                            File.Delete(zipDestino); // Eliminar la copia existente
                        }

                        string t1 = zipDestino.Replace(".zip", "");
                        if (Directory.Exists(t1))
                        {
                            Log.Information("Eliminado el archivo .zip con nombre " + t1);
                            Directory.Delete(t1, true); // Eliminar la copia existente
                        }

                        // Mover el archivo ZIP de A a B
                        File.Move(archivo, zipDestino);
                        Log.Information($"Archivo {nombreArchivo} movido exitosamente.");

                        // Crear carpeta de extracción
                        string carpetaExtraccion = Path.Combine(rutaB, Path.GetFileNameWithoutExtension(nombreArchivo));
                        Directory.CreateDirectory(carpetaExtraccion);

                        // Descomprimir el ZIP en la carpeta de extracción
                        using (ZipFile zip = ZipFile.Read(zipDestino))
                        {
                            foreach (ZipEntry entry in zip)
                            {
                                entry.Extract(carpetaExtraccion, ExtractExistingFileAction.OverwriteSilently);
                                Log.Information($"Archivo {entry.FileName} extraído en: {carpetaExtraccion}");

                                // Si el archivo extraído es un XML, leer su contenido y obtener el nodo DocumentDescription
                                if (Path.GetExtension(entry.FileName).ToLower() == ".xml")
                                {

                                    string xmlPath = Path.Combine(carpetaExtraccion, entry.FileName);

                                    XmlDocument xmlDoc = new XmlDocument();
                                    xmlDoc.Load(xmlPath);

                                    // Crear un NamespaceManager para manejar los namespaces
                                    XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                                    nsmgr.AddNamespace("cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                                    nsmgr.AddNamespace("cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");

                                    // Seleccionar el nodo <cbc:DocumentDescription>
                                    XmlNode documentDescriptionNode = xmlDoc.SelectSingleNode("//cbc:DocumentDescription", nsmgr);


                                    if (documentDescriptionNode != null)
                                    {
                                        DocumentDescriptionpdfSunat = documentDescriptionNode.InnerText;
                                        fullFileNamecrd = entry.FileName;

                                        Log.Information ($"Contenido de DocumentDescription en {entry.FileName}: {documentDescriptionNode.InnerText}");


                                        try
                                        {

                                            //string destino = Path.Combine(rutaDestino, nombreArchivo); // Ruta completa en destino

                                            Log.Information("----------------------- Inicio " + nombreArchivo + "-----------------------");

                                            //ini procesar en SAP , adjuntar
                                            //if (procesarCDRAdjuntoSAP(archivo, nombreArchivo, origenT, destinoT, PathFirmaError))
                                            string valorFinal = destinoT + "\\" + nombreArchivo.Replace(".zip", "") + "\\" + fullFileNamecrd;
                                            if (guardarxmlCdr(valorFinal, fullFileNamecrd.Replace(".xml", ""), DocumentDescriptionpdfSunat,zipDestino))
                                            {
                                                Log.Information($" ✅ Termino exitoso de proceso ,CDR-GRE  {nombreArchivo}");
                                            }
                                            else
                                            {
                                                Log.Error($" ❌ Termino con error de proceso , CDR-GRE  {nombreArchivo}");
                                            }
                                            //fin procesar en SAP , adjuntar


                                        }
                                        catch (Exception ex)
                                        {

                                            Log.Error(ex, ex.ToString());
                                        }
                                        finally
                                        {
                                            Log.Information("----------------------- Fin " + nombreArchivo + "-----------------------");
                                        }

                                    }
                                    else
                                    {
                                        Log.Error ($"El nodo DocumentDescription no se encontró en {entry.FileName}");
                                    }
                                }
                            }
                        }




                    }
                }
                catch (Exception ex)
                {
                    Log.Error(" ❌ Error: " + ex.Message);
                }

            }
            catch (Exception ex)
            {
                Log.Error($"❌ Error: {ex.Message}");
            }


        }


        private bool procesarXMLAdjuntoSAP(string archivot,string namefile ,string oringT ,string destinot ,string PathFirmaErrort) {

            bool val = false;
            
            

            try
            {
                string name = namefile;
                string[] dato = namefile.Split('-');

                string ruc = dato[0];
                string tipoDoc = dato[1];
                string serieDoc = dato[2];
                string correlativo = dato[3];
                string docentry = string.Empty;
                string docTypeSerch = string.Empty;

                Log.Information($"Ruc = {ruc} : tipoDoc = {tipoDoc} : serieDoc = {serieDoc} : correlativo = {correlativo}");

                if (dato.Length == 4)
                {
                    string t1 = @destinot + "\\" + namefile;
                    //ini adjuntar SAP
                    Recordset oRecordSet = null;
                    oRecordSet = (Recordset)Conexion.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    Log.Information("Buscando Entrada en SAP");
                    string query = string.Empty;

                    query = $@"SELECT DocEntry FROM ODLN WHERE U_BPP_MDTD = '09' AND U_BPP_MDSD = '{serieDoc}' AND U_BPP_MDCD = '{correlativo.Replace(".xml", "")}' AND CANCELED <> 'Y'";

                    //logger.Debug(query);
                    oRecordSet.DoQuery(query);

                    if (oRecordSet.RecordCount > 0)
                    {
                        while (!oRecordSet.EoF)
                        {
                            docentry = (oRecordSet.Fields.Item("DocEntry").Value.ToString());
                            docTypeSerch = "ODLN";
                            oRecordSet.MoveNext();
                        }
                        Log.Information("Se encontro el documento ODLN (Entrega) en SAP con DocEntry = " + docentry);
                        val = true;
                    }
                    else
                    {


                        query = string.Empty;
                        query = $@"SELECT DocEntry FROM OWTR WHERE U_BPP_MDTD = '09' AND U_BPP_MDSD = '{serieDoc}' AND U_BPP_MDCD = '{correlativo.Replace(".xml", "")}' AND CANCELED <> 'Y'";

                        oRecordSet.DoQuery(query);

                        if (oRecordSet.RecordCount > 0)
                        {
                            while (!oRecordSet.EoF)
                            {
                                docentry = (oRecordSet.Fields.Item("DocEntry").Value.ToString());
                                docTypeSerch = "OWTR";
                                oRecordSet.MoveNext();
                            }
                            Log.Information("Se encontro el documento OWTR (Transferencia de Stock) en SAP con DocEntry = " + docentry);
                            val = true;
                        }
                        else
                        {

                            Log.Error("No se encontraron en (Entrada) ni (Transferencia de Stock)");
                            val = false;
                        }

                    }

                    switch (docTypeSerch)
                    {
                        case "ODLN":

                            // 🔹 Obtener la Entrega de Mercancía (ODLN)
                            Documents oDelivery = (Documents)Conexion.oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                            if (!oDelivery.GetByKey(Convert.ToInt32(docentry.Trim())))
                            {
                                Log.Information("❌ No se encontró la Entrega de Mercancía.");
                            }

                            // 🔹 Asignar el adjunto a la entrega
                            oDelivery.UserFields.Fields.Item("udfxmlFile").Value = t1;

                            // 🔹 Actualizar la Entrega con el adjunto
                            int resultadoEntrega = oDelivery.Update();
                            if (resultadoEntrega == 0)
                            {
                                Log.Information("✅ Se adjunto satisfactoriamente, XML a Entrega");
                                val = true;
                            }
                            else
                            {
                                Log.Error($"❌ Error SAP [U]: {Conexion.oCompany.GetLastErrorDescription()}");
                                val = false;
                            }

                            break;

                        case "OWTR":

                            // 🔹 Obtener la Entrega de Mercancía (ODLN)
                            StockTransfer oTransferencia = null;

                            oTransferencia = (StockTransfer)Conexion.oCompany.GetBusinessObject(BoObjectTypes.oStockTransfer);
                            if (!oTransferencia.GetByKey(Convert.ToInt32(docentry.Trim())))
                            {
                                Log.Information("❌ No se encontró la Transferencia de Stock con DocEntry: " + docentry);
                            }

                            // 🔹 Asignar el adjunto a la Transferencia de Stock
                            oTransferencia.UserFields.Fields.Item(udfxmlFile).Value = t1;

                            // 🔹 Actualizar la Entrega con el adjunto
                            int resultadoEntrega1 = oTransferencia.Update();
                            if (resultadoEntrega1 == 0)
                            {
                                Log.Information("✅ Se adjunto satisfactoriamente, XML a Transferencia de Stock");
                                val = true;
                            }
                            else
                            {
                                Log.Error($"❌ Error SAP [U] : {Conexion.oCompany.GetLastErrorDescription()}");
                                val = false;
                            }
                            break;

                        default:
                            break;
                    }

                    if (val)
                    {
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
                        Log.Information($"Archivo movido a  : {destinot} ");

                        //destinot

                        //fin 
                    }

                    if (!val)
                    {
                        //ini
                        PathFirmaErrort = PathFirmaErrort + "\\" + namefile;
                        // 📌 Verificar si el archivo ya existe en la carpeta destino
                        if (File.Exists(PathFirmaErrort))
                        {
                            File.Delete(PathFirmaErrort); // Eliminar el archivo para poder moverlo
                            Log.Information($"Archivo existente eliminado: {namefile}");
                        }

                        // 📌 Mover el archivo
                        File.Move(archivot, PathFirmaErrort);
                        Log.Information($"Archivo movido a  : {PathFirmaErrort} ");

                        //destinot

                        //fin 
                    }

                }
                else
                {

                    Log.Error("Nombre del archivo incompleto");
                    val = false;
                }


            }
            catch (Exception ex1)
            {

                Log.Error(ex1, ex1.Message.ToString());
                val=false; 
            }

            return val;

        }

        
        private bool guardarxmlCdr(string valorfinal, string namefile , string valorPdfSunatLink , string zipDestino)
        {

            bool val = false;

            try
            {
                string name = namefile;
                string[] dato = namefile.Split('-');

                string ruc = dato[1];
                string tipoDoc = dato[2];
                string serieDoc = dato[3];
                string correlativo = dato[4];
                string docentry = string.Empty;
                string docTypeSerch = string.Empty;

                Log.Information($"Ruc = {ruc} : tipoDoc = {tipoDoc} : serieDoc = {serieDoc} : correlativo = {correlativo}");

                if (dato.Length == 5)
                {
                    string t1 = valorfinal;
                    //ini adjuntar SAP
                    Recordset oRecordSet = null;
                    oRecordSet = (Recordset)Conexion.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    Log.Information("Buscando Entrada en SAP");
                    string query = string.Empty;

                    query = $@"SELECT DocEntry FROM ODLN WHERE U_BPP_MDTD = '09' AND U_BPP_MDSD = '{serieDoc}' AND U_BPP_MDCD = '{correlativo.Replace(".xml", "")}' AND CANCELED <> 'Y'";

                    //logger.Debug(query);
                    oRecordSet.DoQuery(query);

                    if (oRecordSet.RecordCount > 0)
                    {
                        while (!oRecordSet.EoF)
                        {
                            docentry = (oRecordSet.Fields.Item("DocEntry").Value.ToString());
                            docTypeSerch = "ODLN";
                            oRecordSet.MoveNext();
                        }
                        Log.Information("Se encontro el documento ODLN (Entrega) en SAP con DocEntry = " + docentry);
                        val = true;
                    }
                    else
                    {


                        query = string.Empty;
                        query = $@"SELECT DocEntry FROM OWTR WHERE U_BPP_MDTD = '09' AND U_BPP_MDSD = '{serieDoc}' AND U_BPP_MDCD = '{correlativo.Replace(".xml", "")}' AND CANCELED <> 'Y'";

                        oRecordSet.DoQuery(query);

                        if (oRecordSet.RecordCount > 0)
                        {
                            while (!oRecordSet.EoF)
                            {
                                docentry = (oRecordSet.Fields.Item("DocEntry").Value.ToString());
                                docTypeSerch = "OWTR";
                                oRecordSet.MoveNext();
                            }
                            Log.Information("Se encontro el documento OWTR (Transferencia de Stock) en SAP con DocEntry = " + docentry);
                            val = true;
                        }
                        else
                        {

                            Log.Error("No se encontraron en (Entrada) ni (Transferencia de Stock)");
                            val = false;
                        }

                    }

                    switch (docTypeSerch)
                    {
                        case "ODLN":

                            // 🔹 Obtener la Entrega de Mercancía (ODLN)
                            Documents oDelivery = (Documents)Conexion.oCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                            if (!oDelivery.GetByKey(Convert.ToInt32(docentry.Trim())))
                            {
                                Log.Information("❌ No se encontró la Entrega de Mercancía.");
                            }

                            // 🔹 Asignar el adjunto a la entrega
                            oDelivery.UserFields.Fields.Item(udfcdr).Value = t1;
                            //oDelivery.UserFields.Fields.Item("U_GMI_V1PDFS").Value = valorPdfSunatLink;
                            oDelivery.UserFields.Fields.Item(udfpdfsunat).Value = valorPdfSunatLink;



                            // 🔹 Actualizar la Entrega con el adjunto
                            int resultadoEntrega = oDelivery.Update();
                            if (resultadoEntrega == 0)
                            {
                                Log.Information("✅ Se adjunto satisfactoriamente, CDR-XML a Entrega");
                                val = true;
                            }
                            else
                            {
                                Log.Error($"❌ Error SAP [U]: {Conexion.oCompany.GetLastErrorDescription()}");

                                //// ini delete

                                //if (File.Exists(zipDestino))
                                //{
                                //    Log.Information("Eliminado la carpeta con nombre " + zipDestino.ToString());
                                //    File.Delete(zipDestino); // Eliminar la copia existente
                                //}

                                //string t2 = zipDestino.Replace(".zip", "");
                                //if (Directory.Exists(t2))
                                //{
                                //    Log.Information("Eliminado el archivo .zip con nombre " + t2);
                                //    Directory.Delete(t2, true); // Eliminar la copia existente
                                //}

                                //// Mover el archivo ZIP de A a B
                                //File.Move(t1, zipDestino);
                                //Log.Information($"Archivo {t1} movido exitosamente.");
                                //// fin delete

                                val = false;
                            }

                            break;

                        case "OWTR":

                            // 🔹 Obtener la Entrega de Mercancía (ODLN)
                            StockTransfer oTransferencia = null;

                            oTransferencia = (StockTransfer)Conexion.oCompany.GetBusinessObject(BoObjectTypes.oStockTransfer);
                            if (!oTransferencia.GetByKey(Convert.ToInt32(docentry.Trim())))
                            {
                                Log.Information("❌ No se encontró la Transferencia de Stock con DocEntry: " + docentry);
                            }

                            // 🔹 Asignar el adjunto a la Transferencia de Stock
                            oTransferencia.UserFields.Fields.Item(udfcdr).Value = t1;
                            oTransferencia.UserFields.Fields.Item(udfpdfsunat).Value = valorPdfSunatLink;


                            // 🔹 Actualizar la Entrega con el adjunto
                            int resultadoEntrega1 = oTransferencia.Update();
                            if (resultadoEntrega1 == 0)
                            {
                                Log.Information("✅ Se adjunto satisfactoriamente, CDR-XML a Transferencia de Stock");
                                val = true;
                            }
                            else
                            {
                                Log.Error($"❌ Error SAP [U] : {Conexion.oCompany.GetLastErrorDescription()}");

                                ////ini delete

                                //if (File.Exists(zipDestino))
                                //{
                                //    Log.Information("Eliminado la carpeta con nombre " + zipDestino.ToString());
                                //    File.Delete(zipDestino); // Eliminar la copia existente
                                //}

                                //string t2 = zipDestino.Replace(".zip", "");
                                //if (Directory.Exists(t2))
                                //{
                                //    Log.Information("Eliminado el archivo .zip con nombre " + t2);
                                //    Directory.Delete(t2, true); // Eliminar la copia existente
                                //}

                                //// Mover el archivo ZIP de A a B
                                //File.Move(t1, zipDestino);
                                //Log.Information($"Archivo {t1} movido exitosamente.");
                                //// fin delete

                                val = false;
                            }
                            break;

                        default:
                            break;
                    }

                    ////if (val)
                    ////{
                    ////    //ini
                    ////    destinot = destinot + "\\" + namefile;
                    ////    // 📌 Verificar si el archivo ya existe en la carpeta destino
                    ////    if (File.Exists(destinot))
                    ////    {
                    ////        File.Delete(destinot); // Eliminar el archivo para poder moverlo
                    ////        Log.Information($"Archivo existente eliminado: {namefile}");
                    ////    }

                    ////    // 📌 Mover el archivo
                    ////    File.Move(archivot, destinot);
                    ////    Log.Information($"Archivo movido a  : {destinot} ");

                    ////    //destinot

                    ////    //fin 
                    ////}

                    ////if (!val)
                    ////{
                    ////    //ini
                    ////    PathFirmaErrort = PathFirmaErrort + "\\" + namefile;
                    ////    // 📌 Verificar si el archivo ya existe en la carpeta destino
                    ////    if (File.Exists(PathFirmaErrort))
                    ////    {
                    ////        File.Delete(PathFirmaErrort); // Eliminar el archivo para poder moverlo
                    ////        Log.Information($"Archivo existente eliminado: {namefile}");
                    ////    }

                    ////    // 📌 Mover el archivo
                    ////    File.Move(archivot, PathFirmaErrort);
                    ////    Log.Information($"Archivo movido a  : {PathFirmaErrort} ");

                    ////    //destinot

                    ////    //fin 
                    ////}

                }
                else
                {

                    Log.Error("[WS] Nombre del archivo incompleto ");
                    val = false;
                }


            }
            catch (Exception ex1)
            {

                Log.Error(ex1, ex1.Message.ToString());
                val = false;
            }

            return val;

        }



    }
}
