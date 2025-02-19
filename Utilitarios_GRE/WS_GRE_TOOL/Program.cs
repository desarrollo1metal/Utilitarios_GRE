using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using log4net.Appender;
using log4net.Layout;
using log4net.Repository.Hierarchy;
using SAPbobsCOM;
using Ionic.Zip;
using System.Data.SqlClient;


namespace WS_GRE_TOOL
{
    public class Program
    {
        
        string versiónTools = string.Empty;
        static bool AbortarEnError = false;
        static bool CerrarAlFinalizar = false;
        public static clsConfig ConfiguracionGeneral;
        private static bool Actualizar = false;

        public static string udfxmlFile = string.Empty;
        public static string udfcdr = string.Empty;
        public static string udfpdfsunat = string.Empty;
        static string versionNow = "1.2.5";
        string descripProd = "GMI SERVICIO_ST2";
        string codProd = "GMI_ST2";

        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(typeof(Program));
        [STAThread]
        public static void Main(string[] args)
        {
            logger.Debug("1nicio");
            Console.SetWindowSize(100, 20);
            logger.Debug("fin");

            Program programa = new Program();
            programa.Iniciar();

            Environment.Exit(0);

        }

        public  string ObtenerNumVersion()
        {
            
            string bdversion = string.Empty;
            string consultaSQL = $"SELECT U_GMI_VRSION FROM [dbo].[@GMI_VER1] WHERE U_GMI_CODPRD = '{codProd}'";

            Recordset oRecordSet3 = null;
            try
            {
                oRecordSet3= (Recordset)Conexion.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordSet3.DoQuery(consultaSQL);

                int filas = oRecordSet3.RecordCount;
                if (!oRecordSet3.EoF) bdversion = oRecordSet3.Fields.Item(0).Value.ToString();
            }
            catch { }
            finally
            {
                //VSSQLFactory.LiberarObjeto(oRecordSet);
                try
                {
                    if (oRecordSet3 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet3);
                }
                catch { }
                try
                {
                    oRecordSet3 = null;
                }
                catch { }
                GC.Collect();
            }



            return bdversion;
        }

        private void ValidarVersion()
        {

            logger.Info("iniciar validad versión ");

            string bdversion = string.Empty;
            string consultaSQL = $"SELECT U_GMI_VRSION FROM [dbo].[@GMI_VER1] WHERE U_GMI_CODPRD = '{codProd}'";
            logger.Info("Consulta SQL " + consultaSQL);
            //string query = "SELECT u_vs_vrsion from [dbo].[@vs_prd1] where u_vs_codprd='" + AppMain.codigoproducto + "'";
            //00001_MSSQL
            Recordset oRecordSet = null;
            int filas = 0;
            try
            {
                oRecordSet = (Recordset)Conexion.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                //string qry = UTIL.VSSQLFactory.GetScript(1, new string[] { AppMain.codigoproducto });
                oRecordSet.DoQuery(consultaSQL); //00001
                filas = oRecordSet.RecordCount;
                if (!oRecordSet.EoF) bdversion = oRecordSet.Fields.Item(0).Value.ToString();
            }
            catch
            {
                bdversion = null;
            }
            finally
            {
                if (oRecordSet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }

            try
            {
                //#if DEBUG
                //                return;
                //#endif
                if (bdversion == null || bdversion == "")
                {
                    //insert nuevo de cero

                    //logger.Info("Instalando addon");
                    //AppMain.MensajeAdvertencia("Instalando addon");
                    //CargarEstructura();
                    if (filas == 0)
                    {
                       

                        SAPbobsCOM.Recordset oRecordSet4 = null;
                        try
                        {
                            //String query = "insert into [dbo].[@vs_prd1](Code,LineId,Object,LogInst,U_VS_CODPRD,U_VS_DSCPRD,U_VS_VRSION) values('VS',(SELECT COUNT(0) FROM [dbo].[@vs_prd1])+1,'BPVS_OPRD',NULL,'" + AppMain.codigoproducto+ "','"+AppMain.descripcionproducto+"','" +AppMain.versionAddon+ "')";
                            string query = "insert into [dbo].[@GMI_VER1](Code,LineId,Object,LogInst,U_GMI_CODPRD,U_GMI_DSCPRD,U_GMI_VRSION) values('GMI',(SELECT COUNT(0) FROM [dbo].[@GMI_VER1])+1,'BGMI_TVER',NULL,'" + codProd + "','" + descripProd + "','" + versionNow + "')";
                            oRecordSet4 = (Recordset)Conexion.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            oRecordSet4.DoQuery(query);

                            logger.Debug("Se registra nuevo producto codProd  = " + codProd + " - descripProd  = " + descripProd + " - versionNow = " + versionNow);
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex.Message, ex);
                        }
                        finally
                        {
                            if (oRecordSet4 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet4);
                            oRecordSet4 = null;
                            GC.Collect();
                        }
                    }
                    //    ProductoBL.RegistrarProducto();
                    //else
                    //    ProductoBL.ActualizarVersion();
                    ////AppMain.MensajeError("Addon instalado correctamente, reinicie SAP Business One");
                }
                else
                {
                    if (VersionMenor(versionNow, bdversion))
                    {
                        logger.Info("Actualizando version del addon");
                        Actualizar = true;
                        //AppMain.MensajeAdvertencia("Actualizando version del addon");
                        CargarEstructura();

                        //ProductoBL.ActualizarVersion();
                        //actualizar version
                        //string consultaSQLupdate = $"UPDATE U_GMI_VRSION FROM [dbo].[@GMI_VER1] WHERE U_GMI_CODPRD = '{producto}' AND ";

                        
                        string consultaSQLupdate = $@" UPDATE [@GMI_VER1]  SET U_GMI_VRSION = '{versionNow}'  WHERE U_GMI_CODPRD = '{codProd}';";

                        //ini actualizar

                        SAPbobsCOM.Recordset oRecordSet1 = null;
                        try
                        {
                            //String query = "update [dbo].[@vs_prd1] set u_vs_vrsion='" + AppMain.versionAddon + "' where u_vs_codprd='" + AppMain.codigoproducto + "'";
                            oRecordSet1 = (Recordset)Conexion.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            oRecordSet1.DoQuery(consultaSQLupdate);

                            logger.Debug("Consulta SQL " + consultaSQLupdate);
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex.Message, ex);
                        }
                        finally
                        {
                            if (oRecordSet1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet1);
                            oRecordSet1 = null;
                            GC.Collect();
                        }

                        //fin actualizar

                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
            finally
            {
                if (oRecordSet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }


        public  void Iniciar()
        {
            SetUpLogger();

            

            //fin

            //logger.Debug("Inicio del sercio");

            //log4net.Config.XmlConfigurator.Configure();
            try
            {
                //Console.Title = "TOOL - GMI";
                //logger.Debug("TOOL - GMI");
                //logger.Error("probando Errror");
                System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();

                logger.Info("---------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                logger.Info("Inicializando Tool GMI ");
                stopwatch.Start();

                udfxmlFile = "U_XMLV3";
                udfcdr = "U_CDRV3";
                udfpdfsunat = "U_SUNATPDFV3";


                if (!LeerConfig()) return;



                if (Conexion.ConfiguracionGeneral.Sociedades.Length == 0)
                {
                    logger.Error("No se encontraron BD en el archivo de configuración, por favor revise");
                    return;
                }

                logger.Debug("BD encontradas en archivo son " + Conexion.ConfiguracionGeneral.Sociedades.Length.ToString());


                foreach (SociedadBD sociedad in Conexion.ConfiguracionGeneral.Sociedades)
                {
                    string msj;


                    // Verificar si las propiedades PathFirmaOrigen y PathFirmaOrigenIp3 existen o no son null
                    string pathFirmaOrigen = sociedad.PathFirmaOrigen ?? string.Empty;
                    string pathFirmaOrigenIp3 = sociedad.PathFirmaOrigenIp3 ?? string.Empty;

                    if (string.IsNullOrEmpty(pathFirmaOrigen) || string.IsNullOrEmpty(pathFirmaOrigenIp3))
                    {
                        logger.Error($"Uno o ambos paths no están configurados para la sociedad: {sociedad.DbName}");
                        continue; // Saltar a la siguiente sociedad si los paths no están configurados
                    }

                    // Verificar si los directorios existen y contienen archivos
                    bool existenArchivosEnFirmaOrigen = Directory.Exists(pathFirmaOrigen) && Directory.GetFiles(pathFirmaOrigen).Length > 0;
                    bool existenArchivosEnFirmaOrigenIp3 = Directory.Exists(pathFirmaOrigenIp3) && Directory.GetFiles(pathFirmaOrigenIp3).Length > 0;

                    if (!existenArchivosEnFirmaOrigen && !existenArchivosEnFirmaOrigenIp3)
                    {
                        logger.Error($"No se encontraron archivos en los directorios: {pathFirmaOrigen} y {pathFirmaOrigenIp3}");
                        continue; // Saltar a la siguiente sociedad si no hay archivos
                    }

                    logger.Debug("Se encontro archivos pendiente de procesar...");

                    // Si hay archivos en al menos uno de los directorios, proceder con la conexión
                    Conexion.InitializeCompany(sociedad);
                    Conexion.oCompany.Connect();

                    if (Conexion.oCompany.Connected == false)
                    {
                        int rpta = 0;
                        Conexion.oCompany.GetLastError(out rpta, out msj);
                        logger.Error(rpta.ToString() + " -- " + msj.ToString());
                        //logger.Error(rpta.ToString() + " -- " + msj.ToString());
                    }
                    else
                    {
                        logger.Debug("--------------------------------------------------------------------------------------------------");
                        logger.Debug("Conectado satisfactoriamente a BD " + sociedad.DbName + " con DI API");
                        //crear campos
                        //recien  validamos versiones

                        //INI   version
                        ValidarVersion();
                        //FIN   version


                        Conexion.InicializarVarGlob();


                        //INI XML de Documento firmado
                        logger.Debug("************************** INI del proceso de XML FILE **************************");
                        if (sociedad.PathFirmaOrigen != string.Empty)
                        {
                            //logger.Debug("Sociedad.PathFirmaOrigen ");
                            procesarXML(sociedad.PathFirmaOrigen, sociedad.PathProcesadoFirma, sociedad.PathFirmaError);
                        }

                        if (sociedad.PathFirmaOrigenIp3 != string.Empty)
                        {
                            //logger.Debug("Sociedad.PathFirmaOrigenIp3 ");
                            procesarXML(sociedad.PathFirmaOrigenIp3, sociedad.PathProcesadoFirma, sociedad.PathFirmaError);
                        }
                        logger.Debug("************************** FIN del proceso de XML FILE **************************");
                        //FIN XML de Documento firmado


                        //INI CDR de respuesta
                        logger.Debug("------------------------- INI del proceso de  CDR       -------------------------");
                        if (sociedad.PathCdrZip != string.Empty)
                        {
                            //logger.Debug("Sociedad.PathCdrZip 1");
                            procesarCdr(sociedad.PathCdrZip, sociedad.PathCdrProcesado, sociedad.PathCdrError);
                        }

                        if (sociedad.PathCdrZipIp3 != string.Empty)
                        {
                            //logger.Debug("Sociedad.PathCdrZip 3");
                            procesarCdr(sociedad.PathCdrZipIp3, sociedad.PathCdrProcesado, sociedad.PathCdrError);
                        }
                        logger.Debug("------------------------- FIN del proceso de  CDR       -------------------------");
                        //FIN CDR de respuesta

                        logger.Debug("Finalizando el Proceso de la sociedad(BD)  " + sociedad.DbName);
                        Conexion.oCompany.Disconnect();
                        logger.Debug("Desconectado el Proceso de la sociedad(BD) " + sociedad.DbName);

                        //logger.Debug("-------------------------------------------------");
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                    Conexion.DestroyCompany();
                }


                stopwatch.Stop();
                logger.Info("Tiempo total de ejecución: " + stopwatch.Elapsed.ToString());
                logger.Info("---------------------------------------------------------------------------------------------------------------------------------------------------------------------");

                //Util.FileMGMT.EliminarTemporales(AppDomain.CurrentDomain.BaseDirectory);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }

        static bool LeerConfig()
        {
            try
            {
                string ruta = AppDomain.CurrentDomain.BaseDirectory + "Config.xml";
                if (!File.Exists(ruta))
                {
                    logger.Error("No se encontró el archivo de configuración Config.xml en la ruta " + ruta.ToString());
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
                logger.Error( ex.ToString());
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        public static void procesarXML(string origenT, string destinoT, string PathFirmaError)
        {

            string rutaOrigen = @origenT;
            string rutaDestino = @destinoT;

            try
            {
                // 📌 Asegurar que la carpeta destino existe
                if (!Directory.Exists(rutaDestino))
                {
                    Directory.CreateDirectory(rutaDestino);
                    logger.Debug($"Carpeta creada: {rutaDestino}");
                }

                // 📌 Asegurar que la carpeta destino de errores existe
                if (!Directory.Exists(PathFirmaError))
                {
                    Directory.CreateDirectory(PathFirmaError);
                    logger.Debug($"Carpeta error creada: {PathFirmaError}");
                }

                // 📌 Obtener todos los archivos XML en la carpeta origen
                string[] archivos = Directory.GetFiles(rutaOrigen, "*.xml");

                logger.Debug($"Carpeta Origen: {rutaOrigen}");
                logger.Debug($"Carpeta Destino: {rutaDestino}");


                logger.Debug("Inicio de mover xml " + destinoT);

                foreach (string archivo in archivos)
                {
                    string nombreArchivo = Path.GetFileName(archivo); // Obtener solo el nombre
                    int totalArchivos = archivos.Length;
                    logger.Debug("Numero de archivos XML " + totalArchivos);

                    try
                    {

                        string destino = Path.Combine(rutaDestino, nombreArchivo); // Ruta completa en destino

                        logger.Debug("----------------------- Inicio XML " + nombreArchivo + "-----------------------");
                        //ini procesar en SAP , adjuntar
                        if (procesarXMLAdjuntoSAP(archivo, nombreArchivo, origenT, destinoT, PathFirmaError))
                        {

                            logger.Debug($"Termino exito de proceso , GRE  {nombreArchivo}");
                        }
                        else
                        {
                            logger.Error($"Termino con error de proceso , GRE  {nombreArchivo}");
                        }
                        //fin procesar en SAP , adjuntar


                    }
                    catch (Exception ex)
                    {

                        logger.Error( ex.ToString());
                    }
                    finally
                    {
                        logger.Debug("----------------------- Fin XML " + nombreArchivo + "-----------------------");
                    }

                }

                //logger.Debug("✅ Proceso completado  de BD " + );
            }
            catch (Exception ex)
            {
                logger.Error($"❌ Error: {ex.Message}");
            }


        }

        public static  void procesarCdr(string origenT, string destinoT, string PathFirmaError)
        {

            string rutaOrigen = @origenT;
            string rutaDestino = @destinoT;
            string fullFileNamecrd = string.Empty;
            string DocumentDescriptionpdfSunat = string.Empty;
            bool existos = true;

            try
            {
                // 📌 Asegurar que la carpeta destino existe
                if (!Directory.Exists(rutaDestino))
                {
                    Directory.CreateDirectory(rutaDestino);
                    logger.Debug($"Carpeta creada CDR : {rutaDestino}");
                }

                // 📌 Asegurar que la carpeta destino de errores existe
                if (!Directory.Exists(PathFirmaError))
                {
                    Directory.CreateDirectory(PathFirmaError);
                    logger.Debug($"Carpeta error creada CDR : {PathFirmaError}");
                }

                // 📌 Obtener todos los archivos CDR.ZIP en la carpeta origen
                string[] archivos = Directory.GetFiles(rutaOrigen, "*.zip");

                logger.Debug($"Carpeta Origen: {rutaOrigen}");
                logger.Debug($"Carpeta Destino: {rutaDestino}");


                logger.Debug("Inicio de mover xml " + destinoT);


                string rutaA = @rutaOrigen; // Ruta origen
                string rutaB = @rutaDestino; // Ruta destino

                try
                {

                    try
                    {
                        // Obtener todos los archivos ZIP en la carpeta de origen
                        string[] archivosZip2 = Directory.GetFiles(rutaOrigen, "*.zip");

                        foreach (string archivoZip2 in archivosZip2)
                        {

                            string nombreArchivo = Path.GetFileName(archivoZip2);
                            string zipDestino = Path.Combine(rutaB, nombreArchivo);

                            logger.Debug("----------------------- Inicio " + nombreArchivo + "-----------------------");


                            if (File.Exists(zipDestino))
                            {
                                logger.Debug("Eliminado la carpeta con nombre " + zipDestino.ToString());
                                File.Delete(zipDestino); // Eliminar la copia existente
                            }

                            string t1 = zipDestino.Replace(".zip", "");
                            if (Directory.Exists(t1))
                            {
                                logger.Debug("Eliminado el archivo .zip con nombre " + t1);
                                Directory.Delete(t1, true); // Eliminar la copia existente
                            }

                            // Mover el archivo ZIP de A a B
                            File.Move(archivoZip2, zipDestino);
                            logger.Debug($"Archivo {nombreArchivo} movido exitosamente.");

                            // Crear carpeta de extracción
                            string carpetaExtraccion = Path.Combine(rutaB, Path.GetFileNameWithoutExtension(nombreArchivo));
                            Directory.CreateDirectory(carpetaExtraccion);

                            using (ZipFile zip = ZipFile.Read(zipDestino))
                            {
                                zip.ExtractAll(carpetaExtraccion, ExtractExistingFileAction.OverwriteSilently);
                                //Console.WriteLine($"Archivo ZIP {Path.GetFileName(archivoZip2)} descomprimido en: {carpetaExtraccion}");
                                logger.Debug($"Archivo {nombreArchivo} extraído en: {carpetaExtraccion}");

                                foreach (ZipEntry entry in zip)
                                {
                                    entry.Extract(carpetaExtraccion, ExtractExistingFileAction.OverwriteSilently);
                                    logger.Debug($"Archivo {entry.FileName} extraído en: {carpetaExtraccion}");

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

                                            logger.Debug($"Contenido de DocumentDescription en {entry.FileName}: {documentDescriptionNode.InnerText}");
                                        }

                                    }

                                }


                                //SAP UPDATE
                                try
                                {

                                    //ini procesar en SAP , adjuntar
                                    string valorFinal = destinoT + "\\" + nombreArchivo.Replace(".zip", "") + "\\" + fullFileNamecrd;
                                    if (guardarxmlCdr(valorFinal, fullFileNamecrd.Replace(".xml", ""), DocumentDescriptionpdfSunat))
                                    {
                                        logger.Debug($" ✅ Termino exitoso de proceso ,CDR-GRE  {nombreArchivo}");
                                    }
                                    else
                                    {
                                        logger.Error($" ❌ Termino con error de proceso , CDR-GRE  {nombreArchivo}");
                                        existos = false;
                                    }
                                    //fin procesar en SAP , adjuntar

                                }
                                catch (Exception ex)
                                {

                                    logger.Error( ex.ToString());
                                }
                                finally
                                {

                                }


                            } // Aquí se libera el archivo ZIP automáticamente

                            if (!existos)
                            {
                                string nuevoNombreArchivo1 = Path.Combine(PathFirmaError, Path.GetFileName(archivoZip2));

                                if (File.Exists(nuevoNombreArchivo1))
                                {

                                    File.Delete(nuevoNombreArchivo1); // Eliminar la copia existente
                                    logger.Debug("Archivo eliminado " + nuevoNombreArchivo1.ToString());
                                }

                                File.Move(zipDestino, nuevoNombreArchivo1);
                                logger.Debug($"Archivo ZIP  {Path.GetFileName(zipDestino)} movido a: {nuevoNombreArchivo1}");

                                zipDestino = zipDestino.Replace(".zip", "");
                                if (Directory.Exists(zipDestino))
                                {
                                    Directory.Delete(zipDestino, true);
                                    logger.Debug("Carpeta eliminada " + zipDestino.ToString());
                                    //File.Delete(zipDestino); // Eliminar la copia existente
                                }

                            }

                            existos = true;


                            logger.Debug("----------------------- Fin " + nombreArchivo + "-----------------------");
                        }
                    }
                    catch (Exception ex)
                    {

                        logger.Error($"Error: {ex.Message}");

                    }

                }
                catch (Exception ex)
                {
                    logger.Error(" ❌ Error: " + ex.Message);
                }

            }
            catch (Exception ex)
            {
                logger.Error($"❌ Error: {ex.Message}");
            }


        }

        private static bool procesarXMLAdjuntoSAP(string archivot, string namefile, string oringT, string destinot, string PathFirmaErrort)
        {

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

                logger.Debug($"Ruc = {ruc} : tipoDoc = {tipoDoc} : serieDoc = {serieDoc} : correlativo = {correlativo}");

                if (dato.Length == 4)
                {
                    string t1 = @destinot + "\\" + namefile;
                    //ini adjuntar SAP
                    Recordset oRecordSet = null;
                    oRecordSet = (Recordset)Conexion.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    logger.Debug("Buscando Entrada en SAP");
                    string query = string.Empty;

                    query = $@"SELECT DocEntry FROM ODLN WHERE U_BPP_MDTD = '09' AND U_BPP_MDSD = '{serieDoc}' AND U_BPP_MDCD = '{correlativo.Replace(".xml", "")}' AND CANCELED <> 'Y'";
                    logger.Debug("Query " + query);

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
                        logger.Debug("Se encontro el documento ODLN (Entrega) en SAP con DocEntry = " + docentry);
                        val = true;
                    }
                    else
                    {


                        query = string.Empty;
                        query = $@"SELECT DocEntry FROM OWTR WHERE U_BPP_MDTD = '09' AND U_BPP_MDSD = '{serieDoc}' AND U_BPP_MDCD = '{correlativo.Replace(".xml", "")}' AND CANCELED <> 'Y'";
                        logger.Debug("Query " + query);
                        oRecordSet.DoQuery(query);

                        if (oRecordSet.RecordCount > 0)
                        {
                            while (!oRecordSet.EoF)
                            {
                                docentry = (oRecordSet.Fields.Item("DocEntry").Value.ToString());
                                docTypeSerch = "OWTR";
                                oRecordSet.MoveNext();
                            }
                            logger.Debug("Se encontro el documento OWTR (Transferencia de Stock) en SAP con DocEntry = " + docentry);
                            val = true;
                        }
                        else
                        {

                            logger.Error("No se encontraron en (Entrada) ni (Transferencia de Stock)");
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
                                logger.Debug("❌ No se encontró la Entrega de Mercancía.");
                            }

                            // 🔹 Asignar el adjunto a la entrega
                            oDelivery.UserFields.Fields.Item(udfxmlFile).Value = t1;

                            // 🔹 Actualizar la Entrega con el adjunto
                            int resultadoEntrega = oDelivery.Update();
                            if (resultadoEntrega == 0)
                            {
                                logger.Debug("✅ Se adjunto satisfactoriamente, XML a Entrega");
                                val = true;
                            }
                            else
                            {
                                string erro1 = Conexion.oCompany.GetLastErrorDescription();
                                logger.Error($"❌ Error SAP [U]: {Conexion.oCompany.GetLastErrorDescription()}");
                                val = false;
                            }

                            break;

                        case "OWTR":

                            // 🔹 Obtener la Entrega de Mercancía (ODLN)
                            StockTransfer oTransferencia = null;

                            oTransferencia = (StockTransfer)Conexion.oCompany.GetBusinessObject(BoObjectTypes.oStockTransfer);
                            if (!oTransferencia.GetByKey(Convert.ToInt32(docentry.Trim())))
                            {
                                logger.Debug("❌ No se encontró la Transferencia de Stock con DocEntry: " + docentry);
                            }

                            // 🔹 Asignar el adjunto a la Transferencia de Stock
                            oTransferencia.UserFields.Fields.Item(udfxmlFile).Value = t1;

                            // 🔹 Actualizar la Entrega con el adjunto
                            int resultadoEntrega1 = oTransferencia.Update();
                            if (resultadoEntrega1 == 0)
                            {
                                logger.Debug("✅ Se adjunto satisfactoriamente, XML a Transferencia de Stock");
                                val = true;
                            }
                            else
                            {
                                string erro1 = Conexion.oCompany.GetLastErrorDescription();
                                logger.Error($"❌ Error SAP [U] : {Conexion.oCompany.GetLastErrorDescription()}");
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
                            logger.Debug($"Archivo existente eliminado: {namefile}");
                        }

                        // 📌 Mover el archivo
                        File.Move(archivot, destinot);
                        logger.Debug($"Archivo movido a  : {destinot} ");

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
                            logger.Debug($"Archivo existente eliminado: {namefile}");
                        }

                        // 📌 Mover el archivo
                        File.Move(archivot, PathFirmaErrort);
                        logger.Debug($"Archivo movido a  : {PathFirmaErrort} ");

                        //destinot

                        //fin 
                    }

                }
                else
                {

                    logger.Error("Nombre del archivo incompleto");
                    val = false;
                }


            }
            catch (Exception ex1)
            {

                logger.Error( ex1.Message.ToString());
                val = false;
            }

            return val;

        }

        private static bool guardarxmlCdr(string valorfinal, string namefile, string valorPdfSunatLink)
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

                logger.Debug($"Ruc = {ruc} : tipoDoc = {tipoDoc} : serieDoc = {serieDoc} : correlativo = {correlativo}");

                if (dato.Length == 5)
                {
                    string t1 = valorfinal;
                    //ini adjuntar SAP
                    Recordset oRecordSet = null;
                    oRecordSet = (Recordset)Conexion.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    logger.Debug("Buscando Entrada en SAP");
                    string query = string.Empty;

                    query = $@"SELECT DocEntry FROM ODLN WHERE U_BPP_MDTD = '09' AND U_BPP_MDSD = '{serieDoc}' AND U_BPP_MDCD = '{correlativo.Replace(".xml", "")}' AND CANCELED <> 'Y'";
                    logger.Debug("Query " + query);
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
                        logger.Debug("Se encontro el documento ODLN (Entrega) en SAP con DocEntry = " + docentry);
                        val = true;
                    }
                    else
                    {


                        query = string.Empty;
                        query = $@"SELECT DocEntry FROM OWTR WHERE U_BPP_MDTD = '09' AND U_BPP_MDSD = '{serieDoc}' AND U_BPP_MDCD = '{correlativo.Replace(".xml", "")}' AND CANCELED <> 'Y'";
                        logger.Debug("Query " + query);
                        oRecordSet.DoQuery(query);

                        if (oRecordSet.RecordCount > 0)
                        {
                            while (!oRecordSet.EoF)
                            {
                                docentry = (oRecordSet.Fields.Item("DocEntry").Value.ToString());
                                docTypeSerch = "OWTR";
                                oRecordSet.MoveNext();
                            }
                            logger.Debug("Se encontro el documento OWTR (Transferencia de Stock) en SAP con DocEntry = " + docentry);
                            val = true;
                        }
                        else
                        {

                            logger.Error("No se encontraron en (Entrada) ni (Transferencia de Stock)");
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
                                logger.Debug("❌ No se encontró la Entrega de Mercancía.");
                            }

                            // 🔹 Asignar el adjunto a la entrega
                            oDelivery.UserFields.Fields.Item(udfcdr).Value = t1;
                            //oDelivery.UserFields.Fields.Item("U_GMI_V1PDFS").Value = valorPdfSunatLink;
                            oDelivery.UserFields.Fields.Item(udfpdfsunat).Value = valorPdfSunatLink;



                            // 🔹 Actualizar la Entrega con el adjunto
                            int resultadoEntrega = oDelivery.Update();
                            if (resultadoEntrega == 0)
                            {
                                logger.Debug("✅ Se adjunto satisfactoriamente, CDR-XML a Entrega");
                                val = true;
                            }
                            else
                            {
                                string erro1 = Conexion.oCompany.GetLastErrorDescription();
                                logger.Error($"❌ Error SAP [U]: {Conexion.oCompany.GetLastErrorDescription()}");
                                val = false;
                            }

                            break;

                        case "OWTR":

                            // 🔹 Obtener la Entrega de Mercancía (ODLN)
                            StockTransfer oTransferencia = null;

                            oTransferencia = (StockTransfer)Conexion.oCompany.GetBusinessObject(BoObjectTypes.oStockTransfer);
                            if (!oTransferencia.GetByKey(Convert.ToInt32(docentry.Trim())))
                            {
                                logger.Debug("❌ No se encontró la Transferencia de Stock con DocEntry: " + docentry);
                            }

                            // 🔹 Asignar el adjunto a la Transferencia de Stock
                            oTransferencia.UserFields.Fields.Item(udfcdr).Value = t1;
                            oTransferencia.UserFields.Fields.Item(udfpdfsunat).Value = valorPdfSunatLink;


                            // 🔹 Actualizar la Entrega con el adjunto
                            int resultadoEntrega1 = oTransferencia.Update();
                            if (resultadoEntrega1 == 0)
                            {
                                logger.Debug("✅ Se adjunto satisfactoriamente, CDR-XML a Transferencia de Stock");
                                val = true;
                            }
                            else
                            {
                                string erro1 = Conexion.oCompany.GetLastErrorDescription();
                                logger.Error($"❌ Error SAP [U] : {Conexion.oCompany.GetLastErrorDescription()}");


                                val = false;
                            }
                            break;

                        default:
                            break;
                    }


                }
                else
                {

                    logger.Error("[WS] Nombre del archivo incompleto ");
                    val = false;
                }


            }
            catch (Exception ex1)
            {

                logger.Error(ex1.Message.ToString());
                val = false;
            }

            return val;

        }

        public static void CargarEstructura()
        {
            try
            {

                sapObjUser sapObj = new sapObjUser(Conexion.oCompany);

                //sapObj.CreaCampoMD("ODLN", "PRV1", "GRE Correlativo1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None,8, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, null);

                string udfxmlFilet  = udfxmlFile.Replace("U_","");
                string udfcdrt      = udfcdr.Replace("U_", "");
                string udfpdfsunatt = udfpdfsunat.Replace("U_", "");

                logger.Debug( "Campos crear " + udfxmlFilet + " - " +  udfcdrt +" - "+ udfpdfsunatt);

                sapObj.CreaCampoMD("ODLN", udfxmlFilet, "GRE XML", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 256000, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, null);
                sapObj.CreaCampoMD("ODLN", udfcdrt, "GRE CDR", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 256000, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, null);
                sapObj.CreaCampoMD("ODLN", udfpdfsunatt, "GRE PDF SUNAT", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 256000, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, null);

                //udfxmlFile = "U_GRE_XML";
                //udfcdr = "U_GRE_CDR";
                //udfpdfsunat = "U_GRE_SUNAT";


            }
            catch(Exception ex1)
            {
                logger.Error(ex1.Message.ToString());

            }

        }


        private static void SetUpLogger()
        {
            var ts = DateTime.Now.ToString("yyyy-MM-dd");
            log4net.GlobalContext.Properties["ts"] = ts;
            System.Collections.Specialized.NameValueCollection _settings = System.Configuration.ConfigurationManager.AppSettings;

            Hierarchy hierarchy = (Hierarchy)log4net.LogManager.GetRepository();

            PatternLayout patternLayout = new PatternLayout();
            patternLayout.ConversionPattern = "%-5p %d %5rms %-22.22c{1} %-30.30M - %m%n";
            patternLayout.ActivateOptions();

            PatternLayout patternLayoutConsole = new PatternLayout();
            patternLayout.ConversionPattern = "%date %-5level: %message%newline";
            patternLayout.ActivateOptions();

            RollingFileAppender roller = new RollingFileAppender();
            roller.AppendToFile = true;
            roller.File = @"Log\Mensajes_.log";
            roller.Layout = patternLayout;
            roller.MaxSizeRollBackups = 10;
            roller.MaximumFileSize = "10MB";
            roller.RollingStyle = RollingFileAppender.RollingMode.Date;
            roller.StaticLogFileName = false;
            roller.Threshold = log4net.Core.Level.All;
            roller.DatePattern = "yyyy-MM-dd";
            roller.PreserveLogFileNameExtension = true;
            roller.ActivateOptions();
            hierarchy.Root.AddAppender(roller);

            ConsoleAppender console = new ConsoleAppender();
            console.Layout = patternLayoutConsole;
            console.ActivateOptions();
            hierarchy.Root.AddAppender(console);

            hierarchy.Root.Level = hierarchy.LevelMap[_settings.Get("loglevel")] == null ? log4net.Core.Level.Error : hierarchy.LevelMap[_settings.Get("loglevel")];
            hierarchy.Configured = true;
        }

        private  bool VersionMenor(string versionactual, string versionlocal)
        {
            int[] vactual = Version(versionactual);
            int[] vlocal = Version(versionlocal);

            if (vactual.Length == 3)
            {
                if (vlocal[0] < vactual[0])
                    return true;
                else if (vlocal[0] == vactual[0] && vlocal[1] < vactual[1])
                    return true;
                else if (vlocal[0] == vactual[0] && vlocal[1] == vactual[1] && vlocal[2] < vactual[2])
                    return true;

            }

            if (vactual.Length == 4)
            {
                if (vlocal[0] < vactual[0])
                    return true;
                else if (vlocal[0] == vactual[0] && vlocal[1] < vactual[1])
                    return true;
                else if (vlocal[0] == vactual[0] && vlocal[1] == vactual[1] && vlocal[2] < vactual[2])
                    return true;
                else if (vlocal[0] == vactual[0] && vlocal[1] == vactual[1] && vlocal[2] == vactual[2] && vlocal[3] < vactual[3])
                    return true;
            }


            return false;
        }

        private  int[] Version(string version)
        {
            string phrase = version;
            string[] words = phrase.Split('.');

            int[] ver = Array.ConvertAll(words, int.Parse);


            if (ver.Length == 3)
            {
                //ver = new int[4];
                int[] pospunto = new int[3];
                pospunto[0] = version.IndexOf(".");
                pospunto[1] = version.IndexOf(".", pospunto[0] + 1);
                pospunto[2] = version.IndexOf(".", pospunto[1] + 1);
                try
                {
                    ver[0] = int.Parse(version.Substring(0, pospunto[0]));
                    ver[1] = int.Parse(version.Substring(pospunto[0] + 1, pospunto[1] - pospunto[0] - 1));
                    //string parche = version.Substring(pospunto[1] + 1, pospunto[2] - pospunto[1] - 1);
                    //ver[2] = int.Parse(NormalizarVersion(parche));
                    string parche = version.Substring(pospunto[1] + 1, pospunto[2] - pospunto[1] - 1);
                    ver[2] = int.Parse(NormalizarVersion(parche));
                }
                catch { }
                //return ver;
            }

            if (ver.Length == 4)
            {
                //ver = new int[4];
                int[] pospunto = new int[3];
                pospunto[0] = version.IndexOf(".");
                pospunto[1] = version.IndexOf(".", pospunto[0] + 1);
                pospunto[2] = version.IndexOf(".", pospunto[1] + 1);
                try
                {
                    ver[0] = int.Parse(version.Substring(0, pospunto[0]));
                    ver[1] = int.Parse(version.Substring(pospunto[0] + 1, pospunto[1] - pospunto[0] - 1));
                    string parche = version.Substring(pospunto[1] + 1, pospunto[2] - pospunto[1] - 1);
                    ver[2] = int.Parse(NormalizarVersion(parche));
                    string revision = version.Substring(pospunto[2] + 1, version.Length - pospunto[2] - 1);
                    ver[3] = int.Parse(NormalizarVersion(revision));
                }
                catch { }
                //return ver;
            }

            return ver;
        }

        string NormalizarVersion(string ver)
        {
            byte[] asciiBytes = System.Text.Encoding.ASCII.GetBytes(ver);
            string normalizado = "";
            foreach (byte digito in asciiBytes)
            {
                int ascii = int.Parse(digito.ToString());

                if (ascii > 57 || ascii < 48)
                    break;
                normalizado += Convert.ToChar(digito).ToString();
            }
            return normalizado;
        }


    }
}
