using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using System.Configuration;

namespace Utilitarios_GRE
{

    sealed class Conexion
    {

        public static Company oCompany;
        public static clsConfig ConfiguracionGeneral;

        public static void InitializeCompany(SociedadBD sociedad)
        {
            oCompany = new Company();
            switch (ConfiguracionGeneral.ServidorSAP.DbType)
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
            oCompany.DbUserName = ConfiguracionGeneral.ServidorSAP.DbUserName;
            oCompany.DbPassword = ConfiguracionGeneral.ServidorSAP.DbPassword;
            oCompany.Server = ConfiguracionGeneral.ServidorSAP.Server;
            oCompany.CompanyDB = sociedad.DbName;
            oCompany.UserName = sociedad.SAPuser;
            oCompany.Password = sociedad.SAPpassword;
            oCompany.language = BoSuppLangs.ln_Spanish_La;
            oCompany.LicenseServer = ConfiguracionGeneral.ServidorSAP.LicenseServer;
            oCompany.UseTrusted = false;
        }

        public static void DestroyCompany()
        {
            if (oCompany != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany);
            oCompany = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public static void InicializarVarGlob()
        {
            Recordset oRecordSet = null;
            try
            {
                string sSQL = string.Empty;
                //sSQL = Util.VSSQLFactory.GetScript(oCompany, 6, null);

                ////oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                ////oRecordSet.DoQuery(sSQL);
                ////if (!oRecordSet.EoF)
                ////{
                ////    string ARET = oRecordSet.Fields.Item("U_BPP_PPAR").Value;
                ////    string RCode = oRecordSet.Fields.Item("U_BPP_PPRI").Value;
                ////    //Procesar.SociedadRetenedora = (ARET == "Y");
                ////    //Procesar.CodRet = RCode;
                ////    //Procesar.FullLog = Conexion.ConfiguracionGeneral.LogCompleto;
                ////}
                ////else
                ////{
                ////    //Procesar.SociedadRetenedora = false;
                ////    //Procesar.CodRet = "";
                ////}
            }
            catch { }
            finally
            {
                if (oRecordSet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }
    }

}
