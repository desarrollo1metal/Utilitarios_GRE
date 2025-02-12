using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Serilog;

namespace Utilitarios_GRE
{
    class sapObjUser
    {
        private SAPbouiCOM.Application sboApplication;
        private SAPbobsCOM.Company sboCompany;
        private int m_iErrCode = 0;
        private string m_sErrMsg = "";



        //public sapObjUser(SAPbouiCOM.Application _sboApplication, SAPbobsCOM.Company _sboCompany)
        //{
        //    sboApplication = _sboApplication;
        //    sboCompany = _sboCompany;
        //}

        public sapObjUser( SAPbobsCOM.Company _sboCompany)
        {
            //sboApplication = _sboApplication;
            sboCompany = _sboCompany;
        }

        public void CreaCampoMD(string NombreTabla, string NombreCampo, string DescCampo, SAPbobsCOM.BoFieldTypes TipoCampo,
           SAPbobsCOM.BoFldSubTypes SubTipo, int Tamano, SAPbobsCOM.BoYesNoEnum Obligatorio, string[] validValues,
            string[] validDescription, string valorPorDef, string tablaVinculada, string UDOVinculado)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;//Objeto para la gestion de campos de usuarios de SBO
            try
            {
                if (NombreTabla == null) NombreTabla = "";
                if (NombreCampo == null) NombreCampo = "";
                if (Tamano == 0) Tamano = 10;
                if (validValues == null) validValues = new string[0];
                if (validDescription == null) validDescription = new string[0];
                if (valorPorDef == null) valorPorDef = "";
                if (tablaVinculada == null) tablaVinculada = "";
                if (UDOVinculado == null) UDOVinculado = "";

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)sboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);//generamos una instacia del objeto para la gestion de campos SBO
                oUserFieldsMD.TableName = NombreTabla;//asignamos el nombre de la tabla en donde deseamos crear el campo
                oUserFieldsMD.Name = NombreCampo;//asignamos el nombre del campo a crear
                oUserFieldsMD.Description = DescCampo;//asignamos la descripcion del campo a crear
                oUserFieldsMD.Type = TipoCampo;//ajustamos el tipo del campo a crear
                if (TipoCampo != SAPbobsCOM.BoFieldTypes.db_Date) oUserFieldsMD.EditSize = Tamano;//verificamos que el tipo de campo no sea de tipo fecha para poder asignarle un tamaño
                oUserFieldsMD.SubType = SubTipo;//asignamos el subtipo del campo

                if (tablaVinculada != "") oUserFieldsMD.LinkedTable = tablaVinculada;//asignamos la tabla vinculada
                if (UDOVinculado != "") oUserFieldsMD.LinkedUDO = UDOVinculado;
                else
                {
                    if (validValues.Length > 0)//verificamos que el campo posea valores validos
                    {
                        for (int i = 0; i <= (validValues.Length - 1); i++)//Asignamos los valores validos para el campo
                        {
                            oUserFieldsMD.ValidValues.Value = validValues[i];
                            if (validDescription.Length > 0) oUserFieldsMD.ValidValues.Description = validDescription[i];
                            else oUserFieldsMD.ValidValues.Description = validValues[i];
                            oUserFieldsMD.ValidValues.Add();
                        }
                    }
                    oUserFieldsMD.Mandatory = Obligatorio;
                    if (valorPorDef != "") oUserFieldsMD.DefaultValue = valorPorDef;//Asignamos el valor por defecto para el campo
                }

                m_iErrCode = oUserFieldsMD.Add();//Obtenemos el valor de repuesta de SBO al crear el campo
                if (m_iErrCode != 0)//Validamos que no se hayan producido errores
                {
                    sboCompany.GetLastError(out m_iErrCode, out m_sErrMsg);//Obtenemos el mensaje de error ocurrido
                    Log.Error("codigo error : "  + m_iErrCode.ToString() + " -- mensaje : " + m_sErrMsg.ToString());

                    if (m_iErrCode != -5002) { }
                    //sboApplication.StatusBar.SetText(Util.NombAddon + ": Configuración verificada",
                    //    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else {

                    Log.Information(" Se ha creado el campo de usuario: " + NombreCampo + "en la tabla: " + NombreTabla);
                }
                //sboApplication.StatusBar.SetText(Util.NombAddon + " Se ha creado el campo de usuario: " + NombreCampo
                //        + "en la tabla: " + NombreTabla, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                //sboApplication.StatusBar.SetText(Util.NombAddon +
                //    ": Configuración verificada",
                //    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                Log.Error(ex,ex.Message);
                string error = ex.ToString();
            }
            finally
            {
                //Util.liberarObjeto(oUserFieldsMD);//Liberamos el objeto para la creacion del campo

                try
                {
                    if (oUserFieldsMD != null)
                    {
                        
                        Marshal.ReleaseComObject(oUserFieldsMD);
                        Log.Information("Objeto Liberado");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex,ex.Message);  
                    
                }

                oUserFieldsMD = null;
                GC.Collect();
            }
        }

    }
}
