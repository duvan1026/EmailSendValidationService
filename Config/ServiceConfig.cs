using Slyg.Data.Schemas;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailSendValidationService.Config
{

    public static class ParameterSystem
    {
        public const string ThreadCountServiceESV = "thread_Count_ServiceESV";
        public const string EmailEvidencePageWidth = "email_Evidence_Page_Width";
        public const string EmailEvidencePageHeigth = "email_Evidence_Page_Heigth";
        public const string ValueDefaultMarginTop = "value_Default_Margin_Top";
    }

    public class ServiceConfig
    {

        #region " Estructuras "

        [Serializable]
        public struct TypeConnectionString
        {
            public string Security;
            public string Archiving;
            public string Core;
            public string Imaging;
            public string Integration;
            public string OCR;
            public string Tools;
            public string Softrac;
            public string Core_Risks;
            public string Imaging_Risks;
            public string Archiving_Risks;
        }

        [Serializable]
        public struct TypeParametersString
        {
            public int ThreadCountServiceESV;
            public int EmailEvidencePageWidth;
            public int EmailEvidencePageHeigth;
            public int ValueDefaultMarginTop;
        }

        #endregion

        #region " Enumeraciones"
        public enum Modulo : byte
        {
            Security = 0,
            Imaging = 2,
            Core = 6,
            Archiving = 7,
            PunteoBanAgrario = 13,
            Integration = 30,
            Tools = 3,
            Softrac = 35,
            Core_Risks = 38,
            Imaging_Risks = 39,
            Archiving_Risks = 40,
            OCR = 41
        }

        #endregion

        #region " Propiedades "

        public int usuario_log { get; private set; }
        public int Intervalo { get; set; }
        public string ConnectionStringSecurity { get; }
        public string SystemLogPath { get; }
        
        public string NameService { get; }

        #endregion


        #region " Constructores "

        public ServiceConfig()
        {
            usuario_log = 2;                                                // TODO: Ajustar de acuerdo al usuario
            Intervalo = 5000;
            ConnectionStringSecurity = ConfigurationManager.AppSettings["ConnectionStringSecurity"];
            SystemLogPath = ConfigurationManager.AppSettings["SystemLogPath"];
            NameService = "EmailSendValidation";
        }

        #endregion

        #region " Metodos"

        private DBSecurity.SchemaSecurity.TBL_ModuloDataTable GetModuloDataTable()
        {
            DBSecurity.DBSecurityDataBaseManager dbmSecurity = null;

            try
            {
                dbmSecurity = new DBSecurity.DBSecurityDataBaseManager(ConnectionStringSecurity);
                dbmSecurity.Connection_Open(Program.Config.usuario_log);

                DBSecurity.SchemaSecurity.TBL_ModuloDataTable moduloDataTable = dbmSecurity.SchemaSecurity.TBL_Modulo.DBGet(null);

                return moduloDataTable;
            }
            catch
            {
                throw;
            }
            finally
            {
                dbmSecurity?.Connection_Close();
            }
        }

        #endregion


        #region " funciones "

        public ServiceConfig.TypeConnectionString GetCadenasConexion()
        {
            ServiceConfig.TypeConnectionString cadenas = new ServiceConfig.TypeConnectionString();

            var connectionStrings = GetModuloDataTable();

            foreach (var Modulo in connectionStrings)
            {
                switch ((ServiceConfig.Modulo)Modulo.id_Modulo)
                {
                    case ServiceConfig.Modulo.Security:
                        cadenas.Security = Modulo.ConnectionString;
                        break;
                    case ServiceConfig.Modulo.Imaging:
                        cadenas.Imaging = Modulo.ConnectionString;
                        break;
                    case ServiceConfig.Modulo.Integration:
                        cadenas.Integration = Modulo.ConnectionString;
                        break;
                    case ServiceConfig.Modulo.Core:
                        cadenas.Core = Modulo.ConnectionString;
                        break;
                    case ServiceConfig.Modulo.Archiving:
                        cadenas.Archiving = Modulo.ConnectionString;
                        break;
                    case ServiceConfig.Modulo.Tools:
                        cadenas.Tools = Modulo.ConnectionString;
                        break;
                    case ServiceConfig.Modulo.Softrac:
                        cadenas.Softrac = Modulo.ConnectionString;
                        break;
                    case ServiceConfig.Modulo.OCR:
                        cadenas.OCR = Modulo.ConnectionString;
                        break;
                }
            }

            return cadenas;
        }

        public ServiceConfig.TypeParametersString GetParametersSystem()
        {
            ServiceConfig.TypeParametersString parameters = new ServiceConfig.TypeParametersString();

            var connectionStrings = GetParametersSystemDataTable();

            foreach (var parameter in connectionStrings)
            {
                switch (parameter.Nombre_Parametro)
                {
                    case ParameterSystem.ThreadCountServiceESV:
                        parameters.ThreadCountServiceESV = int.TryParse(parameter.Valor_Parametro, out int result) ? result : 5;                                       // se establecen 5 hilos por defecto
                        break;
                    case ParameterSystem.EmailEvidencePageWidth:
                        parameters.EmailEvidencePageWidth = int.TryParse(parameter.Valor_Parametro, out int resultPageWidth) ? resultPageWidth : 718;                                       // se establecen 5 hilos por defecto
                        break;
                    case ParameterSystem.EmailEvidencePageHeigth:
                        parameters.EmailEvidencePageHeigth = int.TryParse(parameter.Valor_Parametro, out int resultPageHeigth) ? resultPageHeigth : 1024;                                       // se establecen 5 hilos por defecto
                        break;
                    case ParameterSystem.ValueDefaultMarginTop:
                        parameters.ValueDefaultMarginTop = int.TryParse(parameter.Valor_Parametro, out int resultMarginTop) ? resultMarginTop : 20;                                       // se establecen 5 hilos por defecto
                        break;
                }
            }

            return parameters;
        }

        private DBTools.SchemaConfig.TBL_Parametro_SistemaDataTable GetParametersSystemDataTable()
        {
            DBTools.DBToolsDataBaseManager dbmTools = null;

            try
            {
                dbmTools = new DBTools.DBToolsDataBaseManager(Program.ConnectionStrings.Tools);
                dbmTools.Connection_Open();

                DBTools.SchemaConfig.TBL_Parametro_SistemaDataTable parametersDataTable = dbmTools.SchemaConfig.TBL_Parametro_Sistema.DBFindByfk_Entidadfk_Proyecto(0, 0);

                return parametersDataTable;
            }
            catch
            {
                throw;
            }
            finally
            {
                dbmTools?.Connection_Close();
            }
        }

        #endregion

    }
}
