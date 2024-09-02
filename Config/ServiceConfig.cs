using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailSendValidationService.Config
{
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
            public string OCR;
            public string Tools;
            public string Softrac;
            public string Core_Risks;
            public string Imaging_Risks;
            public string Archiving_Risks;
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

        #endregion


        #region " Constructores "

        public ServiceConfig()
        {
            usuario_log = 2;                                                // TODO: Ajustar de acuerdo al usuario
            Intervalo = 5000;
            ConnectionStringSecurity = ConfigurationManager.AppSettings["ConnectionStringSecurity"];
            SystemLogPath = ConfigurationManager.AppSettings["SystemLogPath"];
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

        #endregion

    }
}
