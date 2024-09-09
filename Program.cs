using EmailSendValidationService.Config;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace EmailSendValidationService
{
    internal static class Program
    {

        #region " Declaraciones "

        public static ServiceConfig Config = new ServiceConfig();
        public static ServiceConfig.TypeConnectionString ConnectionStrings;
        public static ServiceConfig.TypeParametersString ConnectionParameterSystemStrings;
        public static DBImaging.SchemaConfig.CTA_ProyectoSimpleType ProyectoImagingRow = new DBImaging.SchemaConfig.CTA_ProyectoSimpleType();

        #endregion

        #region " Propiedades "

        internal static string AppPath
        {
            get
            {
                return System.Windows.Forms.Application.StartupPath.TrimEnd('\\') + "\\";
            }
        }

        public static string TempPath
        {
            get
            {
                return "temp\\";
            }
        }

        #endregion

        #region " Metodos "

        internal static string AppDataPath
        {
            get
            {
                DateTime today = DateTime.Now;                                                      // Obtener la fecha actual
                string fechaActual = today.ToString("yyyyMMdd");                                    // Formato AAAAMMDD

                // Construir la ruta del archivo de registro
                string _SystemLogPath = Config.SystemLogPath;
                if (!Directory.Exists(_SystemLogPath)) Directory.CreateDirectory(_SystemLogPath);   // Crear el directorio si no existe

                _SystemLogPath += $"\\log_{fechaActual}.txt";

                return _SystemLogPath;
            }
        }



        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new Service1()
            };
            ServiceBase.Run(ServicesToRun);
        }


        #endregion
    }
}
