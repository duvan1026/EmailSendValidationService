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
    public static class Program
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
        // Mueve la lógica de inicialización a este método
        public static void Initialize()
        {
            // Asigna las variables que necesitas
            ConnectionStrings = Config.GetCadenasConexion();

            if (ConnectionStrings.Security == "")
            {
                throw new Exception("No se pudo obtener la cadena de conexión a la base de datos Security");
            }

            if (ConnectionStrings.Core == "")
            {
                throw new Exception("No se pudo obtener la cadena de conexión a la base de datos Core");
            }

            if (ConnectionStrings.Imaging == "")
            {
                throw new Exception("No se pudo obtener la cadena de conexión a la base de datos Imaging");
            }

            if (ConnectionStrings.OCR == "")
            {
                throw new Exception("No se pudo obtener la cadena de conexión a la base de datos OCR");
            }
        }



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
        public static void Main()
        {
            //ServiceBase[] ServicesToRun;
            //ServicesToRun = new ServiceBase[]
            //{
            //    new Service1()
            //};
            //ServiceBase.Run(ServicesToRun);

            if (Environment.UserInteractive)
            {
                // Ejecutar el servicio en modo interactivo (como si fuera una consola para depuración)
                var service = new Service1();
                service.IniciarServicio();

                Console.WriteLine("Servicio iniciado en modo interactivo.");
                Console.WriteLine("Presiona cualquier tecla para detener...");
                Console.ReadKey();  // Esperar a que el usuario presione una tecla para simular la detención del servicio

                service.DetenerServicio();
            }
            else
            {
                // Ejecutar el servicio normalmente como un servicio de Windows
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                    new Service1()
                };
                ServiceBase.Run(ServicesToRun);
            }
        }


        #endregion
    }
}
