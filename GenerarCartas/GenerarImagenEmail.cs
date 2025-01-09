using DBCore;
using DBImaging;
using DBIntegration.SchemaSantander;
using DBIntegration;
using EmailSendValidationService;
//using Microsoft.Reporting.WinForms;
using Miharu.Desktop.Library.Config;
using Miharu.FileProvider.Library;
using Slyg.Tools.Imaging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DBImaging.SchemaProcess;
using EmailSendValidationService.Logs;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using Slyg.Tools.Imaging.FreeImageAPI.Plugins;
using Slyg.Tools;
//using Microsoft.ReportingServices.Interfaces;
using ServiceEmailSendValidation.Converters;
using System.Net.Http;
using DBTools;
using EmailSendValidationService.Config;
using PuppeteerSharp;
//using static Miharu.Desktop.Library.Config.DesktopConfig;

namespace ServiceEmailSendValidation.GenerarCartas
{
    public class GenerarImagenEmail
    {
        #region "Declaraciones"

        protected const int MaxThumbnailWidth = 60;
        protected const int MaxThumbnailHeight = 80;
        protected int _ImageCount;
        protected Logs _DataLog;

        Guid identifier;

        protected DBCore.SchemaImaging.TBL_FileRow _FileImagingRow;

        private static readonly object _lockObject = new object();  // Objeto para sincronizar el acceso a escritura DBTools y asegurar ejecución por un solo hilo a la vez.

        #endregion

        #region "Constructor"

        public GenerarImagenEmail(Logs dataLog)
        {
            _DataLog = dataLog;
            identifier = Guid.NewGuid();
        }

        #endregion

        #region "Metodos"

        public void GenerarCartas(DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail,ref IBrowser _browser)
        {
            DBCoreDataBaseManager dbmCore = null;
            DBImagingDataBaseManager dbmImaging = null;

            string fileName = string.Empty;
            string htmlFilePath = string.Empty;

            try
            {
                dbmImaging = new DBImaging.DBImagingDataBaseManager(Program.ConnectionStrings.Imaging);
                dbmCore = new DBCore.DBCoreDataBaseManager(Program.ConnectionStrings.Core);

                dbmImaging.Connection_Open(Program.Config.usuario_log);
                dbmCore.Connection_Open(Program.Config.usuario_log);

                var dtProyecto = dbmImaging.SchemaConfig.CTA_Proyecto.DBFindByfk_Entidadfk_Proyecto(itemfiltertrackingMail.fk_Entidad, itemfiltertrackingMail.fk_Proyecto);

                if (dtProyecto.Count > 0)
                {
                    Program.ProyectoImagingRow = dtProyecto[0].ToCTA_ProyectoSimpleType();

                    var OTDataTable = dbmImaging.SchemaProcess.CTA_OT_Servidor_Centro_Procesamiento.DBFindByfk_Entidadfk_Proyecto(itemfiltertrackingMail.fk_Entidad, itemfiltertrackingMail.fk_Proyecto);

                    if (OTDataTable != null && OTDataTable.Rows.Count > 0)
                    {
                        var filterOTDataTable = new DBImaging.SchemaProcess.CTA_OT_Servidor_Centro_ProcesamientoDataTable();

                        OTDataTable
                            .Where(row => row.fk_Expediente == itemfiltertrackingMail.fk_Expediente && row.fk_Folder == itemfiltertrackingMail.fk_Folder)
                            .CopyToDataTable()
                            .Rows                                                           // divide en filas
                            .Cast<DataRow>()
                            .ToList()
                            .ForEach(row => filterOTDataTable.ImportRow(row));    // almacena cada fila en una fila de filterOTDataTable

                        if (OTDataTable.Count > 0)
                        {
                            var firstRowOTData = filterOTDataTable[0];

                            // Asigna valores por defecto si es necesario por valores NULL
                            itemfiltertrackingMail = AssignDefaultValuesWhenNull(itemfiltertrackingMail);

                            // Actualiza data en TBL_queue y TBL_TrackingMail para enviar correo
                            //var sendEmailDate = ProcessEmailQueueAndUpdateTracking(itemfiltertrackingMail);
                            var sendEmailDate = DateTime.Now;

                            // Construcción de rutas de archivo
                            string fullPath = Path.Combine(Program.AppPath, Program.TempPath);
                            EnsureDirectoryExists(fullPath);                            
                            htmlFilePath = Path.Combine(fullPath, $"{identifier}.html");    // ruta HTML Correo enviado
                            fileName = Path.Combine(fullPath, $"{identifier}.tif");         // ruta imagen TIF del correo enviado

                            // Obtener valores de CC y CCO con validación
                            string ccQueue = GetFormattedEmailQueue(itemfiltertrackingMail.CC_Queue);
                            string ccoQueue = GetFormattedEmailQueue(itemfiltertrackingMail.CCO_Queue);
                            string emailRecipients = $"{ccQueue} {ccoQueue}".Trim();

                            var emailFrom = itemfiltertrackingMail.EmailFrom;
                            var emailSendDate = sendEmailDate;
                            var cultura = CultureInfo.GetCultureInfo("es-ES");
                            var emailAddress = itemfiltertrackingMail.EmailAddress_Queue;
                            var emailSubject = itemfiltertrackingMail.Subject_Queue;
                            var emailMessage = itemfiltertrackingMail.Message_Queue;

                            // Generar HTML del correo enviado
                            var emailHtmlBuilder = new EmailHtmlBuilder(emailFrom, emailSendDate, cultura, emailAddress, emailRecipients, emailSubject, emailMessage);
                            string htmlContent = emailHtmlBuilder.GenerateHtml();                            
                            File.WriteAllText(htmlFilePath, htmlContent);             // Guardar el contenido HTML en un archivo temporal

                            // Conversión de HTML a TIFF
                            var pageWidth = Program.ConnectionParameterSystemStrings.EmailEvidencePageWidth;
                            var pageHeigth = Program.ConnectionParameterSystemStrings.EmailEvidencePageHeigth;
                            var marginTop = Program.ConnectionParameterSystemStrings.ValueDefaultMarginTop;
                            
                            var converter = new HtmlToTiffConverter(fullPath, pageWidth, pageHeigth, marginTop, _DataLog, _browser);  // Instanciar el convertidor                            
                            converter.ConvertHtmlToTiff(htmlFilePath, fileName);                                                      // Convertir HTML a TIFF
                            
                            short foliosImageEmail = (short)ImageManager.GetFolios(fileName);                                        // Folios de la evidencia de correo.
                            short fileImageEmail = (short)itemfiltertrackingMail.fk_File;                                            // File Carta Respuesta tracking mail

                            var servidor = dbmImaging.SchemaCore.CTA_Servidor.DBFindByfk_Entidadid_Servidor(firstRowOTData.fk_Entidad_Servidor, firstRowOTData.fk_Servidor)[0].ToCTA_ServidorSimpleType();
                            var centro = dbmImaging.SchemaSecurity.CTA_Centro_Procesamiento.DBFindByfk_Entidadfk_Sedeid_Centro_Procesamiento(firstRowOTData.fk_Entidad_Procesamiento, firstRowOTData.fk_Sede_Procesamiento_Cargue, firstRowOTData.fk_Centro_Procesamiento_Cargue)[0].ToCTA_Centro_ProcesamientoSimpleType();

                            if (centro != null && servidor != null)
                            {
                                FileProviderManager manager = null;

                                try
                                {
                                    manager = new FileProviderManager(servidor, centro, ref dbmImaging, Program.Config.usuario_log);
                                    manager.Connect();

                                    _ImageCount = GetImageCount(ref manager, itemfiltertrackingMail);

                                    var _FileProcessRow = GetFileProcessRow(ref dbmCore,itemfiltertrackingMail);
                                    
                                    if(_FileProcessRow != null)
                                    {
                                        var dtDocumentoImagin = dbmImaging.SchemaConfig.CTA_Documento.DBFindByfk_Entidadfk_Proyectoid_DocumentoEliminado(itemfiltertrackingMail.fk_Entidad, itemfiltertrackingMail.fk_Proyecto, _FileProcessRow.fk_Documento, false);
                                        
                                        if(dtDocumentoImagin != null && dtDocumentoImagin.Count > 0)
                                        {
                                            var dtDocumentoImaginRow = dtDocumentoImagin[0];
                                            if (dtDocumentoImaginRow.id_Documento_Correo_Evidencia == 0)
                                            {
                                                throw new InvalidOperationException("El valor de id_Documento_Correo_Evidencia es 0 para el documento con ID: " + _FileProcessRow.fk_Documento.ToString() + ". Este valor no es válido. Por favor, configure un valor válido.");
                                            }

                                            var formato = Utilities.GetEnumFormat(Program.ProyectoImagingRow.id_Formato_Imagen_Entrada.ToString());
                                            var compresion = Utilities.GetEnumCompression((DesktopConfig.FormatoImagenEnum)Program.ProyectoImagingRow.id_Formato_Imagen_Salida);

                                            var dtFile = dbmCore.SchemaProcess.TBL_File.DBFindByfk_Expedientefk_Folder(itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder);
                                            if (dtFile == null || dtFile.Count == 0) return;

                                            // Obtén el valor máximo de id_file
                                            short maxIdFile = dtFile.Max(file => file.id_File);
                                            fileImageEmail = (short)(maxIdFile + 1);

                                            var expedienteImageEmail = itemfiltertrackingMail.fk_Expediente;
                                            var folderImageEmail = itemfiltertrackingMail.fk_Folder;

                                            try
                                            {

                                                for (int folio = 1; folio <= (foliosImageEmail + _ImageCount); folio++)
                                                {
                                                    byte[] dataImage = null;
                                                    byte[] dataImageThumbnail = null;

                                                    if (folio <= foliosImageEmail)
                                                    {
                                                        dataImage = ImageManager.GetFolioData(fileName, folio, formato, compresion);
                                                        dataImageThumbnail = ImageManager.GetThumbnailData(fileName, folio, folio, MaxThumbnailWidth, MaxThumbnailHeight)[0];
                                                    }
                                                    else
                                                    {
                                                        short currentFolio = (short)(folio - foliosImageEmail);
                                                        manager.GetFolio(expedienteImageEmail, folderImageEmail, (short)itemfiltertrackingMail.fk_File, _FileImagingRow.id_Version, currentFolio, ref dataImage, ref dataImageThumbnail);
                                                    }

                                                    // Debug - muestra las imagenes para almacenar en el nuevo file
                                                    //string FileName2 = fullPath + identifier + "_" + folio.ToString() + ".tif";
                                                    //using (var fs = new FileStream(FileName2, FileMode.Create))
                                                    //{
                                                    //    fs.Write(dataImage, 0, dataImage.Length);
                                                    //    fs.Close();
                                                    //}//////////////////////////////////////////

                                                    if (folio == 1)
                                                    {
                                                        Guid guidImage = Guid.NewGuid();

                                                        var fileImgType = new DBCore.SchemaImaging.TBL_FileType();
                                                        fileImgType.fk_Expediente = expedienteImageEmail;
                                                        fileImgType.fk_Folder = (short)folderImageEmail;
                                                        fileImgType.fk_File = fileImageEmail;
                                                        fileImgType.id_Version = 1;
                                                        fileImgType.File_Unique_Identifier = guidImage;
                                                        fileImgType.Folios_Documento_File = (short)(foliosImageEmail + _ImageCount);
                                                        fileImgType.Tamaño_Imagen_File = 0;
                                                        fileImgType.Nombre_Imagen_File = "";
                                                        fileImgType.Key_Cargue_Item = "";
                                                        fileImgType.Save_FileName = "";
                                                        fileImgType.fk_Content_Type = Program.ProyectoImagingRow.Extension_Formato_Imagen_Salida;
                                                        fileImgType.fk_Usuario_Log = Program.Config.usuario_log;
                                                        fileImgType.Validaciones_Opcionales = false;
                                                        fileImgType.Es_Anexo = false;
                                                        fileImgType.fk_Anexo = null;
                                                        fileImgType.fk_Entidad_Servidor = firstRowOTData.fk_Entidad_Servidor;
                                                        fileImgType.fk_Servidor = firstRowOTData.fk_Servidor;
                                                        fileImgType.Fecha_Creacion = SlygNullable.SysDate;
                                                        fileImgType.Fecha_Transferencia = null;
                                                        fileImgType.En_Transferencia = false;
                                                        fileImgType.fk_Entidad_Servidor_Transferencia = null;
                                                        fileImgType.fk_Servidor_Transferencia = null;

                                                        var fileProcesType = new DBCore.SchemaProcess.TBL_FileType();
                                                        fileProcesType.fk_Expediente = expedienteImageEmail;
                                                        fileProcesType.fk_Folder = (short)folderImageEmail;
                                                        fileProcesType.id_File = fileImageEmail;
                                                        fileProcesType.File_Unique_Identifier = guidImage;
                                                        fileProcesType.fk_Documento = dtDocumentoImaginRow.id_Documento_Correo_Evidencia;
                                                        fileProcesType.Folios_File = ((SlygNullable<short>)(foliosImageEmail + _ImageCount));
                                                        fileProcesType.Monto_File = 0;
                                                        fileProcesType.CBarras_File = expedienteImageEmail.ToString() + folderImageEmail.ToString() + fileImageEmail;

                                                        var FileEstadoType = new DBCore.SchemaProcess.TBL_File_EstadoType();
                                                        FileEstadoType.fk_Expediente = expedienteImageEmail;
                                                        FileEstadoType.fk_Folder = (short)folderImageEmail;
                                                        FileEstadoType.fk_File = fileImageEmail;
                                                        FileEstadoType.Modulo = new Slyg.Tools.SlygNullable<byte>((byte)ServiceConfig.Modulo.Imaging);
                                                        FileEstadoType.fk_Estado = 38;  // estado Indexado 
                                                        FileEstadoType.fk_Usuario = Program.Config.usuario_log;
                                                        FileEstadoType.Fecha_Log = DateTime.Now;

                                                        lock (_lockObject)
                                                        {
                                                            dbmCore.SchemaProcess.TBL_File.DBInsert(fileProcesType);
                                                            dbmCore.SchemaProcess.TBL_File_Estado.DBInsert(FileEstadoType);
                                                            dbmCore.SchemaImaging.TBL_File.DBInsert(fileImgType);
                                                            manager.CreateItem((long)expedienteImageEmail, folderImageEmail, fileImageEmail, 1, Program.ProyectoImagingRow.Extension_Formato_Imagen_Salida, identifier);
                                                        }
                                                    }

                                                    lock (_lockObject)
                                                    {
                                                        manager.CreateFolio((long)expedienteImageEmail, (short)folderImageEmail, fileImageEmail, 1, (short)folio, dataImage, dataImageThumbnail, false);
                                                    }
                                                }

                                                // Actualiza data en TBL_queue y TBL_TrackingMail para enviar correo
                                                ProcessEmailQueueAndUpdateTracking(itemfiltertrackingMail, sendEmailDate);
                                            }
                                            catch
                                            {
                                                lock (_lockObject)
                                                {
                                                    // eliminamos si existe el registro
                                                    var dataTableProcessFile = dbmCore.SchemaProcess.TBL_File.DBGet(expedienteImageEmail, folderImageEmail, fileImageEmail);
                                                    if (dataTableProcessFile != null && dataTableProcessFile.Count > 0)
                                                    {
                                                        dbmCore.SchemaProcess.TBL_File.DBDelete(expedienteImageEmail, folderImageEmail, fileImageEmail);
                                                    }

                                                    // eliminamos si existe el registro
                                                    var dataTableImagingFile = dbmCore.SchemaImaging.TBL_File.DBGet(expedienteImageEmail, folderImageEmail, fileImageEmail, 1);
                                                    if (dataTableImagingFile != null && dataTableImagingFile.Count > 0)
                                                    {
                                                        dbmCore.SchemaImaging.TBL_File.DBDelete(expedienteImageEmail, folderImageEmail, fileImageEmail, dataTableImagingFile[0].id_Version);
                                                    }

                                                    // Borrar los folios
                                                    manager.DeleteItem(expedienteImageEmail, folderImageEmail, fileImageEmail, 1);
                                                    throw;
                                                }
                                            }
                                        }                                        
                                    }                                                                       
                                }
                                catch
                                {
                                    throw;
                                }
                                finally
                                {
                                    if ((manager != null))
                                         manager.Disconnect();
                                }
                            }         
                            DeleteFilesIfExist(fileName, htmlFilePath);    // Eliminar archivo HTML y imagen Tiff correo enviado          
                        }                                               
                    }
                }
            }
            catch
            {
                DeleteFilesIfExist(fileName, htmlFilePath);    // Eliminar archivo HTML y imagen Tiff correo enviado
                throw;
            }
            finally
            {
                if ((dbmImaging != null))
                    dbmImaging.Connection_Close();
                if ((dbmCore != null))
                    dbmCore.Connection_Close();
            }
        }

        private void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        /// <summary>
        /// Elimina los archivos especificados si existen en las rutas proporcionadas.
        /// </summary>
        /// <param name="filePath">Ruta del archivo a eliminar</param>
        public void DeleteFilesIfExist(params string[] filePaths)
        {
            foreach (var filePath in filePaths)
            {
                // Validar si la ruta no es null ni está vacía antes de continuar
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
        }

        private DBTools.SchemaMail.TBL_Tracking_MailRow AssignDefaultValuesWhenNull(DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail)
        {

            itemfiltertrackingMail.fk_Usuario = Convert.IsDBNull(itemfiltertrackingMail["fk_Usuario"])
                                                ? 2  // cambiar numero
                                                : itemfiltertrackingMail.fk_Usuario;

            itemfiltertrackingMail.CC_Queue = Convert.IsDBNull(itemfiltertrackingMail["CC_Queue"])
                                                            ? ";"
                                                            : itemfiltertrackingMail.CC_Queue;

            itemfiltertrackingMail.CCO_Queue = Convert.IsDBNull(itemfiltertrackingMail["CCO_Queue"])
                                                            ? ";"
                                                            : itemfiltertrackingMail.CCO_Queue;

            itemfiltertrackingMail.Attach_Queue = Convert.IsDBNull(itemfiltertrackingMail["Attach_Queue"])
                                                            ? new byte[0] // Si está vacío, inicializa un arreglo vacío de bytes
                                                            : itemfiltertrackingMail.Attach_Queue;

            itemfiltertrackingMail.EmailFromDisplay = Convert.IsDBNull(itemfiltertrackingMail["EmailFromDisplay"])
                                                            ? string.Empty
                                                            : itemfiltertrackingMail.EmailFromDisplay;

            itemfiltertrackingMail.Detalle_Envio = Convert.IsDBNull(itemfiltertrackingMail["Detalle_Envio"])
                                                            ? string.Empty
                                                            : itemfiltertrackingMail.Detalle_Envio;

            return itemfiltertrackingMail;
        }




        #endregion

        #region "Funciones"

        /// <summary>
        /// Evalúa la existencia de un folio verificando un rango de índices.
        /// Busca el folio desde <paramref name="startIndex"/> hasta <paramref name="endIndex"/> 
        /// (por defecto es 1000). El método actualiza el último índice donde se encontró el folio.
        /// </summary>
        /// /// <param name="manager">Referencia a la instancia de <see cref="FileProviderManager"/> utilizada para gestionar operaciones de archivos.</param>
        /// <param name="itemfiltertrackingMail">La fila de seguimiento de correo utilizada como filtro para la verificación de la existencia del folio.</param>
        /// <param name="startIndex">El índice desde el cual comenzar a verificar la existencia del folio.</param>
        /// <param name="endIndex">El índice hasta el cual verificar la existencia del folio (por defecto es 1000).</param>
        /// <returns>El último índice donde se encontró que el folio existe. Retorna -1 si no se encontró ningún folio en el rango especificado.</returns>
        public short EvaluateFolio(ref FileProviderManager manager, DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail, short startIndex, short endIndex = 1000)
        {
            short lastTrueIndex = -1; // Inicializamos el índice del último true

            for (short i = startIndex; i <= endIndex; i++)
            {
                if (ExistFolio(ref manager, itemfiltertrackingMail, i, 1))
                {
                    lastTrueIndex = i; // Guardamos el índice donde fue true
                    continue; // Continuamos evaluando
                }
                else
                {
                    break; // Salimos si ExistFolio devuelve false
                }
            }

            return (lastTrueIndex <= startIndex)? startIndex: lastTrueIndex; // Retornamos el último índice donde fue true si no es menor al indice de inicio
        }

        private bool ExistFolio(ref FileProviderManager manager, DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail, short fileImageEmail, int folio)
        {
            bool existFolio = false;

            try
            {
                existFolio =  manager.ExistFolio(itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, fileImageEmail, 1, (short)folio);
                
                return existFolio;
            }
            catch 
            {
                return existFolio;
            }
        }


        private DBCore.SchemaProcess.TBL_FileRow GetFileProcessRow(ref DBCoreDataBaseManager dbmCore, DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail)
        {
            try
            {
                var fileProcessDataTable = dbmCore.SchemaProcess.TBL_File.DBFindByfk_Expedientefk_Folderid_File(itemfiltertrackingMail.fk_Expediente,
                                                                                                                itemfiltertrackingMail.fk_Folder,
                                                                                                                (SlygNullable<short>)itemfiltertrackingMail.fk_File);
                
                if(fileProcessDataTable != null && fileProcessDataTable.Count > 0)
                {
                    return fileProcessDataTable[0];
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                _DataLog.AddErrorEntry($"¡¡WARNING!! Expediente: {itemfiltertrackingMail.fk_Expediente}, No contiene file generación de cartas, error:" + ex.Message);
                throw;
            }           
        }



        private string GetFormattedEmailQueue(string emailQueue)
        {
            if (!string.IsNullOrWhiteSpace(emailQueue) && emailQueue != ";")
            {
                return $" {emailQueue}".Trim();
            }
            return string.Empty;
        }

        private DataRow GetSeguimientoCorreoRow(DBImaging.SchemaProcess.CTA_OT_Servidor_Centro_ProcesamientoRow dataOTServidor, DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail)
        {

           DBToolsDataBaseManager dbmTools = null;

            try
            {
                dbmTools = new DBTools.DBToolsDataBaseManager(Program.ConnectionStrings.Tools);
                dbmTools.Connection_Open();

                var seguimientoCorreoDataTable = dbmTools.SchemaMail.PA_Seguimiento_Correo.DBExecute(dataOTServidor.fk_Entidad, 
                                                                                                    dataOTServidor.fk_Proyecto, 
                                                                                                    dataOTServidor.fk_fecha_proceso, 
                                                                                                    dataOTServidor.fk_fecha_proceso, 
                                                                                                    dataOTServidor.fk_OT,
                                                                                                    -1);

                // Verifica si la tabla es nula o no tiene filas
                if (seguimientoCorreoDataTable == null || seguimientoCorreoDataTable.Rows.Count == 0) return null;

                // Filtrar usando LINQ
                var filteredRows = seguimientoCorreoDataTable.AsEnumerable()
                    .Where(row =>
                        row.Field<long>("fk_Expediente") == itemfiltertrackingMail.fk_Expediente &&
                        row.Field<short>("fk_Folder") == itemfiltertrackingMail.fk_Folder &&
                        row.Field<short>("id_File") == itemfiltertrackingMail.fk_File
                    );

                // Verificar si el filtro devolvió filas
                var resultRow = filteredRows.FirstOrDefault();
                if (resultRow == null) return null;

                return resultRow;
            }
            catch
            {
                throw;
            }
            finally
            {
                if ((dbmTools != null))
                    dbmTools.Connection_Close();
            }
        }


        public void ProcessEmailQueueAndUpdateTracking(DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail, DateTime currentDate)
        {
            lock (_lockObject)
            {
                DBToolsDataBaseManager dbmTools = null;

                try
                {
                    dbmTools = new DBTools.DBToolsDataBaseManager(Program.ConnectionStrings.Tools);
                    dbmTools.Connection_Open();

                    dbmTools.SchemaMail.PA_Insert_Queue.DBExecute(itemfiltertrackingMail.id_Tracking_Mail, currentDate, itemfiltertrackingMail.EmailAddress_Queue);
                }
                catch
                {
                    throw;
                }
                finally
                {
                    if ((dbmTools != null))
                        dbmTools.Connection_Close();
                }
            }
        }

        /*
        public DateTime ProcessEmailQueueAndUpdateTracking(DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail)
        {
            lock (_lockObject)
            {
                DBToolsDataBaseManager dbmTools = null;

                try
                {
                    dbmTools = new DBTools.DBToolsDataBaseManager(Program.ConnectionStrings.Tools);
                    dbmTools.Connection_Open();

                    var currentDate = DateTime.Now;
                    dbmTools.SchemaMail.PA_Insert_Queue.DBExecute(itemfiltertrackingMail.id_Tracking_Mail, currentDate, itemfiltertrackingMail.EmailAddress_Queue);

                    return currentDate;
                }
                catch
                {
                    throw;
                }
                finally
                {
                    if ((dbmTools != null))
                        dbmTools.Connection_Close();
                }
            }
        }*/

        /// <summary>
        /// Obtiene el numero de folios de un file
        /// </summary>
        /// <param name="nManager"></param>
        /// <param name="itemfiltertrackingMail">fila de archivos con los datos para extraer los folios</param>
        /// <returns></returns>
        private int GetImageCount(ref FileProviderManager nManager, DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail)
        {
            int imageCount;

            try
            {
                var fileImagingDataTable = GetFileDataTable(itemfiltertrackingMail);
                if (fileImagingDataTable == null || fileImagingDataTable.Count == 0)
                {
                    _DataLog.AddErrorEntry($"¡¡WARNING!! Expediente: {itemfiltertrackingMail.fk_Expediente}, No contiene Datos obtenidos de file en dbmCore.SchemaImaging.TBL_File");
                    return 0;
                }

                _FileImagingRow = fileImagingDataTable[fileImagingDataTable.Count - 1];

                if (_FileImagingRow.Es_Anexo)
                {
                    imageCount = nManager.GetFolios(_FileImagingRow.fk_Anexo);
                }
                else
                {
                    imageCount = nManager.GetFolios(_FileImagingRow.fk_Expediente, _FileImagingRow.fk_Folder, _FileImagingRow.fk_File, _FileImagingRow.id_Version);
                }
                return imageCount;
            }
            catch ( Exception ex)
            {
                _DataLog.AddErrorEntry($"¡¡WARNING!! Expediente: {itemfiltertrackingMail.fk_Expediente}, No contiene o no se obtuvo imagenes de la Carta de respuesta para procesar, error:" + ex.Message);
                return 0;
            }

        }

        /// <summary>
        /// Obtiene los datos de un archivo desde la base de datos de Core.
        /// </summary>
        /// <param name="currentBlockedFile">Objeto que representa la fila de captura bloqueada en el panel de control.</param>
        /// <returns>Una tabla de datos que contiene la información del archivo, o null si ocurre un error.</returns>
        private DBCore.SchemaImaging.TBL_FileDataTable GetFileDataTable(DBTools.SchemaMail.TBL_Tracking_MailRow filtertrackingMailRow)
        {
            DBCore.DBCoreDataBaseManager dbmCore = null;

            try
            {
                dbmCore = new DBCore.DBCoreDataBaseManager(Program.ConnectionStrings.Core);
                dbmCore.Connection_Open(Program.Config.usuario_log);

                DBCore.SchemaImaging.TBL_FileDataTable fileDataTable = dbmCore.SchemaImaging.TBL_File.DBGet(filtertrackingMailRow.fk_Expediente,
                                                                                                                filtertrackingMailRow.fk_Folder,
                                                                                                                (Slyg.Tools.SlygNullable<short>)filtertrackingMailRow.fk_File,
                                                                                                                null);
                return fileDataTable;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (dbmCore != null) dbmCore.Connection_Close();
            }
        }

        #endregion

    }
}
