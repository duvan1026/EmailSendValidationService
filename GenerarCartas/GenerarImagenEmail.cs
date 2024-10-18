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
                            var NotificacionRow = GetSeguimientoCorreoRow(firstRowOTData, itemfiltertrackingMail);

                            if (NotificacionRow != null)
                            {
                                int TipoNotificacion = (int)NotificacionRow["TipoNotificacion"];
                                var NotificacionParametrosDataTable = new DBCore.SchemaProcess.CTA_Notificacion_ParametrosDataTable();

                                lock (_lockObject)
                                {
                                    NotificacionParametrosDataTable = dbmCore.SchemaProcess.PA_Notificacion_Parametros.DBExecute(itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, (SlygNullable<short>)(itemfiltertrackingMail.fk_File - 1), TipoNotificacion);
                                }

                                foreach (var tblNotificacionParametroRow in NotificacionParametrosDataTable)
                                {
                                    switch (tblNotificacionParametroRow.fk_Parametro_tipo)
                                    {
                                        case 1:  // Subject
                                            itemfiltertrackingMail.Subject_Queue = itemfiltertrackingMail.Subject_Queue.Replace(tblNotificacionParametroRow.Parametro, tblNotificacionParametroRow.Valor_Parametro).Replace("<b>", "").Replace("</b>", "");
                                            break;
                                        case 2:  // Message
                                            itemfiltertrackingMail.Message_Queue = itemfiltertrackingMail.Message_Queue.Replace(tblNotificacionParametroRow.Parametro, tblNotificacionParametroRow.Valor_Parametro);
                                            break;
                                    }
                                }
                            }                                                   

                            // Asigna valores por defecto si es necesario por valores NULL
                            itemfiltertrackingMail = AssignDefaultValuesWhenNull(itemfiltertrackingMail);

                            // Actualiza data en TBL_queue y TBL_TrackingMail para enviar correo
                            var sendEmailDate = ProcessEmailQueueAndUpdateTracking(itemfiltertrackingMail);

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
                            
                            var converter = new HtmlToTiffConverter(fullPath, pageWidth, pageHeigth, marginTop);   // Instanciar el convertidor                            
                            converter.ConvertHtmlToTiff(htmlFilePath, fileName, ref _browser);                                   // Convertir HTML a TIFF
                            

                            short folios = (short)ImageManager.GetFolios(fileName);
                            var FolioBitmap = ImageManager.GetFolioBitmap(fileName, folios);
                            short fileImageEmail = (short)itemfiltertrackingMail.fk_File;

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

                                            var formato = Utilities.GetEnumFormat(Program.ProyectoImagingRow.id_Formato_Imagen_Entrada.ToString());
                                            var compresion = Utilities.GetEnumCompression((DesktopConfig.FormatoImagenEnum)Program.ProyectoImagingRow.id_Formato_Imagen_Salida);

                                            // Evalua el ultimo folio de ese expediente folder para almacenar la evidencia del correo
                                            short lastFolio = EvaluateFolio(ref manager, itemfiltertrackingMail, fileImageEmail);
                                            fileImageEmail = (short)(lastFolio + 1);

                                            for (int folio = 1; folio <= (folios + _ImageCount); folio++)
                                            {
                                                byte[] dataImage = null;
                                                byte[] dataImageThumbnail = null;

                                                if (folio <= folios)
                                                {
                                                    dataImage = ImageManager.GetFolioData(fileName, folio, formato, compresion);
                                                    dataImageThumbnail = ImageManager.GetThumbnailData(fileName, folio, folio, MaxThumbnailWidth, MaxThumbnailHeight)[0];
                                                }
                                                else
                                                {
                                                    short currentFolio = (short)(folio - folios);
                                                    manager.GetFolio(itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, (short)itemfiltertrackingMail.fk_File, _FileImagingRow.id_Version, currentFolio, ref dataImage, ref dataImageThumbnail);
                                                }

                                                // Debug - muestra las imagenes para almacenar en el nuevo file
                                                //string FileName2 = fullPath + identifier + "_" + folio.ToString() + ".tif";
                                                //using (var fs = new FileStream(FileName2, FileMode.Create))
                                                //{
                                                //    fs.Write(dataImage, 0, dataImage.Length);
                                                //    fs.Close();
                                                //}//////////////////////////////////////////

                                                // Verifica que exista el file para proceder a actualizarlo o insertarlo
                                                bool existFolio = ExistFolio(ref manager, itemfiltertrackingMail, fileImageEmail, (short)folio);

                                                if (folio == 1)
                                                {
                                                    var fileImgType = new DBCore.SchemaImaging.TBL_FileType();
                                                    fileImgType.fk_Expediente = itemfiltertrackingMail.fk_Expediente;
                                                    fileImgType.fk_Folder = (short)itemfiltertrackingMail.fk_Folder;
                                                    fileImgType.fk_File = fileImageEmail;
                                                    fileImgType.id_Version = 1;
                                                    fileImgType.File_Unique_Identifier = Guid.NewGuid();
                                                    fileImgType.Folios_Documento_File = (short)(folios + _ImageCount);
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
                                                    fileProcesType.fk_Expediente = itemfiltertrackingMail.fk_Expediente;
                                                    fileProcesType.fk_Folder = (short)itemfiltertrackingMail.fk_Folder;
                                                    fileProcesType.id_File = fileImageEmail;
                                                    fileProcesType.File_Unique_Identifier = Guid.NewGuid();
                                                    fileProcesType.fk_Documento = dtDocumentoImaginRow.id_Documento_Correo_Evidencia;
                                                    fileProcesType.Folios_File = ((SlygNullable<short>)(folios + _ImageCount));
                                                    fileProcesType.Monto_File = 0;
                                                    fileProcesType.CBarras_File = itemfiltertrackingMail.fk_Expediente.ToString() + itemfiltertrackingMail.fk_Folder.ToString() + fileImageEmail;

                                                    var FileEstadoType = new DBCore.SchemaProcess.TBL_File_EstadoType();
                                                    FileEstadoType.fk_Expediente = itemfiltertrackingMail.fk_Expediente;
                                                    FileEstadoType.fk_Folder = (short)itemfiltertrackingMail.fk_Folder;
                                                    FileEstadoType.fk_File = fileImageEmail;
                                                    FileEstadoType.Modulo = new Slyg.Tools.SlygNullable<byte>((byte)ServiceConfig.Modulo.Imaging);
                                                    FileEstadoType.fk_Estado = 38;  // estado Indexado 
                                                    FileEstadoType.fk_Usuario = Program.Config.usuario_log;
                                                    FileEstadoType.Fecha_Log = DateTime.Now;

                                                    lock (_lockObject)
                                                    {
                                                        // Validar que exista el file para proceder a actualizarlo o insertarlo                                                      
                                                        if (existFolio)
                                                        {
                                                            dbmCore.SchemaImaging.TBL_File.DBUpdate(fileImgType, (long)itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, fileImageEmail, 1);
                                                            dbmCore.SchemaProcess.TBL_File.DBUpdate(fileProcesType, (long)itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, fileImageEmail);
                                                            dbmCore.SchemaProcess.TBL_File_Estado.DBUpdate(FileEstadoType, (long)itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, fileImageEmail, FileEstadoType.Modulo);
                                                        }
                                                        else
                                                        {                                                        
                                                            manager.CreateItem((long)itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, fileImageEmail, 1, Program.ProyectoImagingRow.Extension_Formato_Imagen_Salida, identifier);
                                                            dbmCore.SchemaImaging.TBL_File.DBInsert(fileImgType);
                                                            dbmCore.SchemaProcess.TBL_File.DBInsert(fileProcesType);
                                                            dbmCore.SchemaProcess.TBL_File_Estado.DBInsert(FileEstadoType);
                                                        }
                                                    }
                                                }

                                                lock (_lockObject)
                                                {
                                                    // Validar que exista el file para proceder a actualizarlo o insertarlo                                                    
                                                    if (existFolio)
                                                    {
                                                        manager.UpdateFolio((long)itemfiltertrackingMail.fk_Expediente, (short)itemfiltertrackingMail.fk_Folder, fileImageEmail, 1, (short)folio, dataImage, dataImageThumbnail);
                                                    }
                                                    else
                                                    {
                                                        manager.CreateFolio((long)itemfiltertrackingMail.fk_Expediente, (short)itemfiltertrackingMail.fk_Folder, fileImageEmail, 1, (short)folio, dataImage, dataImageThumbnail, false);
                                                    }
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
                                                                                                    dataOTServidor.fk_OT);

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
                    Guid guidQueue = Guid.NewGuid();  // Generar un nuevo Guid para el id_Queue

                    // Insertamos en TBL_Queue para enviar email
                    var dataTBLQueueType = new DBTools.SchemaMail.TBL_QueueType
                    {
                        id_Queue = guidQueue,
                        fk_Entidad = itemfiltertrackingMail.fk_Entidad,
                        fk_Usuario = itemfiltertrackingMail.fk_Usuario,
                        Fecha_Queue = currentDate,  // Fecha actual
                        EmailAddress_Queue = itemfiltertrackingMail.EmailAddress_Queue,
                        CC_Queue = itemfiltertrackingMail.CC_Queue,
                        CCO_Queue = itemfiltertrackingMail.CCO_Queue,
                        Subject_Queue = itemfiltertrackingMail.Subject_Queue,
                        Message_Queue = itemfiltertrackingMail.Message_Queue,
                        Attach_Queue = itemfiltertrackingMail.Attach_Queue,
                        AttachName_Queue = itemfiltertrackingMail.AttachName_Queue,
                        EmailFrom = itemfiltertrackingMail.EmailFrom,
                        EmailTracking = true,  // Marcar como correo a ser trackeado
                        EmailFromDisplay = itemfiltertrackingMail.EmailFromDisplay
                    };
                    dbmTools.SchemaMail.TBL_Queue.DBInsert(dataTBLQueueType);


                    // Actualizamos estado y fecha de envío en TBL_Tracking_Mail
                    var dataTBLTrackingMail = new DBTools.SchemaMail.TBL_Tracking_MailType
                    {
                        fk_Queue = guidQueue,
                        fk_Entidad = itemfiltertrackingMail.fk_Entidad,
                        fk_Proyecto = itemfiltertrackingMail.fk_Proyecto,
                        fk_Expediente = itemfiltertrackingMail.fk_Expediente,
                        fk_Folder = itemfiltertrackingMail.fk_Folder,
                        fk_File = itemfiltertrackingMail.fk_File,
                        fk_Usuario = itemfiltertrackingMail.fk_Usuario,
                        Fecha_Log = itemfiltertrackingMail.Fecha_Log,
                        EmailAddress_Queue = itemfiltertrackingMail.EmailAddress_Queue,
                        CC_Queue = itemfiltertrackingMail.CC_Queue,
                        CCO_Queue = itemfiltertrackingMail.CCO_Queue,
                        Subject_Queue = itemfiltertrackingMail.Subject_Queue,
                        Message_Queue = itemfiltertrackingMail.Message_Queue,
                        Attach_Queue = itemfiltertrackingMail.Attach_Queue,
                        AttachName_Queue = itemfiltertrackingMail.AttachName_Queue,
                        EmailFrom = itemfiltertrackingMail.EmailFrom,
                        EmailFromDisplay = itemfiltertrackingMail.EmailFromDisplay,
                        Fecha_Envio = currentDate,
                        fk_Estado_Correo = 2,
                        Detalle_Envio = itemfiltertrackingMail.Detalle_Envio,
                        IsActive = itemfiltertrackingMail.IsActive
                    };
                    dbmTools.SchemaMail.TBL_Tracking_Mail.DBUpdate(dataTBLTrackingMail, itemfiltertrackingMail.id_Tracking_Mail);

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
        }

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
