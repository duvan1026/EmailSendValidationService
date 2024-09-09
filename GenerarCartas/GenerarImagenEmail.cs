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

namespace ServiceEmailSendValidation.GenerarCartas
{
    public class GenerarImagenEmail
    {
        #region "Declaraciones"

        protected const int MaxThumbnailWidth = 60;
        protected const int MaxThumbnailHeight = 80;

        protected int _ImageCount;
        protected Logs _DataLog;
        protected short _FormatoImagenSalida;
        protected string _ExtensionFormatoImagenSalida;

        protected DBCore.SchemaImaging.TBL_FileRow _FileImagingRow;

        #endregion

        #region "Constructor"

        public GenerarImagenEmail(Logs dataLog)
        {
            _DataLog = dataLog;            
        }

        #endregion

        #region "Metodos"
        public void GenerarCartas(DBTools.SchemaMail.TBL_Tracking_MailDataTable filtertrackingMail)
        {
            DBIntegrationDataBaseManager dbmIntegration = null;
            DBCoreDataBaseManager dbmCore = null;
            DBImagingDataBaseManager dbmImaging = null;
            FileProviderManager manager = null;

            try
            {
                dbmImaging = new DBImaging.DBImagingDataBaseManager(Program.ConnectionStrings.Imaging);
                dbmIntegration = new DBIntegration.DBIntegrationDataBaseManager(Program.ConnectionStrings.Integration);
                dbmCore = new DBCore.DBCoreDataBaseManager(Program.ConnectionStrings.Core);

                dbmImaging.Connection_Open(Program.Config.usuario_log);
                dbmIntegration.Connection_Open(Program.Config.usuario_log);
                dbmCore.Connection_Open(Program.Config.usuario_log);

                var previousfkEntidad = 0;
                var previousfkProyecto = 0;
                var dtProyecto = new DBImaging.SchemaConfig.CTA_ProyectoDataTable();
                var OTDataTable = new DBImaging.SchemaProcess.CTA_OT_Servidor_Centro_ProcesamientoDataTable();

                foreach (var itemfiltertrackingMail in filtertrackingMail)
                {
                    if (previousfkEntidad != itemfiltertrackingMail.fk_Entidad || previousfkProyecto != itemfiltertrackingMail.fk_Proyecto)
                    {
                        dtProyecto = dbmImaging.SchemaConfig.CTA_Proyecto.DBFindByfk_Entidadfk_Proyecto(itemfiltertrackingMail.fk_Entidad, itemfiltertrackingMail.fk_Proyecto);
                        OTDataTable = dbmImaging.SchemaProcess.CTA_OT_Servidor_Centro_Procesamiento.DBFindByfk_Entidadfk_Proyecto(itemfiltertrackingMail.fk_Entidad, itemfiltertrackingMail.fk_Proyecto);
                    }

                    if (dtProyecto.Count > 0)
                    {
                        Program.ProyectoImagingRow = dtProyecto[0].ToCTA_ProyectoSimpleType();
                        _ExtensionFormatoImagenSalida = Program.ProyectoImagingRow.Extension_Formato_Imagen_Salida;
                        _FormatoImagenSalida = Program.ProyectoImagingRow.id_Formato_Imagen_Salida;

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

                            // TO DO: gestionar en caso de que no se capture el expediente folder
                            var firstRowOTData = filterOTDataTable[0];

                            var formato = Utilities.GetEnumFormat(Program.ProyectoImagingRow.id_Formato_Imagen_Entrada.ToString());
                            var compresion = Utilities.GetEnumCompression((DesktopConfig.FormatoImagenEnum)_FormatoImagenSalida);

                            var servidor = dbmImaging.SchemaCore.CTA_Servidor.DBFindByfk_Entidadid_Servidor(firstRowOTData.fk_Entidad_Servidor, firstRowOTData.fk_Servidor)[0].ToCTA_ServidorSimpleType();
                            var centro = dbmImaging.SchemaSecurity.CTA_Centro_Procesamiento.DBFindByfk_Entidadfk_Sedeid_Centro_Procesamiento(firstRowOTData.fk_Entidad_Procesamiento, firstRowOTData.fk_Sede_Procesamiento_Cargue, firstRowOTData.fk_Centro_Procesamiento_Cargue)[0].ToCTA_Centro_ProcesamientoSimpleType();
                            manager = new FileProviderManager(servidor, centro, ref dbmImaging, Program.Config.usuario_log);
                            manager.Connect();

                            string FileName = "";
                            string htmlFilePath = "";
                            byte[] Imagen = null;
                            byte[] Thumbnail = null;
                            var identificador = Guid.NewGuid();

                            string fullPath = Program.AppPath + Program.TempPath;
                            if (!Directory.Exists(fullPath)) Directory.CreateDirectory(fullPath);

                            htmlFilePath = fullPath + identificador.ToString() + ".html";
                            FileName = fullPath + identificador.ToString() + ".tif";

                            string _CC_Queue = "";
                            if (itemfiltertrackingMail.CC_Queue != "" && itemfiltertrackingMail.CC_Queue != ";")
                            {
                                _CC_Queue = itemfiltertrackingMail.CC_Queue;
                            }
                            string _CCO_Queue = "";
                            if (itemfiltertrackingMail.CCO_Queue != "" && itemfiltertrackingMail.CCO_Queue != ";")
                            {
                                _CCO_Queue = " " + itemfiltertrackingMail.CCO_Queue;
                            }

                            string dataCC_CCO = _CC_Queue + _CC_Queue;

                            var emailFrom = itemfiltertrackingMail.EmailFrom;
                            var emailSendDate = DateTime.Now;
                            var cultura = CultureInfo.GetCultureInfo("es-ES");
                            var emailAddress = itemfiltertrackingMail.EmailAddress_Queue;
                            var emailCC_CCO = dataCC_CCO;
                            var emailSubject = itemfiltertrackingMail.Subject_Queue;
                            var emailMessage = itemfiltertrackingMail.Message_Queue;

                            var emailHtmlBuilder = new EmailHtmlBuilder(emailFrom, emailSendDate, cultura, emailAddress, emailCC_CCO, emailSubject, emailMessage);
                            string htmlContent = emailHtmlBuilder.GenerateHtml();

                            File.WriteAllText(htmlFilePath, htmlContent);

                            var converter = new HtmlToTiffConverter(fullPath);                                     // Instanciar el convertidor                            
                            converter.ConvertHtmlToTiffAsync(htmlFilePath, FileName).GetAwaiter().GetResult(); ;   // Convertir HTML a TIFF
                            
                            File.Delete(htmlFilePath);

                            short folios = (short)ImageManager.GetFolios(FileName);
                            var FolioBitmap = ImageManager.GetFolioBitmap(FileName, folios);
                            short fileImageEmail = (short)(itemfiltertrackingMail.fk_File + 1);

                            _ImageCount = GetImageCount(manager, itemfiltertrackingMail);

                            for (int folio = 1; folio <= (folios + _ImageCount); folio++)
                            {
                                byte[] dataImage = null;
                                byte[] dataImageThumbnail = null;

                                if (folio <= folios)
                                {
                                    dataImage = ImageManager.GetFolioData(FileName, folio, formato, compresion);
                                    dataImageThumbnail = ImageManager.GetThumbnailData(FileName, folio, folio, MaxThumbnailWidth, MaxThumbnailHeight)[0];
                                }
                                else
                                {
                                    short currentFolio = (short)(folio - folios);
                                    manager.GetFolio(itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, (short)itemfiltertrackingMail.fk_File, _FileImagingRow.id_Version, currentFolio, ref dataImage, ref dataImageThumbnail);
                                }

                                //To Do: Eliminar ////////////////////////
                                string FileName2 = fullPath + "Copia_" + folio.ToString() + ".tif";
                                using (var fs = new FileStream(FileName2, FileMode.Create))
                                {
                                    fs.Write(dataImage, 0, dataImage.Length);
                                    fs.Close();
                                }//////////////////////////////////////////


                                if (folio == 1)
                                {
                                    manager.CreateItem((long)itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, fileImageEmail, 1, _ExtensionFormatoImagenSalida, identificador);

                                    var fileImgType = new DBCore.SchemaImaging.TBL_FileType();

                                    fileImgType.fk_Expediente = itemfiltertrackingMail.fk_Expediente;
                                    fileImgType.fk_Folder = (short)itemfiltertrackingMail.fk_Folder;
                                    fileImgType.fk_File = fileImageEmail;
                                    fileImgType.id_Version = 1;
                                    fileImgType.File_Unique_Identifier = identificador;
                                    fileImgType.Folios_Documento_File = (short)(folios + _ImageCount);
                                    fileImgType.Tamaño_Imagen_File = 0;
                                    fileImgType.Nombre_Imagen_File = "";
                                    fileImgType.Key_Cargue_Item = "";
                                    fileImgType.Save_FileName = "";
                                    fileImgType.fk_Content_Type = _ExtensionFormatoImagenSalida;
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
                                    dbmCore.SchemaImaging.TBL_File.DBInsert(fileImgType);
                                }

                                manager.CreateFolio((long)itemfiltertrackingMail.fk_Expediente, (short)itemfiltertrackingMail.fk_Folder, fileImageEmail, 1, (short)folio, dataImage, dataImageThumbnail, false);
                            }

                            File.Delete(FileName);                            
                        }
                    }

                    previousfkEntidad = itemfiltertrackingMail.fk_Entidad;
                    previousfkProyecto = itemfiltertrackingMail.fk_Proyecto;
                }
            }
            catch (Exception ex)
            {
                if (dbmCore != null)
                    dbmCore.Transaction_Rollback();
                if (manager != null)
                    manager.TransactionRollback();
                throw;
            }
            finally
            {
                if ((dbmImaging != null))
                    dbmImaging.Connection_Close();
                if ((dbmCore != null))
                    dbmCore.Connection_Close();
                if ((dbmIntegration != null))
                    dbmIntegration.Connection_Close();
            }
        }

        #endregion

        #region "Funciones"


        /// <summary>
        /// Obtiene el numero de folios de un file
        /// </summary>
        /// <param name="nManager"></param>
        /// <param name="itemfiltertrackingMail">fila de archivos con los datos para extraer los folios</param>
        /// <returns></returns>
        private int GetImageCount(FileProviderManager nManager, DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail)
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
            catch
            {
                throw;
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
