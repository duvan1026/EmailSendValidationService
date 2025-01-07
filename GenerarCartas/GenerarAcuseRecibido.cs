using DBCore;
using DBImaging;
using EmailSendValidationService;
using EmailSendValidationService.Config;
using EmailSendValidationService.Logs;
using Miharu.FileProvider.Library;
using PuppeteerSharp;
using Slyg.Tools;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceEmailSendValidation.GenerarCartas
{
    public class GenerarAcuseRecibido
    {

        #region "Declaraciones"

        protected Logs _DataLog;
        private static readonly object _lockObject = new object();  // Objeto para sincronizar el acceso a escritura DBTools y asegurar ejecución por un solo hilo a la vez.

        const int DeliveredToCarrierStatus = 43;

        Guid identifier;

        #endregion

        #region "Constructor"

        public GenerarAcuseRecibido(Logs dataLog)
        {
            _DataLog = dataLog;
            identifier = Guid.NewGuid();
        }

        #endregion

        #region "Metodos"

        public void GenerarAcuse(DBTools.SchemaMail.TBL_Tracking_MailRow itemfiltertrackingMail)
        {
            DBCoreDataBaseManager dbmCore = null;
            DBImagingDataBaseManager dbmImaging = null;
            DBTools.DBToolsDataBaseManager dbmTools = null;

            try
            {
                dbmImaging = new DBImaging.DBImagingDataBaseManager(Program.ConnectionStrings.Imaging);
                dbmCore = new DBCore.DBCoreDataBaseManager(Program.ConnectionStrings.Core);
                dbmTools = new DBTools.DBToolsDataBaseManager(Program.ConnectionStrings.Tools);

                dbmImaging.Connection_Open(Program.Config.usuario_log);
                dbmCore.Connection_Open(Program.Config.usuario_log);
                dbmTools.Connection_Open();

                var fileProcessDataTable = dbmCore.SchemaProcess.TBL_File.DBFindByfk_Expedientefk_Folderid_File(itemfiltertrackingMail.fk_Expediente,itemfiltertrackingMail.fk_Folder,(SlygNullable<short>)itemfiltertrackingMail.fk_File);
                if (fileProcessDataTable == null || fileProcessDataTable.Count == 0) return;

                var dtDocumentoImagin = dbmImaging.SchemaConfig.CTA_Documento.DBFindByfk_Entidadfk_Proyectoid_DocumentoEliminado(itemfiltertrackingMail.fk_Entidad, itemfiltertrackingMail.fk_Proyecto, fileProcessDataTable[0].fk_Documento, false);
                if (dtDocumentoImagin == null || dtDocumentoImagin.Count == 0) return;

                var dtDocumentoImaginRow = dtDocumentoImagin[0];
                if(dtDocumentoImaginRow.id_Documento_Acuse_Recibido == 0)
                {
                    string message = "El ID del documento 'Acuse Recibido' no está configurado para el documento 'Carta Respuesta' para Entidad : " + itemfiltertrackingMail.fk_Entidad.ToString() + " y proyecto : " + itemfiltertrackingMail.fk_Proyecto.ToString() + ". Verifique la configuración en Parametros Documentos";
                    _DataLog.AddErrorEntry($"¡¡WARNING!! " + message);
                    return;
                }

                var dtFile = dbmCore.SchemaProcess.TBL_File.DBFindByfk_Expedientefk_Folder(itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder);
                if (dtFile == null || dtFile.Count == 0) return;

                // Obtén el valor máximo de id_file
                short maxIdFile = dtFile.Max(file => file.id_File);
                short nextFile = (short)(maxIdFile + 1);

                var fileProcesType = new DBCore.SchemaProcess.TBL_FileType();
                fileProcesType.fk_Expediente = itemfiltertrackingMail.fk_Expediente;
                fileProcesType.fk_Folder = (short)itemfiltertrackingMail.fk_Folder;
                fileProcesType.id_File = nextFile;
                fileProcesType.File_Unique_Identifier = identifier;
                fileProcesType.fk_Documento = dtDocumentoImaginRow.id_Documento_Acuse_Recibido; 
                fileProcesType.Folios_File = 1;
                fileProcesType.Monto_File = 0;
                fileProcesType.CBarras_File = itemfiltertrackingMail.fk_Expediente.ToString() + itemfiltertrackingMail.fk_Folder.ToString() + nextFile;

                var FileEstadoType = new DBCore.SchemaProcess.TBL_File_EstadoType();
                FileEstadoType.fk_Expediente = itemfiltertrackingMail.fk_Expediente;
                FileEstadoType.fk_Folder = (short)itemfiltertrackingMail.fk_Folder;
                FileEstadoType.fk_File = nextFile;
                FileEstadoType.Modulo = new Slyg.Tools.SlygNullable<byte>((byte)ServiceConfig.Modulo.Imaging);
                FileEstadoType.fk_Estado = DeliveredToCarrierStatus;
                FileEstadoType.fk_Usuario = Program.Config.usuario_log;
                FileEstadoType.Fecha_Log = DateTime.Now;

                var mailTrackingType = new DBTools.SchemaMail.TBL_Tracking_MailType();
                mailTrackingType.fk_Estado_Correo = (short)DBTools.EnumEstadosCorreos.EntregadoTransportadora;

                lock (_lockObject)
                {
                    dbmCore.SchemaProcess.TBL_File.DBInsert(fileProcesType);
                    dbmCore.SchemaProcess.TBL_File_Estado.DBInsert(FileEstadoType);
                    dbmTools.SchemaProcess.PA_Insertar_Numero_Guia.DBExecute((SlygNullable<int>)itemfiltertrackingMail.fk_Expediente, itemfiltertrackingMail.fk_Folder, nextFile, dtDocumentoImaginRow.id_Documento_Acuse_Recibido, itemfiltertrackingMail.Numero_Guia);
                    dbmTools.SchemaMail.TBL_Tracking_Mail.DBUpdate(mailTrackingType, itemfiltertrackingMail.id_Tracking_Mail);
                }
            }
            catch 
            {
                throw;
            }
            finally
            {               
                if ((dbmTools != null))
                    dbmTools.Connection_Close();

                if ((dbmImaging != null))
                    dbmImaging.Connection_Close();

                if ((dbmCore != null))
                    dbmCore.Connection_Close();
            }
        }

        #endregion
    }
}
