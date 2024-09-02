using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using EmailSendValidationService.Logs;
using EmailSendValidationService.Servicio;
using Slyg.Tools;

namespace EmailSendValidationService
{
    public partial class Service1 : ServiceBase
    {

        #region " Declaraciones "

        private volatile bool stopRequested = false;     

        private Logs.Logs dataLog = new Logs.Logs();

        #endregion

        #region " Constructores "

        public Service1()
        {
            InitializeComponent();
        }

        #endregion

        protected override void OnStart(string[] args)
        {
            IniciarServicio();
        }

        protected override void OnStop()
        {
            DetenerServicio();
        }

        private void IniciarServicio()
        {
            try
            {
                Program.ConnectionStrings = Program.Config.GetCadenasConexion();

                if (Program.ConnectionStrings.Security == "")
                {
                    dataLog.AddErrorEntry("No se pudo obtener la cadena de conexión a la base de datos Secuity");
                    //dataLog.WriteErrorLog("No se pudo obtener la cadena de conexión a la base de datos Secuity");
                    Stop();
                    return;
                }

                if (Program.ConnectionStrings.Core == "")
                {
                    dataLog.AddErrorEntry("No se pudo obtener la cadena de conexión a la base de datos Core");
                    //dataLog.WriteErrorLog("No se pudo obtener la cadena de conexión a la base de datos Core");
                    Stop();
                    return;
                }

                if (Program.ConnectionStrings.Imaging == "")
                {
                    dataLog.AddErrorEntry("No se pudo obtener la cadena de conexión a la base de datos Imaging");
                    //dataLog.WriteErrorLog("No se pudo obtener la cadena de conexión a la base de datos Imaging");
                    Stop();
                    return;
                }

                if (Program.ConnectionStrings.OCR == "")
                {
                    dataLog.AddErrorEntry("No se pudo obtener la cadena de conexión a la base de datos OCR");
                    //dataLog.WriteErrorLog("No se pudo obtener la cadena de conexión a la base de datos OCR");
                    Stop();
                    return;
                }


                Thread workerThread = new Thread(Proceso);

                stopRequested = false;
                workerThread.Start();

            }
            catch (Exception ex)
            {
                dataLog.AddErrorEntry("Error IniciarServicio ex: " + ex.Message + " " + ex.StackTrace);
                dataLog.WriteErrorLog("Error IniciarServicio ex: " + ex.Message + " " + ex.StackTrace);
                Stop();
            }
        }

        private void DetenerServicio()
        {
            stopRequested = true;      // Solicitar detener el servicio
        }

        private void Proceso()
        {
            try
            {
                while (!stopRequested)
                {
                    try
                    {
                        short EmailScheduledStatus = 1;   // To Do: Cambiar por el estado
                        DBTools.SchemaMail.TBL_Tracking_MailDataTable trakingMailTable = LoadEmailScheduledStatus(EmailScheduledStatus);

                        if (trakingMailTable.Count > 0)
                        {

                            DBTools.SchemaMail.TBL_Tracking_MailDataTable distinctTrakingMailTable = new DBTools.SchemaMail.TBL_Tracking_MailDataTable();

                            // Obtener los valores distintos de fk_Entidad y fk_Proyecto y los almacena en distinctTrakingMailTable
                            trakingMailTable
                                .GroupBy(x => new { x.fk_Entidad, x.fk_Proyecto })
                                .Select(g => g.First())                                          // Seleccionar el primer elemento de cada grupo
                                .OrderBy(x => x.fk_Entidad)
                                .ThenBy(x => x.fk_Proyecto)
                                .ThenBy(x => x.fk_Expediente)
                                .CopyToDataTable()
                                .Rows                                                           // divide en filas
                                .Cast<DataRow>()
                                .ToList()
                                .ForEach(row => distinctTrakingMailTable.ImportRow(row)); // almacena cada fila en una fila de distinctDashboardCapturasTable

                            foreach (var item in distinctTrakingMailTable)
                            {
                                // Validar por entidad y proyecto si se encuentra en dia habil.
                                if (ValidationIsTimeInOperatingHours(item.fk_Entidad, item.fk_Proyecto))
                                {

                                    DBTools.SchemaMail.TBL_Tracking_MailDataTable filtertrackingMail = new DBTools.SchemaMail.TBL_Tracking_MailDataTable();

                                    trakingMailTable
                                        .Where(row => row.fk_Entidad == item.fk_Entidad && row.fk_Proyecto == item.fk_Proyecto )
                                        .OrderBy(x => x.fk_Entidad)
                                        .ThenBy(x => x.fk_Proyecto)
                                        .ThenBy(x => x.fk_Expediente)
                                        .CopyToDataTable()
                                        .Rows                                                           // divide en filas
                                        .Cast<DataRow>()
                                        .ToList()
                                        .ForEach(row => filtertrackingMail.ImportRow(row));    // almacena cada fila en una fila de distinctDashboardCapturasTable

                                    foreach (var itemfiltertrackingMail in filtertrackingMail)
                                    {
                                        // TO DO: Crear el documento del cuerpo del correo enviado por cada fila
                                        var data = false;
                                    }
                                }                                
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        dataLog.AddErrorEntry("Error Proceso ex: " + ex.ToString());
                    }

                    if (stopRequested) return;

                    Thread.Sleep(Program.Config.Intervalo);  // Esperar n segundos antes de continuar
                }
            }
            catch (Exception ex)
            {
                dataLog.AddErrorEntry("**TERMINACIÓN HILO PRINCIPAL**: Se ha producido un error durante la ejecución del hilo principal del servicio. Detalles del error: " + ex.ToString());
                //dataLog.WriteErrorLog("Error Proceso ex: " + ex.ToString());
            }
        }

        private DBTools.SchemaMail.TBL_Tracking_MailDataTable  LoadEmailScheduledStatus(short _OCRCaptureState)
        {
            DBTools.DBToolsDataBaseManager dbmTools = null;

            try
            {
                dbmTools = new DBTools.DBToolsDataBaseManager(Program.ConnectionStrings.Tools);
                dbmTools.Connection_Open();

                // Procede a extraer los registros 
                var trackingMailTable = dbmTools.SchemaMail.TBL_Tracking_Mail.DBFindByfk_Estado_Correo(_OCRCaptureState);

                return trackingMailTable;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (dbmTools != null) dbmTools.Connection_Close();
            }
        }

        private bool ValidationIsTimeInOperatingHours(short _fk_Entidad, short _fk_Proyecto)
        {
            DBTools.DBToolsDataBaseManager dbmTools = null;

            try
            {
                dbmTools = new DBTools.DBToolsDataBaseManager(Program.ConnectionStrings.Tools);
                dbmTools.Connection_Open();

                var dataTable = dbmTools.SchemaConfig.PA_Es_Hora_Habil_Correo.DBExecute(_fk_Entidad, _fk_Proyecto);

                if (dataTable != null && dataTable.Rows.Count > 0)
                {
                    bool isTimeInOperatingHours = Convert.ToBoolean(dataTable.Rows[0][0]);
                    return isTimeInOperatingHours;
                }

                // Si no hay datos, retornar false (o manejar según sea necesario)
                return false;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (dbmTools != null) dbmTools.Connection_Close();
            }
        }


    }
}
