using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Remoting.Messaging;
//using static System.Net.Mime.MediaTypeNames;

namespace EmailSendValidationService.Logs
{
    public class Logs
    {
        private static StreamWriter mswSw = null;
        private static int Instancias;

        public Logs()
        {
            if (Instancias == 0)
                AbrirFichero();

            Instancias++;
        }

        protected void AbrirFichero()
        {
            String NombreFichero = GetNombreFichero();

            if (!(new FileInfo(NombreFichero).Exists))
            {
                if (mswSw != null)
                {
                    mswSw.Close();
                    mswSw.Dispose();
                }
                mswSw = new StreamWriter(NombreFichero, true);
            }
            else
            {
                if (mswSw == null)
                {
                    mswSw = new StreamWriter(NombreFichero, true);
                }
            }
        }

        protected string GetNombreFichero()
        {
            String Fichero = Program.AppDataPath;
            //if (!Directory.Exists(Fichero)) Directory.CreateDirectory(Fichero);

            //Fichero += "\\log.txt";
            return Fichero;
        }

        public void AddErrorEntry(string Mensaje)
        {
            Escribir("ERROR", Mensaje);
        }

        public void AddWarningEntry(string Mensaje)
        {
            Escribir("WARNING", Mensaje);
        }

        public void AddInformationEntry(string Mensaje)
        {
            Escribir("INFORMATION", Mensaje);
        }

        protected void Escribir(string Tipo, string Mensaje)
        {
            lock (this)
            {
                AbrirFichero();
                mswSw.WriteLine("--------------------------------------------------------------");
                mswSw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                mswSw.WriteLine("Mensaje: " + Mensaje);
                mswSw.WriteLine("--------------------------------------------------------------");
                mswSw.WriteLine("");
                //mswSw.WriteLine("" + DateTime.Now.ToLongTimeString() + " [" + Tipo + "] " + Mensaje);
                mswSw.Flush();
            }
        }
        public void Dispose()
        {
            lock (this)
            {
                if (--Instancias == 0)
                {
                    mswSw.Close();
                    mswSw.Dispose();
                }
            }
        }

        public void WriteErrorLog(string nMessage)
        {
            //lock (objectLock)
            //{
            try
            {
                //JWriteLog("WriteErrorLog Path: " + Program.AppDataPath + "log.txt", EventLogEntryType.Information);

                //JWriteLog(nMessage, EventLogEntryType.Error);

                //using (StreamWriter sw = new StreamWriter(Program.AppDataPath + "log.txt", true))
                using (StreamWriter sw = new StreamWriter(Program.AppDataPath, true))
                {
                    sw.WriteLine("--------------------------------------------------------------");
                    sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    sw.WriteLine("Mensaje: " + nMessage);
                    sw.WriteLine("--------------------------------------------------------------");
                    sw.WriteLine("");
                }
            }
            catch
            {
                //try { JWriteLog(ex.Message, EventLogEntryType.Error); } catch { }
            }
            //}
        }

        //public void JWriteLog(string mensaje, EventLogEntryType tipo)
        //{
        //    if (!EventLog.SourceExists("SofTracService"))
        //    {
        //        EventLog.CreateEventSource("SofTracService", "Application");
        //    }

        //    using (EventLog eventLog1 = new EventLog())
        //    {
        //        eventLog1.Source = "SofTracService";
        //        eventLog1.WriteEntry(mensaje, tipo);
        //    }
        //}

    }
}
