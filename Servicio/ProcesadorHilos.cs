using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EmailSendValidationService.Servicio
{
    public class ProcesadorHilos
    {
        #region " Declaraciones "

        private int numHilos;
        private List<Thread> ListaHilos = new List<Thread>();
        public static object objLock = new object();

        public Service1 servicio;

        #endregion

        public ProcesadorHilos()
        {
            numHilos = 8;//Program.ConnectionParametersStrings.threadCount;
        }

        #region

        #endregion

        #region " Metodos "
        //public void AgregarHilo(DBImaging.SchemaConfig.CTA_Get_OCR_Captura_File_TypedRow nTransferenciarow, List<DocumentFieldsFiltersParameters> listFiledsFiltersParameters)
        //{
        //    if (!TieneHiloslibres())
        //    {
        //        do
        //        {
        //            Thread.Sleep(100);
        //            if (TieneHiloslibres())
        //            {
        //                break;
        //            }
        //        } while (true);
        //    }

        //    Thread Threads = new Thread(new ParameterizedThreadStart(servicio.ProcesoPrincipalHilo)); // especifica el metodo que será ejecutado por el hilo
        //    Threads.Start(Tuple.Create(nTransferenciarow, listFiledsFiltersParameters));

        //    lock (objLock)
        //    {
        //        ListaHilos.Add(Threads);
        //    }
        //}

        #endregion

        #region " Funciones "
        public bool TieneHiloslibres()
        {
            lock (objLock)
            {
                List<Thread> ListaHilosBorrar = new List<Thread>();
                foreach (Thread hilo in ListaHilos)
                {
                    if (hilo.ThreadState == ThreadState.Stopped)
                    {
                        ListaHilosBorrar.Add(hilo);
                    }
                }
                foreach (Thread hilo in ListaHilosBorrar)
                {
                    ListaHilos.Remove(hilo);
                }
            }
            return ListaHilos.Count < numHilos;
        }

        public bool TerminoHilos()
        {
            foreach (Thread hilo in ListaHilos)
            {
                if (hilo.ThreadState != ThreadState.Stopped)
                {
                    return false;
                }
            }
            return true;
        }

        #endregion

    }
}
