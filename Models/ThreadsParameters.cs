using PuppeteerSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceEmailSendValidation.Models
{
    public class ThreadsParameters
    {
        public DBTools.SchemaMail.TBL_Tracking_MailRow TransferenciaRow { get; set; }
        public IBrowser Browser { get; set; }

        public ThreadsParameters(DBTools.SchemaMail.TBL_Tracking_MailRow transferenciaRow, IBrowser browser)
        {
            TransferenciaRow = transferenciaRow;
            Browser = browser;
        }
    }
}
