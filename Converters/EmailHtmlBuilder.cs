using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceEmailSendValidation.Converters
{
    public class EmailHtmlBuilder
    {
        private readonly string emailFrom;
        private readonly string emailSendDate;
        private readonly string emailAddress;
        private readonly string emailCC_CCO;
        private readonly string emailSubject;
        private readonly string emailMessage;

        public EmailHtmlBuilder(string emailFrom, DateTime fechaActual, CultureInfo cultura, string emailAddress, string emailCC_CCO, string emailSubject, string emailMessage)
        {
            this.emailFrom = emailFrom;
            this.emailSendDate = fechaActual.ToString("dddd, d 'de' MMMM 'de' yyyy HH:mm", cultura);
            this.emailAddress = emailAddress;
            this.emailCC_CCO = emailCC_CCO.Trim();  // Elimina los espacios al inicio y al final
            this.emailSubject = emailSubject;
            this.emailMessage = emailMessage;
        }

        public string GenerateHtml()
        {
            var sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang='es'>");
            sb.AppendLine("<head>");
            sb.AppendLine("    <meta charset='UTF-8'>");
            sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
            sb.AppendLine("    <title>Detalles del Correo Electrónico</title>");
            sb.AppendLine("    <style>");
            sb.AppendLine("        body {");
            sb.AppendLine("            font-family: Arial, sans-serif;");
            sb.AppendLine("            margin: 0;");
            sb.AppendLine("            padding: 20px;");
            sb.AppendLine("            color: #333;");
            sb.AppendLine("        }");
            sb.AppendLine("        .container {");
            sb.AppendLine("            max-width: 800px;");
            sb.AppendLine("            margin: auto;");
            sb.AppendLine("            padding: 20px;");
            sb.AppendLine("            border: 1px solid #ddd;");
            sb.AppendLine("            border-radius: 8px;");
            sb.AppendLine("            background-color: #f9f9f9;");
            sb.AppendLine("        }");
            sb.AppendLine("        .header {");
            sb.AppendLine("            font-size: 24px;");
            sb.AppendLine("            font-weight: bold;");
            sb.AppendLine("            margin-bottom: 10px;");
            sb.AppendLine("            border-bottom: 2px solid #ddd;");
            sb.AppendLine("            padding-bottom: 10px;");
            sb.AppendLine("        }");
            sb.AppendLine("        .section {");
            sb.AppendLine("            margin-bottom: 5px;");
            sb.AppendLine("        }");
            sb.AppendLine("        .section strong {");
            sb.AppendLine("            margin-right: 5px;");
            sb.AppendLine("            margin-bottom: 5px;");
            sb.AppendLine("        }");
            sb.AppendLine("        .section p {");
            sb.AppendLine("            display: inline;");
            sb.AppendLine("            margin: 0;");
            sb.AppendLine("            font-size: 14px;");
            sb.AppendLine("        }");
            sb.AppendLine("    </style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("    <div class='container'>");
            sb.AppendLine("        <div class='header'>");
            sb.AppendLine($"            <strong>Asunto:</strong> {emailSubject}");
            sb.AppendLine("        </div>");
            sb.AppendLine("        <div class='section'>");
            sb.AppendLine($"            <strong>De:</strong>");
            sb.AppendLine($"            <p>{emailFrom}</p>");
            sb.AppendLine("        </div>");
            sb.AppendLine("        <div class='section'>");
            sb.AppendLine($"            <strong>Enviado:</strong>");
            sb.AppendLine($"            <p>{emailSendDate}</p>");
            sb.AppendLine("        </div>");
            sb.AppendLine("        <div class='section'>");
            sb.AppendLine($"            <strong>Para:</strong>");
            sb.AppendLine($"            <p>{emailAddress}</p>");
            sb.AppendLine("        </div>");

            // Conditionally add the CC and CCO section
            if (!string.IsNullOrEmpty(emailCC_CCO))
            {
                sb.AppendLine("        <div class='section'>");
                sb.AppendLine($"            <strong>CC y CCO:</strong>");
                sb.AppendLine($"            <p>{emailCC_CCO}</p>");
                sb.AppendLine("        </div>");
            }

            sb.AppendLine("        <div class='section'>");
            sb.AppendLine($"            <strong>Asunto:</strong>");
            sb.AppendLine($"            <p>{emailSubject}</p>");
            sb.AppendLine("        </div>");
            sb.AppendLine("        <div class='section'>");
            sb.AppendLine("            <br>");
            sb.AppendLine("            <br>");
            sb.AppendLine($"            <p>{emailMessage}</p>");
            sb.AppendLine("        </div>");
            sb.AppendLine("    </div>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb.ToString();
        }
    }
}
