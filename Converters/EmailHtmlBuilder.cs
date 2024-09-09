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
            this.emailCC_CCO = emailCC_CCO;
            this.emailSubject = emailSubject;
            this.emailMessage = emailMessage;
        }

        public string GenerateHtml()
        {
            return $@"
<!DOCTYPE html>
<html lang='es'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Detalles del Correo Electrónico</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            color: #333;
        }}
        .container {{
            max-width: 800px;
            margin: auto;
            padding: 20px;
            border: 1px solid #ddd;
            border-radius: 8px;
            background-color: #f9f9f9;
        }}
        .header {{
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 10px;
            border-bottom: 2px solid #ddd;
            padding-bottom: 10px;
        }}
        .section {{
            margin-bottom: 5px;
        }}
        .section strong {{
            margin-right: 5px; /* Espacio entre el título y el valor */
            //display: block;
            margin-bottom: 5px;
            //font-size: 16px;
        }}
        .section p {{
            display: inline;
            margin: 0;
            font-size: 14px;
            //flex: 1; /* Esto asegura que el texto se alinee a la derecha del encabezado */
        }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <strong>Asunto:</strong> {emailSubject}
        </div>
        <div class='section'>
            <strong>De:</strong>
            <p>{emailFrom}</p>
        </div>
        <div class='section'>
            <strong>Enviado:</strong>
            <p>{emailSendDate}</p>
        </div>
        <div class='section'>
            <strong>Para:</strong>
            <p>{emailAddress}</p>
        </div>
        <div class='section'>
            <strong>CC y CCO:</strong>
            <p>{emailCC_CCO}</p>
        </div>
        <div class='section'>
            <strong>Asunto:</strong>
            <p>{emailSubject}</p>
        </div>
        <div class='section'>
            <br>
            <br>
            <p>{emailMessage}</p>
        </div>
    </div>
</body>
</html>";
        }
    }
}
