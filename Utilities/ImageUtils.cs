using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceEmailSendValidation.Utilities
{
    public static class ImageUtils
    {

        /// <summary>
        /// Convierte una imagen en binario a una cadena Base64 en formato HTML <img>.
        /// </summary>
        /// <param name="imageBinary">Arreglo de bytes de la imagen.</param>
        /// <returns>Etiqueta HTML <img> con la imagen en formato Base64.</returns>
        public static string ConvertImageToBase64Html(byte[] imageBinary)
        {
            string imageFormat = GetImageFormat(imageBinary)
            string base64String = Convert.ToBase64String(imageBinary);                          // Convierte el arreglo de bytes (la imagen) en una cadena Base64
            string htmlImage = $"<img src='data:image/{imageFormat};base64,{base64String}' />"; //cadena en formato HTML img con el tipo de la imagen (jpeg, png, etc.)

            return htmlImage;
        }

        /// <summary>
        /// Convierte un arreglo de bytes en una cadena Base64.
        /// </summary>
        /// <param name="imageBinary">Arreglo de bytes de la imagen.</param>
        /// <returns>Cadena en Base64 que representa los datos binarios.</returns>
        public static string ConvertBytesToBase64(byte[] imageBinary)
        {
            // Convierte el arreglo de bytes en una cadena Base64 y retorna la cadena
            return Convert.ToBase64String(imageBinary);
        }

        /// <summary>
        /// Determina el formato de la imagen a partir de los primeros bytes.
        /// </summary>
        /// <param name="imageBinary">Arreglo de bytes de la imagen.</param>
        /// <returns>Formato de la imagen (jpeg, png, bmp, gif).</returns>
        public static string GetImageFormat(byte[] imageBinary)
        {
            // Usa los primeros bytes de la imagen para determinar su formato
            if (imageBinary.Length > 4)
            {
                // Verifica los formatos comunes de imagen
                if (imageBinary[0] == 0xFF && imageBinary[1] == 0xD8) return "jpeg"; // JPEG
                if (imageBinary[0] == 0x89 && imageBinary[1] == 0x50) return "png";  // PNG
                if (imageBinary[0] == 0x42 && imageBinary[1] == 0x4D) return "bmp";  // BMP
                if (imageBinary[0] == 0x47 && imageBinary[1] == 0x49) return "gif";  // GIF
            }

            // Si no se identifica el formato, retorna un valor por defecto (en este caso, jpeg)
            return "jpeg";
        }

    }
}
