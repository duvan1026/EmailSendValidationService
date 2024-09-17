using ImageMagick;
using PuppeteerSharp.Media;
using PuppeteerSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace ServiceEmailSendValidation.Converters
{
    public class HtmlToTiffConverter
    {
        #region "Declaraciones"

        private readonly string _outputFolderPath;
        private readonly int _pageWidth;
        private readonly int _pageHeight;
        private readonly int _defaultMarginTop;

        #endregion

        #region "Constructores"

        public HtmlToTiffConverter(string outputFolderPath, int pageWidth, int pageHeight, int defaultMarginTop)
        {
            _outputFolderPath = outputFolderPath;
            _pageWidth = pageWidth;
            _pageHeight = pageHeight;
            _defaultMarginTop = defaultMarginTop;
        }

        #endregion

        #region "Metodos"

        public async Task ConvertHtmlToTiffAsync(string htmlFilePath, string outputTiffPath)
        {
            try
            {
                // Descarga y configuración de Puppeteer
                var browserFetcher = new BrowserFetcher();
                await browserFetcher.DownloadAsync();
                //var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
                var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = false });  // modo para renderizar en servicio windows

                // Crear una nueva página y cargar el contenido HTML
                var page = await browser.NewPageAsync();
                var htmlContent = File.ReadAllText(htmlFilePath);
                await page.SetContentAsync(htmlContent);

                // Esperar a que el cuerpo de la página esté completamente renderizado
                await page.WaitForSelectorAsync("body");

                // Asegurar que todas las imágenes se hayan cargado completamente
                await page.EvaluateFunctionAsync(@"() => {
                    return new Promise((resolve) => {
                        const checkIfImagesLoaded = () => {
                            if (Array.from(document.images).every(img => img.complete)) {
                                resolve();
                            } else {
                                setTimeout(checkIfImagesLoaded, 100);
                            }
                        };
                        checkIfImagesLoaded();
                    });
                }");


                // Establecer el tamaño de la vista de la página
                await page.SetViewportAsync(new ViewPortOptions
                {
                    Width = _pageWidth,
                    Height = _pageHeight 
                });


                /// DEBUG
                //// Capturar la imagen de la página completa
                //var screenshotPath = Path.Combine(_outputFolderPath, $"{Path.GetFileNameWithoutExtension(outputTiffPath + "_test")}.png");
                //await page.ScreenshotAsync(screenshotPath, new ScreenshotOptions
                //{
                //    FullPage = true  // Captura toda la página
                //});

                // Capturar las imágenes por partes
                var tempImagePaths = await CapturePageScreenshotsAsync((Page)page);

                // Convertir las imágenes en un archivo TIFF
                CreateTiffFromImages(tempImagePaths, outputTiffPath);

                // Eliminar las imágenes temporales
                CleanupTemporaryImages(tempImagePaths); 

                await browser.CloseAsync();
            }
            catch
            {
                throw;
            }

        }

        private void EnsureOutputDirectoryExists()
        {
            if (!Directory.Exists(_outputFolderPath))
            {
                Directory.CreateDirectory(_outputFolderPath);
            }
        }

        private void CreateTiffFromImages(List<string> imagePaths, string outputTiffPath)
        {
            try
            {
                using (var tiff = new MagickImageCollection())
                {
                    foreach (var imagePath in imagePaths)
                    {
                        tiff.Add(imagePath);
                    }

                    tiff.Write(outputTiffPath);
                }
            }
            catch
            {
                throw;
            }
        }

        private void CleanupTemporaryImages(List<string> imagePaths)
        {
            foreach (var imagePath in imagePaths)
            {
                if (File.Exists(imagePath))
                {
                    File.Delete(imagePath);
                }
            }
        }

        #endregion

        #region "Funciones"

        private async Task<List<string>> CapturePageScreenshotsAsync(Page page)
        {
            var tempImagePaths = new List<string>();
            var tempImageCorrectPaths = new List<string>();
            int currentPage = 1;
            bool hasMoreContent = true;
            Guid identifier = Guid.NewGuid();

            // Obtener la posición superior e inferior de todas las líneas de texto visibles en la página
            var linePositions = await page.EvaluateFunctionAsync<List<Dictionary<string, int>>>(@"
                 () => {
                     const lines = [];
                     const elements = document.querySelectorAll('p, h1, h2, h3, div, span, td, th, table, li, a, b, strong, i, em'); // Seleccionar más elementos de texto
                     elements.forEach(el => {
                         if (el.innerText.trim() !== '') { // Asegurarse de que el elemento tiene texto visible
                             const rect = el.getBoundingClientRect();
                             lines.push({
                                 top: Math.floor(rect.top + window.scrollY),   // Posición Y superior
                                 bottom: Math.ceil(rect.bottom + window.scrollY) // Posición Y inferior
                             });
                         }
                     });
                     return lines;
                 }
             ");

            int previousPageHeight = 0;
            var bottomValue = linePositions[0]["bottom"];

            while (hasMoreContent)
            {
                // Determinar el recorte basado en la última línea de texto visible dentro de la altura de la página
                int pageHeightAdjusted = _pageHeight; // Ajustar si es necesario
                int lastVisibleLineY = linePositions
                    .Where(line => line["bottom"] <= (currentPage * pageHeightAdjusted))
                    .Select(line => line["bottom"])
                    .LastOrDefault();

                if (lastVisibleLineY == 0) lastVisibleLineY = currentPage * pageHeightAdjusted;

                var screenshotPath = Path.Combine(_outputFolderPath, $"{identifier}_{currentPage}.png");        // Definir la ruta para la captura de pantalla

                var bodyHeight = await page.EvaluateExpressionAsync<double>(@"document.body.scrollHeight"); // Obtiene el valor mayor de hight renderizado
                bool isBodyMax = bodyHeight <= (_pageHeight * currentPage);                                 // Evaluar si hay más contenido para capturar  

                // Calcular la altura y posición de la página a capturar
                int pageHeight = (currentPage == 1) ? lastVisibleLineY : lastVisibleLineY - previousPageHeight;
                pageHeight = (!isBodyMax) ? pageHeight : (bottomValue - previousPageHeight);
                int pageY = (currentPage == 1) ? 0 : previousPageHeight;

                // Actualizar la altura de la página previa
                previousPageHeight = (currentPage == 1) ? pageHeight : (previousPageHeight + pageHeight);

                // Capturar la imagen de la vista actual
                await page.ScreenshotAsync(screenshotPath, new ScreenshotOptions
                {
                    Clip = new Clip
                    {
                        X = 0,
                        Y = pageY,
                        Width = _pageWidth,
                        Height = pageHeight
                    },
                    FullPage = false
                });

                tempImagePaths.Add(screenshotPath);

                // Evaluar si hay más contenido para capturar
                if (bodyHeight <= (_pageHeight * currentPage))
                {
                    hasMoreContent = false;
                }
                else
                {
                    await page.EvaluateExpressionAsync($"window.scrollTo(0, {lastVisibleLineY});");
                    currentPage++;
                }
            }

            // Ajustar las imágenes de acuerdo con los márgenes de las páginas
            tempImageCorrectPaths = AdjustCapturedImagesWithMargins(tempImagePaths);

            return tempImageCorrectPaths;
        }

        private List<string> AdjustCapturedImagesWithMargins(List<string> tempImagePaths)
        {
            var tempImageCorrectPaths = new List<string>();

            // Ajustar la primera imagen para agregar espacio en blanco en la parte inferior
            if (tempImagePaths.Count > 0)
            {
                var firstImagePath = tempImagePaths[0];
                string newImagePath = AdjustImageForMargin(firstImagePath, isBottom: true);
                tempImageCorrectPaths.Add(newImagePath);
            }

            // Ajustar las imágenes intermedias para agregar espacio en blanco en la parte superior e inferior
            for (int i = 1; i < tempImagePaths.Count - 1; i++)
            {
                var imagePath = tempImagePaths[i];
                string newImagePath = AdjustImageForMargin(imagePath, isBottom: true, isTop: true);
                tempImageCorrectPaths.Add(newImagePath);
            }

            // Ajustar la última imagen para agregar espacio en blanco en la parte superior
            if (tempImagePaths.Count > 1)
            {
                var lastImagePath = tempImagePaths.Last();
                string newImagePath = AdjustImageForMargin(lastImagePath, isBottom: false);
                tempImageCorrectPaths.Add(newImagePath);
            }

            CleanupTemporaryImages(tempImagePaths);

            return tempImageCorrectPaths;
        }

        private string AdjustImageForMargin(string imagePath, bool isBottom, bool isTop = false)
        {
            try
            {
                using (var image = System.Drawing.Image.FromFile(imagePath))
                {
                    // Definir el tamaño de la nueva imagen
                    var newWidth = image.Width;
                    int margin = (isBottom && isTop) ? ((_pageHeight - image.Height) / 2) : (_pageHeight - image.Height);

                    using (var newImage = new Bitmap(newWidth, _pageHeight))
                    using (var graphics = Graphics.FromImage(newImage))
                    {
                        // Rellenar con blanco
                        graphics.Clear(Color.White);

                        int x = 0;
                        int y = 0;

                        // Calcular la posición de la imagen original en la nueva imagen
                        if (isTop)
                        {
                            y += margin; // Ajustar la posición si se agrega margen en la parte superior
                        }
                        else if (!isBottom && !isTop)
                        {
                            y = _defaultMarginTop; // To Do: valor por defecto para dejar espacio en Top ultima Pagina
                        }

                        // Dibujar la imagen original sobre la nueva imagen
                        graphics.DrawImage(image, x, y);

                        string newImagePath = Path.Combine(Path.GetDirectoryName(imagePath), "modified_" + Path.GetFileName(imagePath));

                        newImage.Save(newImagePath, System.Drawing.Imaging.ImageFormat.Png);
                        return newImagePath;
                    }
                }
            }
            catch
            {
                throw;
            }
        }

        /* Original 
        private async Task<List<string>> CapturePageScreenshotsAsync(Page page)
        {
            try
            {
                var tempImagePaths = new List<string>();
                var tempImageCorrectPaths = new List<string>();
                int currentPage = 1;
                bool hasMoreContent = true;

                const int margin = 20; // Ajusta este valor según sea necesario

                while (hasMoreContent)
                {
                    var screenshotPath = Path.Combine(_outputFolderPath, $"temp_page{currentPage}.png");

                    // Capturar la imagen de la vista actual
                    await page.ScreenshotAsync(screenshotPath, new ScreenshotOptions
                    {
                        Clip = new Clip
                        {
                            X = 0,
                            Y = (currentPage - 1) * (_pageHeight - margin),//_pageHeight,
                            Width = _pageWidth,
                            Height = _pageHeight - margin
                        },
                        FullPage = false
                    });

                    tempImagePaths.Add(screenshotPath);

                    // Evaluar si hay más contenido para capturar
                    var bodyHeight = await page.EvaluateExpressionAsync<double>(@"document.body.scrollHeight");
                    if (bodyHeight <= _pageHeight * currentPage)
                    {
                        hasMoreContent = false;
                    }
                    else
                    {
                        await page.EvaluateExpressionAsync(@"window.scrollBy(0, window.innerHeight)");
                        currentPage++;
                    }
                }

                // Ajustar la primera imagen para agregar espacio en blanco en la parte inferior
                if (tempImagePaths.Count > 0)
                {
                    var firstImagePath = tempImagePaths[0];
                    string newImagePath = AdjustImageForMargin(firstImagePath, margin, isBottom: true);
                    tempImageCorrectPaths.Add(newImagePath);
                }

                // Ajustar las imágenes intermedias para agregar espacio en blanco en la parte superior e inferior
                for (int i = 1; i < tempImagePaths.Count - 1; i++)
                {
                    var imagePath = tempImagePaths[i];
                    string newImagePath = AdjustImageForMargin(imagePath, margin, isBottom: true, isTop: true);
                    tempImageCorrectPaths.Add(newImagePath);
                }

                // Ajustar la última imagen para agregar espacio en blanco en la parte superior
                if (tempImagePaths.Count > 1)
                {
                    var lastImagePath = tempImagePaths.Last();
                    string newImagePath = AdjustImageForMargin(lastImagePath, margin, isBottom: false);
                    tempImageCorrectPaths.Add(newImagePath);
                }

                CleanupTemporaryImages(tempImagePaths);

                return tempImageCorrectPaths;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }*/

        /* Original
        private string AdjustImageForMargin(string imagePath, int margin, bool isBottom, bool isTop = false)
        {
            try
            {
                using (var image = System.Drawing.Image.FromFile(imagePath))
                {
                    // Definir el tamaño de la nueva imagen
                    var newWidth = image.Width;
                    var newHeight = image.Height + margin;

                    if (isTop)
                    {
                        newHeight += margin; // Agregar margen en la parte superior
                    }
                    if (isBottom)
                    {
                        newHeight += margin; // Agregar margen en la parte inferior
                    }

                    using (var newImage = new Bitmap(newWidth, newHeight))
                    using (var graphics = Graphics.FromImage(newImage))
                    {
                        // Rellenar con blanco
                        graphics.Clear(Color.White);

                        // Copiar la imagen original sobre el nuevo lienzo
                        //graphics.DrawImage(image, 0, isBottom ? 0 : margin, image.Width, image.Height);

                        // Calcular la posición de la imagen original en la nueva imagen
                        int x = 0;
                        int y = margin;

                        if (isTop)
                        {
                            y += margin; // Ajustar la posición si se agrega margen en la parte superior
                        }

                        // Dibujar la imagen original sobre la nueva imagen
                        graphics.DrawImage(image, x, y);

                        string newImagePath = Path.Combine(Path.GetDirectoryName(imagePath), "modified_" + Path.GetFileName(imagePath));

                        newImage.Save(newImagePath, System.Drawing.Imaging.ImageFormat.Png);
                        return newImagePath;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }*/

        #endregion
    }
}
