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
using System.Threading;
using System.Collections.Concurrent;
using EmailSendValidationService;
using PuppeteerSharp.BrowserData;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;
using EmailSendValidationService.Logs;
//using OpenQA.Selenium;
//using OpenQA.Selenium.Chrome;

namespace ServiceEmailSendValidation.Converters
{
    public class HtmlToTiffConverter
    {
        #region "Declaraciones"

        private readonly string _outputFolderPath;
        private readonly int _pageWidth;
        private readonly int _pageHeight;
        private readonly int _defaultMarginTop;
        protected Logs _DataLog;
        private readonly IBrowser _browser;
        private readonly object _browserLock = new object();

        //private static readonly object _lockObject = new object();  // Objeto para sincronizar el acceso a escritura DBTools y asegurar ejecución por un solo hilo a la vez.

        #endregion

        #region "Constructores"

        public HtmlToTiffConverter(string outputFolderPath, int pageWidth, int pageHeight, int defaultMarginTop, Logs dataLog, IBrowser browser)
        {
            _DataLog = dataLog;
            _outputFolderPath = outputFolderPath;
            _pageWidth = pageWidth;
            _pageHeight = pageHeight;
            _defaultMarginTop = defaultMarginTop;
            _browser = browser;
        }

        #endregion

        #region "Metodos"

        public void ConvertHtmlToTiff(string htmlFilePath, string outputTiffPath)
        {
            IPage page = null;

            try
            {
                if (_browser != null)
                {
                    lock (_browserLock)
                    {
                        // Abrir una nueva página sincrónicamente
                        page = _browser.NewPageAsync().GetAwaiter().GetResult();
                    }

                    // Establecer el tiempo de espera para la página
                    page.DefaultTimeout = 300000; // Establecer el tiempo de espera a 5 minutos (300,000 ms)

                    // Llamar al método interno que realiza la conversión (también necesitas una versión sincrónica)
                    ConvertHtmlToTiffInternal((Page)page, htmlFilePath, outputTiffPath);
                }
                else
                {
                    _DataLog.AddErrorEntry("[Error] No se tiene valor para la instancia del navegador, por lo tanto no se puede crear la pagina para el Render del HTML");
                }
            }
            catch (Exception ex)
            {
                throw; 
            }
            finally
            {
                // Asegurarse de cerrar la página y el navegador
                if (page != null)
                {
                    try
                    {
                        page.CloseAsync().GetAwaiter().GetResult();
                        page.DisposeAsync().GetAwaiter().GetResult();
                    }
                    catch (Exception ex)
                    {
                        _DataLog.AddErrorEntry("[Error] Error closing page: " + ex.Message);
                    }
                }
            }
        }

        public void ConvertHtmlToTiffInternal(Page page, string htmlFilePath, string outputTiffPath)
        {
            try
            {
                var htmlContent = File.ReadAllText(htmlFilePath);            // lee el html
                page.SetContentAsync(htmlContent).GetAwaiter().GetResult();

                System.Threading.Thread.Sleep(2000);

                // Establecer el tamaño de la vista de la página
                page.SetViewportAsync(new ViewPortOptions
                {
                    Width = _pageWidth,
                    Height = _pageHeight
                }).GetAwaiter().GetResult();

                // Esperar a que el documento esté completamente cargado
                page.WaitForFunctionAsync("() => document.readyState === 'complete'").GetAwaiter().GetResult();

                // Verificar que el contenedor principal está presente y visible
                page.WaitForFunctionAsync("() => { const container = document.querySelector('.container'); return container && container.offsetHeight > 0; }").GetAwaiter().GetResult();

                // Asegurarse de que hay contenido en el encabezado y al menos una sección
                page.WaitForFunctionAsync("() => { const headerText = document.querySelector('.header').innerText; const sections = document.querySelectorAll('.section'); return headerText.trim().length > 0 && sections.length > 0; }").GetAwaiter().GetResult();

                // Esperar hasta que todas las imágenes estén completamente cargadas
                page.WaitForFunctionAsync("() => document.readyState === 'complete' && Array.from(document.images).every(img => img.complete)").GetAwaiter().GetResult();

                // Esperar otros 2 segundos de manera sincrónica
                System.Threading.Thread.Sleep(2000);

                ///// DEBUG
                ////// Capturar la imagen de la página completa
                //var screenshotPath = Path.Combine(_outputFolderPath, $"{Path.GetFileNameWithoutExtension(outputTiffPath)}_test.png");
                //await page.ScreenshotAsync(screenshotPath, new ScreenshotOptions
                //{
                //    FullPage = true  // Captura toda la página
                //});

                var tempImagePaths = CapturePageScreenshots((Page)page);   // Capturar las imágenes por partes
                CreateTiffFromImages(tempImagePaths, outputTiffPath);                 // Convertir las imágenes en un archivo TIFF
                CleanupTemporaryImages(tempImagePaths);                               // Eliminar las imágenes temporales
            }
            catch
            {
                throw;
            }
        }

        //// Método que gestiona la creación de navegadores y la conversión de HTML a TIFF
        //public async Task ConvertHtmlToTiffAsync(string htmlFilePath, string outputTiffPath)
        //{
        //    var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
        //    var page = (Page)await browser.NewPageAsync();                                               // Crear una página para este navegador

        //    await ConvertHtmlToTiffInternalAsync(page, htmlFilePath, outputTiffPath);                // Cargar el contenido HTML y realizar la conversión a TIFF

        //    await browser.CloseAsync();
        //    await page.CloseAsync();            
        //}

        //public async Task ConvertHtmlToTiffInternalAsync(Page page, string htmlFilePath, string outputTiffPath)
        //{
        //    try
        //    {
        //        var htmlContent = File.ReadAllText(htmlFilePath);            // lee el html
        //        await page.SetContentAsync(htmlContent);

        //        await Task.Delay(2000);

        //        // Establecer el tamaño de la vista de la página
        //        await page.SetViewportAsync(new ViewPortOptions
        //        {
        //            Width = _pageWidth,
        //            Height = _pageHeight
        //        });

        //        // Esperar a que el documento esté completamente cargado
        //        await page.WaitForFunctionAsync("() => document.readyState === 'complete'");

        //        // Verificar que el contenedor principal está presente y visible
        //        //await page.WaitForSelectorAsync(".container");
        //        await page.WaitForFunctionAsync("() => { const container = document.querySelector('.container'); return container && container.offsetHeight > 0; }");

        //        // Asegúrate de que hay contenido en el encabezado y al menos una sección
        //        await page.WaitForFunctionAsync("() => { const headerText = document.querySelector('.header').innerText; const sections = document.querySelectorAll('.section'); return headerText.trim().length > 0 && sections.length > 0; }");

        //        await page.WaitForFunctionAsync("() => document.readyState === 'complete' && Array.from(document.images).every(img => img.complete)");

        //        await Task.Delay(2000);

        //        ///// DEBUG
        //        ////// Capturar la imagen de la página completa
        //        //var screenshotPath = Path.Combine(_outputFolderPath, $"{Path.GetFileNameWithoutExtension(outputTiffPath)}_test.png");
        //        //await page.ScreenshotAsync(screenshotPath, new ScreenshotOptions
        //        //{
        //        //    FullPage = true  // Captura toda la página
        //        //});

        //        var tempImagePaths = await CapturePageScreenshotsAsync((Page)page);   // Capturar las imágenes por partes
        //        CreateTiffFromImages(tempImagePaths, outputTiffPath);                 // Convertir las imágenes en un archivo TIFF
        //        CleanupTemporaryImages(tempImagePaths);                               // Eliminar las imágenes temporales
        //    }
        //    catch
        //    {
        //        throw;
        //    }
        //}



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

        private List<string> CapturePageScreenshots(Page page)
        {
            var tempImagePaths = new List<string>();
            var tempImageCorrectPaths = new List<string>();
            int currentPage = 1;
            bool hasMoreContent = true;
            Guid identifier = Guid.NewGuid();

            // Obtener la posición superior e inferior de todas las líneas de texto visibles en la página
            var linePositions = page.EvaluateFunctionAsync<List<Dictionary<string, int>>>(@"
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
             ").GetAwaiter().GetResult();  // Convertir a llamada sincrónica

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

                var bodyHeight = page.EvaluateExpressionAsync<double>(@"document.body.scrollHeight").GetAwaiter().GetResult(); // Obtener el valor sincrónicamente
                bool isBodyMax = bodyHeight <= (_pageHeight * currentPage);                                                    // Evaluar si hay más contenido para capturar  

                // Calcular la altura y posición de la página a capturar
                int pageHeight = (currentPage == 1) ? lastVisibleLineY : lastVisibleLineY - previousPageHeight;
                pageHeight = (!isBodyMax) ? pageHeight : (bottomValue - previousPageHeight);
                int pageY = (currentPage == 1) ? 0 : previousPageHeight;

                // Actualizar la altura de la página previa
                previousPageHeight = (currentPage == 1) ? pageHeight : (previousPageHeight + pageHeight);

                // Capturar la imagen de la vista actual
                page.ScreenshotAsync(screenshotPath, new ScreenshotOptions
                {
                    Clip = new Clip
                    {
                        X = 0,
                        Y = pageY,
                        Width = _pageWidth,
                        Height = pageHeight
                    },
                    FullPage = false
                }).GetAwaiter().GetResult();  // Convertir a llamada sincrónica

                tempImagePaths.Add(screenshotPath);

                // Evaluar si hay más contenido para capturar
                if (bodyHeight <= (_pageHeight * currentPage))
                {
                    hasMoreContent = false;
                }
                else
                {
                    page.EvaluateExpressionAsync($"window.scrollTo(0, {lastVisibleLineY});").GetAwaiter().GetResult();  // Convertir a llamada sincrónica
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
        #endregion
    }
}
