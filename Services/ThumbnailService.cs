using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace TocBuilder_dotnet_framework.Services
{
    public class ThumbnailService : IDisposable
    {
        private bool _disposed;

        public List<Models.SlideItem> GetSlides(string filePath)
        {
            var slides = new List<Models.SlideItem>();
            string tempDir = Path.Combine(Path.GetTempPath(), $"ppt_thumbs_{Guid.NewGuid()}");
            Directory.CreateDirectory(tempDir);

            Application pptApp = null;
            Presentation pres = null;

            try
            {
                pptApp = new Application();
                pres = pptApp.Presentations.Open(filePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

                float slideWidth = pres.PageSetup.SlideWidth;
                float slideHeight = pres.PageSetup.SlideHeight;
                float aspect = slideWidth / slideHeight;

                int previewWidth = 320;
                int previewHeight = (int)(previewWidth / aspect);

                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    string thumbPath = Path.Combine(tempDir, $"slide_{i}.png");
                    try
                    {
                        pres.Slides[i].Export(thumbPath, "PNG", previewWidth * 2, previewHeight * 2);
                    }
                    catch
                    {
                        CreateFallbackThumbnail(thumbPath, i, previewWidth, previewHeight);
                    }

                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.UriSource = new Uri(thumbPath);
                    bitmap.EndInit();
                    bitmap.Freeze();

                    slides.Add(new Models.SlideItem
                    {
                        Number = i,
                        Thumbnail = bitmap,
                        IsSelected = true
                    });
                }
            }
            finally
            {
                if (pres != null) { pres.Close(); Marshal.ReleaseComObject(pres); }
                if (pptApp != null) { pptApp.Quit(); Marshal.ReleaseComObject(pptApp); }

                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }

            return slides;
        }

        private void CreateFallbackThumbnail(string filePath, int slideNumber, int width, int height)
        {
            using (Bitmap bmp = new Bitmap(width, height))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.LightGray);
                g.DrawString($"Slide {slideNumber}",
                    new System.Drawing.Font("Arial", 20),
                    Brushes.Black,
                    new PointF(10, 10));
                bmp.Save(filePath, ImageFormat.Png);
            }
        }

        public (float Width, float Height) GetSlideDimensions(string pptxPath)
        {
            try
            {
                using (var doc = PresentationDocument.Open(pptxPath, isEditable: false))
                {
                    var presentationPart = doc.PresentationPart;
                    var presentation = presentationPart?.Presentation;

                    var slideSize = presentation?.SlideSize;
                    if (slideSize != null)
                    {
                        // (в EMU)
                        long? cx = slideSize.Cx;
                        long? cy = slideSize.Cy;

                        if (cx.HasValue && cy.HasValue)
                        {
                            const double EMU_PER_POINT = 12700.0;
                            float widthPts = (float)(cx.Value / EMU_PER_POINT);
                            float heightPts = (float)(cy.Value / EMU_PER_POINT);
                            return (widthPts, heightPts);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Для отладки можно логировать: Debug.WriteLine($"GetSlideDimensions error: {ex}");
            }

            // Fallback на 16:9 (960×540 pt)
            return (LayoutConstants.DefaultSlideWidth, LayoutConstants.DefaultSlideHeight);
        }

        public void Dispose()
        {
            if (!_disposed) { _disposed = true; GC.SuppressFinalize(this); }
        }
    }
}
