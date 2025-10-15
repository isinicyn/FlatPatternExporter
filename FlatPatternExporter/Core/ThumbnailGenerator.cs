using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using DxfRenderer;
using FlatPatternExporter.Converters;
using FlatPatternExporter.Services;
using FlatPatternExporter.UI.Windows;
using Inventor;
using Svg.Skia;

namespace FlatPatternExporter.Core;

public class ThumbnailGenerator
{
    private readonly ApprenticeServerComponent _apprentice = new();
    /// <summary>
    /// Converts SVG string to BitmapImage
    /// </summary>
    public static BitmapImage? ConvertSvgToBitmapImage(string svgContent)
    {
        if (string.IsNullOrEmpty(svgContent))
            return null;

        try
        {
            var svg = new SKSvg();
            var svgDocument = svg.FromSvg(svgContent);

            if (svgDocument == null)
                return null;

            var bounds = svgDocument.CullRect;

            if (bounds.Width <= 0 || bounds.Height <= 0)
                return null;

            using var surface = SkiaSharp.SKSurface.Create(new SkiaSharp.SKImageInfo((int)bounds.Width, (int)bounds.Height));
            if (surface?.Canvas == null)
                return null;

            surface.Canvas.Clear(SkiaSharp.SKColors.White);
            surface.Canvas.DrawPicture(svgDocument);
            surface.Canvas.Flush();

            using var image = surface.Snapshot();
            using var data = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);

            var bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = new MemoryStream(data.ToArray());
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
            bitmapImage.EndInit();
            bitmapImage.Freeze();

            return bitmapImage;
        }
        catch (Exception)
        {
            return null;
        }
    }

    public BitmapImage? GenerateDxfThumbnails(string dxfFilePath, Dispatcher dispatcher)
    {
        if (!System.IO.File.Exists(dxfFilePath)) return null;

        try
        {
            var generator = new DxfThumbnailGenerator();
            var svg = generator.GenerateSvg(dxfFilePath);

            BitmapImage? bitmapImage = null;

            // Convert SVG to BitmapImage
            dispatcher.Invoke(() =>
            {
                bitmapImage = ConvertSvgToBitmapImage(svg);
            });

            return bitmapImage;
        }
        catch (Exception ex)
        {
            dispatcher.Invoke(() =>
            {
                CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_ThumbnailGeneration", ex.Message), LocalizationManager.Instance.GetString("Error_Title"), MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null;
        }
    }

    public async Task<BitmapImage> GetThumbnailAsync(PartDocument document, Dispatcher dispatcher)
    {
        try
        {
            BitmapImage? bitmap = null;
            await dispatcher.InvokeAsync(() =>
            {
                var apprenticeDoc = _apprentice.Open(document.FullDocumentName);

                var thumbnail = apprenticeDoc.Thumbnail;
                var img = IPictureDispConverter.PictureDispToImage(thumbnail);

                if (img != null)
                    using (var memoryStream = new MemoryStream())
                    {
                        img.Save(memoryStream, ImageFormat.Png);
                        memoryStream.Position = 0;

                        bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = memoryStream;
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        bitmap.Freeze();
                    }
            });

            return bitmap!;
        }
        catch (Exception ex)
        {
            dispatcher.Invoke(() =>
            {
                CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_ThumbnailObtaining", ex.Message), LocalizationManager.Instance.GetString("Error_Title"), MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null!;
        }
    }
}