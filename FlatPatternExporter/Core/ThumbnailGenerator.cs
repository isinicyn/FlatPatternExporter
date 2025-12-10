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
using Microsoft.WindowsAPICodePack.Shell;
using Svg.Skia;

namespace FlatPatternExporter.Core;

public class ThumbnailGenerator
{
    private readonly ApprenticeServerComponent? _apprentice = CreateApprenticeServer();

    private static ApprenticeServerComponent? CreateApprenticeServer()
    {
        try
        {
            return new ApprenticeServerComponent();
        }
        catch
        {
            return null;
        }
    }
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

    public async Task<BitmapImage?> GetThumbnailAsync(PartDocument document, Dispatcher dispatcher)
    {
        if (_apprentice is not null)
        {
            try
            {
                return await GetThumbnailViaApprenticeAsync(document, dispatcher);
            }
            catch
            {
                // Fallback to Shell API
            }
        }

        try
        {
            return await GetThumbnailViaShellAsync(document, dispatcher);
        }
        catch (Exception ex)
        {
            dispatcher.Invoke(() =>
            {
                CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_ThumbnailObtaining", ex.Message), LocalizationManager.Instance.GetString("Error_Title"), MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null;
        }
    }

    private Task<BitmapImage?> GetThumbnailViaApprenticeAsync(PartDocument document, Dispatcher dispatcher)
    {
        var tcs = new TaskCompletionSource<BitmapImage?>();

        dispatcher.InvokeAsync(() =>
        {
            ApprenticeServerDocument? apprenticeDoc = null;
            try
            {
                apprenticeDoc = _apprentice!.Open(document.FullDocumentName);
                var thumbnail = apprenticeDoc.Thumbnail;
                var img = IPictureDispConverter.PictureDispToImage(thumbnail);

                if (img != null)
                {
                    using var memoryStream = new MemoryStream();
                    img.Save(memoryStream, ImageFormat.Png);
                    memoryStream.Position = 0;

                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.StreamSource = memoryStream;
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    bitmap.Freeze();

                    tcs.SetResult(bitmap);
                }
                else
                {
                    tcs.SetResult(null);
                }
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
            finally
            {
                apprenticeDoc?.Close();
            }
        });

        return tcs.Task;
    }

    private async Task<BitmapImage?> GetThumbnailViaShellAsync(PartDocument document, Dispatcher dispatcher)
    {
        BitmapImage? bitmap = null;
        await dispatcher.InvokeAsync(() =>
        {
            var filePath = document.FullFileName;

            if (!System.IO.File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");

            using var shellFile = ShellFile.FromFilePath(filePath);
            using var thumbnail = shellFile.Thumbnail.ExtraLargeBitmap;

            if (thumbnail != null)
                using (var memoryStream = new MemoryStream())
                {
                    thumbnail.Save(memoryStream, ImageFormat.Png);
                    memoryStream.Position = 0;

                    bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.StreamSource = memoryStream;
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    bitmap.Freeze();
                }
        });

        return bitmap;
    }
}