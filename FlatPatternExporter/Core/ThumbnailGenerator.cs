using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using DxfRenderer;
using FlatPatternExporter.Services;
using FlatPatternExporter.UI.Windows;
using Svg.Skia;

namespace FlatPatternExporter.Core;

public static class ThumbnailGenerator
{
    // Inventor Summary Information property set (replaces Microsoft FMTID_SummaryInformation since Inventor 6.0)
    private static readonly Guid InventorSummaryInformationId = new("3D38DE39-0588-4C14-BB37-18F4D5DD31C7");

    private const int PidThumbnail = 0x11;
    private const ushort VtClipboardData = 71;
    private const int ThumbnailDecodeWidth = 256;

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

            return CreateFrozenBitmap(data.ToArray());
        }
        catch (Exception)
        {
            return null;
        }
    }

    public static BitmapImage? GenerateDxfThumbnails(string dxfFilePath, Dispatcher dispatcher)
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

    /// <summary>
    /// Gets the thumbnail saved in an Inventor file by reading the embedded PNG via OLE Structured Storage.
    /// Returns null for files without one (legacy WMF/EMF thumbnails). Thread-safe, the result is frozen.
    /// </summary>
    public static BitmapImage? GetThumbnail(string fullFileName)
    {
        return ReadEmbeddedPng(fullFileName) is { } png ? CreateFrozenBitmap(png, ThumbnailDecodeWidth) : null;
    }

    private static BitmapImage CreateFrozenBitmap(byte[] imageData, int decodePixelWidth = 0)
    {
        using var stream = new MemoryStream(imageData);
        var bitmap = new BitmapImage();
        bitmap.BeginInit();
        bitmap.StreamSource = stream;
        bitmap.CacheOption = BitmapCacheOption.OnLoad;
        bitmap.DecodePixelWidth = decodePixelWidth;
        bitmap.EndInit();
        bitmap.Freeze();
        return bitmap;
    }

    private static byte[]? ReadEmbeddedPng(string filePath)
    {
        NativeMethods.IPropertySetStorage? storage = null;
        NativeMethods.IPropertyStorage? propertyStorage = null;
        var values = new NativeMethods.PropVariant[1];

        try
        {
            NativeMethods.StgOpenStorage(filePath, IntPtr.Zero,
                NativeMethods.STGM_READ | NativeMethods.STGM_SHARE_DENY_WRITE, IntPtr.Zero, 0, out storage);

            var fmtid = InventorSummaryInformationId;
            storage.Open(ref fmtid, NativeMethods.STGM_READ | NativeMethods.STGM_SHARE_EXCLUSIVE, out propertyStorage);
            propertyStorage.ReadMultiple(1, [new() { Kind = NativeMethods.PRSPEC_PROPID, PropId = PidThumbnail }], values);

            if (values[0].Vt != VtClipboardData || values[0].Ptr == IntPtr.Zero)
                return null;

            // CLIPDATA: cbSize includes ulClipFmt; Inventor data is a 12-byte header followed by the image
            const int ClipFormatSize = 4;
            const int HeaderSize = 12;
            var clipData = Marshal.PtrToStructure<NativeMethods.ClipData>(values[0].Ptr);
            var imageSize = (int)clipData.Size - ClipFormatSize - HeaderSize;
            if (imageSize <= 8 || clipData.Data == IntPtr.Zero)
                return null;

            var image = new byte[imageSize];
            Marshal.Copy(clipData.Data + HeaderSize, image, 0, imageSize);

            // Older files store WMF/EMF instead of PNG
            return image.AsSpan().StartsWith(NativeMethods.PngSignature) ? image : null;
        }
        catch
        {
            return null;
        }
        finally
        {
            NativeMethods.PropVariantClear(ref values[0]);
            if (propertyStorage != null) Marshal.ReleaseComObject(propertyStorage);
            if (storage != null) Marshal.ReleaseComObject(storage);
        }
    }

    private static class NativeMethods
    {
        public const uint STGM_READ = 0x00000000;
        public const uint STGM_SHARE_EXCLUSIVE = 0x00000010;
        public const uint STGM_SHARE_DENY_WRITE = 0x00000020;
        public const uint PRSPEC_PROPID = 1;

        public static ReadOnlySpan<byte> PngSignature => [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern void StgOpenStorage(string name, IntPtr priority, uint mode, IntPtr exclude, uint reserved,
            [MarshalAs(UnmanagedType.Interface)] out IPropertySetStorage storage);

        [DllImport("ole32.dll")]
        public static extern int PropVariantClear(ref PropVariant value);

        [ComImport, Guid("0000013A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertySetStorage
        {
            void Create(ref Guid fmtid, IntPtr clsid, uint flags, uint mode, out IntPtr propertyStorage);

            void Open(ref Guid fmtid, uint mode, out IPropertyStorage propertyStorage);
        }

        // ReadMultiple is the first vtable slot, the remaining members are not used
        [ComImport, Guid("00000138-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyStorage
        {
            void ReadMultiple(uint count,
                [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PropSpec[] specs,
                [In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PropVariant[] values);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PropSpec
        {
            public uint Kind;
            public IntPtr PropId;
        }

        // PROPVARIANT trimmed to the type tag and the data pointer (x64 layout)
        [StructLayout(LayoutKind.Explicit, Size = 24)]
        public struct PropVariant
        {
            [FieldOffset(0)] public ushort Vt;
            [FieldOffset(8)] public IntPtr Ptr;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ClipData
        {
            public uint Size;
            public int Format;
            public IntPtr Data;
        }
    }
}
