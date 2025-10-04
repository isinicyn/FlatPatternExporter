using System.Collections;
using System.Windows;
using System.Windows.Controls;
using FlatPatternExporter.Models;
using OfficeOpenXml;

namespace FlatPatternExporter.Services;

public class ExcelExportService
{
    private const int ImageSizePixels = 80;
    private const int PixelToEmu = 9525;

    private static double PixelsToPoints(int pixels) => pixels * 72.0 / 96.0;

    public void ExportToExcel(string filePath, DataGrid dataGrid, IEnumerable view)
    {
        var fileInfo = new System.IO.FileInfo(filePath);
        using var package = new ExcelPackage(fileInfo);

        var worksheet = package.Workbook.Worksheets.Add("Export");

        var visibleColumns = dataGrid.Columns.Where(c => c.Visibility == Visibility.Visible).ToList();

        for (int col = 0; col < visibleColumns.Count; col++)
        {
            worksheet.Cells[1, col + 1].Value = visibleColumns[col].Header.ToString();
        }

        int row = 2;
        foreach (PartData item in view)
        {
            bool hasImage = false;
            for (int col = 0; col < visibleColumns.Count; col++)
            {
                var column = visibleColumns[col];
                if (!string.IsNullOrEmpty(column.SortMemberPath))
                {
                    var value = item.GetType().GetProperty(column.SortMemberPath)?.GetValue(item, null);

                    if (value is System.Windows.Media.Imaging.BitmapImage bitmapImage)
                    {
                        try
                        {
                            var imageBytes = BitmapImageToBytes(bitmapImage);
                            if (imageBytes != null && imageBytes.Length > 0)
                            {
                                var imageStream = new System.IO.MemoryStream(imageBytes);
                                var picture = worksheet.Drawings.AddPicture($"Image_R{row}C{col + 1}", imageStream);
                                picture.ChangeCellAnchor(OfficeOpenXml.Drawing.eEditAs.TwoCell);

                                picture.From.Row = row - 1;
                                picture.From.Column = col;
                                picture.From.RowOff = 0;
                                picture.From.ColumnOff = 0;

                                picture.To.Row = row - 1;
                                picture.To.Column = col;
                                picture.To.RowOff = ImageSizePixels * PixelToEmu;
                                picture.To.ColumnOff = ImageSizePixels * PixelToEmu;

                                hasImage = true;
                            }
                        }
                        catch
                        {
                            worksheet.Cells[row, col + 1].Value = "[IMAGE]";
                        }
                    }
                    else
                    {
                        worksheet.Cells[row, col + 1].Value = value;
                    }
                }
            }

            if (hasImage)
            {
                worksheet.Row(row).Height = PixelsToPoints(ImageSizePixels);
            }

            row++;
        }

        if (worksheet.Dimension != null)
        {
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }

        package.Save();
    }

    public void ExportToCsv(string filePath, DataGrid dataGrid, IEnumerable view)
    {
        using var writer = new System.IO.StreamWriter(filePath, false, System.Text.Encoding.UTF8);

        var visibleColumns = dataGrid.Columns.Where(c => c.Visibility == Visibility.Visible).ToList();

        var headers = visibleColumns.Select(c => c.Header.ToString());
        writer.WriteLine(string.Join("\t", headers));

        foreach (PartData item in view)
        {
            var values = visibleColumns.Select(column =>
            {
                if (!string.IsNullOrEmpty(column.SortMemberPath))
                {
                    var value = item.GetType().GetProperty(column.SortMemberPath)?.GetValue(item, null);
                    if (value is System.Windows.Media.Imaging.BitmapImage)
                    {
                        return "[IMAGE]";
                    }
                    return value?.ToString() ?? string.Empty;
                }
                return string.Empty;
            });
            writer.WriteLine(string.Join("\t", values));
        }
    }

    private static byte[]? BitmapImageToBytes(System.Windows.Media.Imaging.BitmapImage bitmapImage)
    {
        try
        {
            var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmapImage));
            using var stream = new System.IO.MemoryStream();
            encoder.Save(stream);
            return stream.ToArray();
        }
        catch
        {
            return null;
        }
    }
}
