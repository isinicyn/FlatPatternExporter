using System.Collections;
using System.Windows;
using System.Windows.Controls;
using FlatPatternExporter.Models;
using ClosedXML.Excel;

namespace FlatPatternExporter.Services;

public class ExcelExportService
{
    private const int ImageSizePixels = 80;

    private static double PixelsToPoints(int pixels) => pixels * 72.0 / 96.0;

    public void ExportToExcel(string filePath, DataGrid dataGrid, IEnumerable view)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Export");

        var visibleColumns = dataGrid.Columns.Where(c => c.Visibility == Visibility.Visible).ToList();

        for (int col = 0; col < visibleColumns.Count; col++)
        {
            worksheet.Cell(1, col + 1).Value = visibleColumns[col].Header.ToString();
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
                                var cell = worksheet.Cell(row, col + 1);

                                var picture = worksheet.AddPicture(imageStream)
                                    .MoveTo(cell, 0, 0, cell, ImageSizePixels, ImageSizePixels);

                                hasImage = true;
                            }
                        }
                        catch
                        {
                            worksheet.Cell(row, col + 1).Value = "[IMAGE]";
                        }
                    }
                    else
                    {
                        worksheet.Cell(row, col + 1).Value = value?.ToString() ?? string.Empty;
                    }
                }
            }

            if (hasImage)
            {
                worksheet.Row(row).Height = PixelsToPoints(ImageSizePixels);
            }

            row++;
        }

        worksheet.Columns().AdjustToContents();

        workbook.SaveAs(filePath);
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
