using System.IO;
using netDxf;
using netDxf.Entities;
using netDxf.Header;

namespace DxfGenerator;

public static class DxfOptimizer
{
    private static readonly Dictionary<string, DxfVersion> AcadVersionToDxfVersion = new()
    {
        { "2000", DxfVersion.AutoCad2000 },
        { "2004", DxfVersion.AutoCad2004 },
        { "2007", DxfVersion.AutoCad2007 },
        { "2010", DxfVersion.AutoCad2010 },
        { "2013", DxfVersion.AutoCad2013 },
        { "2018", DxfVersion.AutoCad2018 }
    };

    public static void OptimizeDxfFile(string dxfFilePath, string acadVersion = "2000")
    {
        try
        {
            var dxf = DxfDocument.Load(dxfFilePath);
            SaveAsVersion(dxf, dxfFilePath, acadVersion);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Ошибка оптимизации DXF: {ex.Message}");
        }
    }

    private static void SaveAsVersion(DxfDocument dxf, string originalFilePath, string acadVersion)
    {
        try
        {
            if (!AcadVersionToDxfVersion.TryGetValue(acadVersion, out var dxfVersion))
            {
                dxfVersion = DxfVersion.AutoCad2000;
                System.Diagnostics.Debug.WriteLine($"Неизвестная версия AutoCAD: {acadVersion}, используется 2000 по умолчанию");
                acadVersion = "2000";
            }
            
            var directory = Path.GetDirectoryName(originalFilePath) ?? "";
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalFilePath);
            var extension = Path.GetExtension(originalFilePath);
            var versionFilePath = Path.Combine(directory, $"{fileNameWithoutExtension}_{acadVersion}{extension}");

            if (File.Exists(versionFilePath))
                return;

            var versionDxf = new DxfDocument();
            versionDxf.DrawingVariables.AcadVer = dxfVersion;
            
            var allEntities = CollectEntities(dxf);
            foreach (var entity in allEntities)
                versionDxf.Entities.Add((EntityObject)entity.Clone());

            versionDxf.Save(versionFilePath);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Ошибка сохранения {acadVersion}: {ex.Message}");
        }
    }

    private static List<EntityObject> CollectEntities(DxfDocument dxf)
    {
        var entities = new List<EntityObject>();
        entities.AddRange(dxf.Entities.Lines);
        entities.AddRange(dxf.Entities.Circles);
        entities.AddRange(dxf.Entities.Arcs);
        entities.AddRange(dxf.Entities.Polylines2D);
        entities.AddRange(dxf.Entities.Polylines3D);
        entities.AddRange(dxf.Entities.Splines);
        entities.AddRange(dxf.Entities.Ellipses);
        return entities;
    }
}