using System.IO;
using netDxf;
using netDxf.Entities;
using netDxf.Header;

namespace DxfGenerator;

public static class DxfOptimizer
{
    public static void OptimizeDxfFile(string dxfFilePath)
    {
        try
        {
            var dxf = DxfDocument.Load(dxfFilePath);
            SaveAsR15(dxf, dxfFilePath);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Ошибка оптимизации DXF: {ex.Message}");
        }
    }

    private static void SaveAsR15(DxfDocument dxf, string originalFilePath)
    {
        try
        {
            var directory = Path.GetDirectoryName(originalFilePath) ?? "";
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalFilePath);
            var extension = Path.GetExtension(originalFilePath);
            var r15FilePath = Path.Combine(directory, $"{fileNameWithoutExtension}_R15{extension}");

            if (File.Exists(r15FilePath))
                return;

            var r15Dxf = new DxfDocument();
            r15Dxf.DrawingVariables.AcadVer = DxfVersion.AutoCad2000;
            
            var allEntities = CollectEntities(dxf);
            foreach (var entity in allEntities)
                r15Dxf.Entities.Add((EntityObject)entity.Clone());

            r15Dxf.Save(r15FilePath);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Ошибка сохранения R15: {ex.Message}");
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