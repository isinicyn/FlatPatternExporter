using netDxf;
using netDxf.Entities;
using FlatPatternExporter.Enums;

namespace FlatPatternExporter.Utilities;

public static class DxfOptimizer
{
    public static void OptimizeDxfFile(string dxfFilePath, AcadVersionType version)
    {
        try
        {
            var dxf = DxfDocument.Load(dxfFilePath);
            ReplaceWithOptimizedVersion(dxf, dxfFilePath, version);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Ошибка оптимизации DXF: {ex.Message}");
        }
    }

    private static void ReplaceWithOptimizedVersion(DxfDocument dxf, string originalFilePath, AcadVersionType version)
    {
        try
        {
            var dxfVersion = AcadVersionMapping.GetDxfVersion(version);
            if (!dxfVersion.HasValue)
            {
                System.Diagnostics.Debug.WriteLine($"Версия AutoCAD {AcadVersionMapping.GetDisplayName(version)} не поддерживается для оптимизации");
                return;
            }

            var optimizedDxf = new DxfDocument();
            optimizedDxf.DrawingVariables.AcadVer = dxfVersion.Value;

            foreach (var entity in dxf.Entities.All)
                optimizedDxf.Entities.Add((EntityObject)entity.Clone());

            optimizedDxf.Save(originalFilePath);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Ошибка оптимизации {AcadVersionMapping.GetDisplayName(version)}: {ex.Message}");
        }
    }
}