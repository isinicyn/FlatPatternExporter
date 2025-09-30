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
            System.Diagnostics.Debug.WriteLine($"DXF optimization error: {ex.Message}");
        }
    }

    private static void ReplaceWithOptimizedVersion(DxfDocument dxf, string originalFilePath, AcadVersionType version)
    {
        try
        {
            var dxfVersion = AcadVersionMapping.GetDxfVersion(version);
            if (!dxfVersion.HasValue)
            {
                System.Diagnostics.Debug.WriteLine($"AutoCAD version {AcadVersionMapping.GetDisplayName(version)} is not supported for optimization");
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
            System.Diagnostics.Debug.WriteLine($"Optimization error for {AcadVersionMapping.GetDisplayName(version)}: {ex.Message}");
        }
    }
}