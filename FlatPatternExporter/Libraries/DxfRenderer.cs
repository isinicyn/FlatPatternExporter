using System.Drawing.Drawing2D;
using System.Globalization;
using System.Text;
using netDxf;
using netDxf.Entities;

namespace DxfRenderer;

public abstract class BaseRenderer
{
    public const int DefaultThumbnailSize = 100;
    protected const double BoundsPadding = 0.9;
    protected const int SplineSubdivisions = 50;
    protected const double MinGeometrySize = 0.05;
    protected const double MinBoundsSize = 1.0;
    protected static readonly double[] CardinalAngles = { 0, 90, 180, 270 };
    protected static readonly Color DefaultEntityColor = Color.Black;
    
    public abstract object Render(DxfDocument dxf, int width, int height);
    public abstract object Render(List<EntityObject> entities, (double MinX, double MinY, double MaxX, double MaxY) bounds, int width, int height);
    
    public static List<EntityObject> CollectEntities(DxfDocument dxf)
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
    
    public (double MinX, double MinY, double MaxX, double MaxY) CalculateBoundingBox(IEnumerable<EntityObject> entities)
    {
        double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
        foreach (var entity in entities)
            UpdateBounds(entity, ref minX, ref minY, ref maxX, ref maxY);
        return (minX, minY, maxX, maxY);
    }
    
    protected void UpdateBounds(EntityObject entity, ref double minX, ref double minY, ref double maxX, ref double maxY)
    {
        if (entity is Line line)
        {
            minX = Math.Min(minX, Math.Min(line.StartPoint.X, line.EndPoint.X));
            minY = Math.Min(minY, Math.Min(line.StartPoint.Y, line.EndPoint.Y));
            maxX = Math.Max(maxX, Math.Max(line.StartPoint.X, line.EndPoint.X));
            maxY = Math.Max(maxY, Math.Max(line.StartPoint.Y, line.EndPoint.Y));
        }
        else if (entity is Circle circle)
        {
            minX = Math.Min(minX, circle.Center.X - circle.Radius);
            minY = Math.Min(minY, circle.Center.Y - circle.Radius);
            maxX = Math.Max(maxX, circle.Center.X + circle.Radius);
            maxY = Math.Max(maxY, circle.Center.Y + circle.Radius);
        }
        else if (entity is Arc arc)
        {
            var startX = arc.Center.X + arc.Radius * Math.Cos(arc.StartAngle * Math.PI / 180);
            var startY = arc.Center.Y + arc.Radius * Math.Sin(arc.StartAngle * Math.PI / 180);
            var endX = arc.Center.X + arc.Radius * Math.Cos(arc.EndAngle * Math.PI / 180);
            var endY = arc.Center.Y + arc.Radius * Math.Sin(arc.EndAngle * Math.PI / 180);

            minX = Math.Min(minX, Math.Min(startX, endX));
            minY = Math.Min(minY, Math.Min(startY, endY));
            maxX = Math.Max(maxX, Math.Max(startX, endX));
            maxY = Math.Max(maxY, Math.Max(startY, endY));

            foreach (var angle in CardinalAngles)
                if (IsAngleBetween(angle, arc.StartAngle, arc.EndAngle))
                {
                    var radian = angle * Math.PI / 180;
                    var x = arc.Center.X + arc.Radius * Math.Cos(radian);
                    var y = arc.Center.Y + arc.Radius * Math.Sin(radian);

                    minX = Math.Min(minX, x);
                    minY = Math.Min(minY, y);
                    maxX = Math.Max(maxX, x);
                    maxY = Math.Max(maxY, y);
                }
        }
        else if (entity is Polyline2D polyline2D)
        {
            foreach (var vertex in polyline2D.Vertexes)
                UpdateBoundsForPoint(new Vector3(vertex.Position.X, vertex.Position.Y, 0), ref minX, ref minY, ref maxX, ref maxY);

            var numSegments = polyline2D.IsClosed ? polyline2D.Vertexes.Count : polyline2D.Vertexes.Count - 1;
            for (var i = 0; i < numSegments; i++)
            {
                var startVertex = polyline2D.Vertexes[i];
                if (startVertex.Bulge == 0) continue;

                var endVertex = polyline2D.Vertexes[(i + 1) % polyline2D.Vertexes.Count];

                var arcGeom = GetArcGeomFromBulge(startVertex, endVertex);
                if (arcGeom == null) continue;

                var (center, radius, startAngle, endAngle) = arcGeom.Value;

                foreach (var angle in CardinalAngles)
                    if (IsAngleBetween(angle, startAngle, endAngle))
                    {
                        var radian = angle * Math.PI / 180;
                        var x = center.X + radius * Math.Cos(radian);
                        var y = center.Y + radius * Math.Sin(radian);

                        minX = Math.Min(minX, x);
                        minY = Math.Min(minY, y);
                        maxX = Math.Max(maxX, x);
                        maxY = Math.Max(maxY, y);
                    }
            }
        }
        else if (entity is Polyline3D polyline3D)
        {
            foreach (var vertex in polyline3D.Vertexes)
            {
                minX = Math.Min(minX, vertex.X);
                minY = Math.Min(minY, vertex.Y);
                maxX = Math.Max(maxX, vertex.X);
                maxY = Math.Max(maxY, vertex.Y);
            }
        }
        else if (entity is Spline spline)
        {
            var splineVertices = InterpolateSpline(spline.ControlPoints.ToList(), spline.Degree, spline.Knots.ToList(), SplineSubdivisions);
            foreach (var point in splineVertices)
            {
                minX = Math.Min(minX, point.X);
                minY = Math.Min(minY, point.Y);
                maxX = Math.Max(maxX, point.X);
                maxY = Math.Max(maxY, point.Y);
            }
        }
        else if (entity is Ellipse ellipse)
        {
            var a = ellipse.MajorAxis;
            var b = ellipse.MinorAxis;

            var majorAxisVector = ellipse.MajorAxis * ellipse.Normal;
            var rotation = Math.Atan2(majorAxisVector.Y, majorAxisVector.X);

            var cosRotation = Math.Cos(rotation);
            var sinRotation = Math.Sin(rotation);

            var points = new[]
            {
                new Vector2(a * cosRotation - b * sinRotation, a * sinRotation + b * cosRotation),
                new Vector2(-a * cosRotation - b * sinRotation, -a * sinRotation + b * cosRotation),
                new Vector2(a * cosRotation + b * sinRotation, a * sinRotation - b * cosRotation),
                new Vector2(-a * cosRotation + b * sinRotation, -a * sinRotation - b * cosRotation)
            };

            foreach (var point in points)
                UpdateBoundsForPoint(new Vector3(point.X + ellipse.Center.X, point.Y + ellipse.Center.Y, 0), ref minX,
                    ref minY, ref maxX, ref maxY);
        }
        else if (entity.GetType().Name == "Polyline")
        {
            var vertexes = GetPolylineVertexes(entity);
            foreach (var vertex in vertexes)
            {
                minX = Math.Min(minX, vertex.X);
                minY = Math.Min(minY, vertex.Y);
                maxX = Math.Max(maxX, vertex.X);
                maxY = Math.Max(maxY, vertex.Y);
            }
        }
    }
    
    protected void UpdateBoundsForPoint(Vector3 point, ref double minX, ref double minY, ref double maxX, ref double maxY)
    {
        minX = Math.Min(minX, point.X);
        minY = Math.Min(minY, point.Y);
        maxX = Math.Max(maxX, point.X);
        maxY = Math.Max(maxY, point.Y);
    }
    
    protected bool IsAngleBetween(double angle, double startAngle, double endAngle)
    {
        angle = (angle + 360) % 360;
        startAngle = (startAngle + 360) % 360;
        endAngle = (endAngle + 360) % 360;

        if (startAngle < endAngle)
            return startAngle <= angle && angle <= endAngle;
        return startAngle <= angle || angle <= endAngle;
    }
    
    protected (Vector2 center, double radius, double startAngle, double endAngle)? GetArcGeomFromBulge(
        Polyline2DVertex startV, Polyline2DVertex endV)
    {
        var bulge = startV.Bulge;
        if (Math.Abs(bulge) < 1e-9) return null;

        var startPoint = startV.Position;
        var endPoint = endV.Position;

        var dx = endPoint.X - startPoint.X;
        var dy = endPoint.Y - startPoint.Y;
        var chordLength = Math.Sqrt(dx * dx + dy * dy);

        if (Math.Abs(chordLength) < 1e-9) return null;

        var angle = 4 * Math.Atan(Math.Abs(bulge));
        var radius = chordLength / (2 * Math.Sin(angle / 2));

        var chordAngle = Math.Atan2(endPoint.Y - startPoint.Y, endPoint.X - startPoint.X);

        var angleToCenter = chordAngle + Math.Sign(bulge) * (Math.PI / 2 - angle / 2);

        var centerX = startPoint.X + Math.Cos(angleToCenter) * radius;
        var centerY = startPoint.Y + Math.Sin(angleToCenter) * radius;

        var center = new Vector2(centerX, centerY);

        var startAngle = Math.Atan2(startPoint.Y - center.Y, startPoint.X - center.X) * 180 / Math.PI;
        var endAngle = Math.Atan2(endPoint.Y - center.Y, endPoint.X - center.X) * 180 / Math.PI;

        if (bulge < 0) (startAngle, endAngle) = (endAngle, startAngle);

        return (center, radius, startAngle, endAngle);
    }
    
    protected List<Vector3> GetPolylineVertexes(EntityObject polyline)
    {
        var vertexes = new List<Vector3>();
        if (polyline.GetType().Name != "Polyline") return vertexes;

        var vertexesProperty = polyline.GetType().GetProperty("Vertexes");
        if (vertexesProperty?.GetValue(polyline) is not IEnumerable<object> vertexesList) return vertexes;

        foreach (var vertex in vertexesList)
        {
            var positionProperty = vertex.GetType().GetProperty("Position");
            if (positionProperty?.GetValue(vertex) is Vector3 position)
                vertexes.Add(position);
        }

        return vertexes;
    }
    
    protected PolylineSmoothType GetPolylineSmoothType(EntityObject entity)
    {
        var smoothTypeProperty = entity.GetType().GetProperty("SmoothType");
        if (smoothTypeProperty?.GetValue(entity) is PolylineSmoothType smoothType)
            return smoothType;
        return PolylineSmoothType.NoSmooth;
    }
    
    protected List<Vector3> InterpolateSpline(List<Vector3> controlPoints, int degree, List<double> knotValues,
        int subdivisions)
    {
        var vertices = new List<Vector3>();
        var step = 1.0 / subdivisions;

        for (var i = 0; i <= subdivisions; i++)
        {
            var t = i * step * (knotValues.Last() - knotValues.First()) + knotValues.First();
            var point = DeBoor(t, degree, controlPoints, knotValues);
            vertices.Add(point);
        }

        return vertices;
    }

    protected Vector3 DeBoor(double t, int degree, List<Vector3> controlPoints, List<double> knotValues)
    {
        var n = controlPoints.Count;

        var s = -1;
        for (var i = degree; i < knotValues.Count - degree - 1; i++)
            if (t >= knotValues[i] && t < knotValues[i + 1])
            {
                s = i;
                break;
            }

        if (s == -1) return controlPoints[controlPoints.Count - 1];

        var dPoints = new Vector3[degree + 1];
        for (var j = 0; j <= degree; j++) dPoints[j] = controlPoints[s - degree + j];

        for (var r = 1; r <= degree; r++)
        for (var j = degree; j >= r; j--)
        {
            var alpha = (t - knotValues[s - degree + j]) / (knotValues[s + 1 - r + j] - knotValues[s - degree + j]);
            dPoints[j] = (1.0 - alpha) * dPoints[j - 1] + alpha * dPoints[j];
        }

        return dPoints[degree];
    }
    
    protected Color GetEntityColor(EntityObject entity)
    {
        Color color;

        color = entity.Color.IsByLayer && entity.Layer != null 
            ? entity.Layer.Color.ToColor() 
            : entity.Color.IsByLayer 
                ? DefaultEntityColor 
                : entity.Color.ToColor();

        if (color.ToArgb() == Color.White.ToArgb()) 
            color = DefaultEntityColor;

        return color;
    }
    
    protected string GetEntityColorHex(EntityObject entity)
    {
        var color = GetEntityColor(entity);
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
    }
}

public class BitmapRenderer : BaseRenderer
{
    public override object Render(DxfDocument dxf, int width, int height)
    {
        var entities = BaseRenderer.CollectEntities(dxf);
        if (entities.Count == 0) return new Bitmap(width, height);

        var bounds = CalculateBoundingBox(entities);
        return Render(entities, bounds, width, height);
    }
    
    public override object Render(List<EntityObject> entities, (double MinX, double MinY, double MaxX, double MaxY) bounds, int width, int height)
    {
        var bitmap = new Bitmap(width, height);

        using var g = Graphics.FromImage(bitmap);
        ConfigureGraphics(g);

        var transform = CalculateTransform(bounds, width, height);
        var penWidth = CalculatePenWidth(bounds, width, height);

        RenderEntities(g, entities, transform, penWidth, height);
        
        return bitmap;
    }
    
    private static void ConfigureGraphics(Graphics g)
    {
        g.Clear(Color.White);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.CompositingQuality = CompositingQuality.HighQuality;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
    }
    
    private (double Scale, double OffsetX, double OffsetY) CalculateTransform(
        (double MinX, double MinY, double MaxX, double MaxY) bounds, int width, int height)
    {
        var rangeX = Math.Max(bounds.MaxX - bounds.MinX, MinBoundsSize);
        var rangeY = Math.Max(bounds.MaxY - bounds.MinY, MinBoundsSize);
        
        var scaleX = width / rangeX;
        var scaleY = height / rangeY;
        var scale = Math.Min(scaleX, scaleY) * BoundsPadding;
        
        if (!double.IsFinite(scale) || scale <= 0)
            scale = 1.0;
        
        var offsetX = (width - rangeX * scale) / 2 - bounds.MinX * scale;
        var offsetY = (height - rangeY * scale) / 2 - bounds.MinY * scale;
        
        if (!double.IsFinite(offsetX)) offsetX = 0;
        if (!double.IsFinite(offsetY)) offsetY = 0;
        
        return (scale, offsetX, offsetY);
    }
    
    private float CalculatePenWidth((double MinX, double MinY, double MaxX, double MaxY) bounds, int width, int height)
    {
        var drawingWidth = bounds.MaxX - bounds.MinX;
        var drawingHeight = bounds.MaxY - bounds.MinY;
        var avgDimension = (drawingWidth + drawingHeight) / 2;
        
        var canvasSize = Math.Min(width, height);
        var scale = canvasSize / avgDimension * BoundsPadding;
        var penWidth = (float)(avgDimension * 0.0225 * scale);
        
        penWidth = Math.Max((float)(avgDimension * 0.015 * scale), 
                           Math.Min((float)(avgDimension * 0.0375 * scale), penWidth));
        
        return Math.Max(1f, penWidth);
    }
    
    private void RenderEntities(Graphics g, IEnumerable<EntityObject> entities, 
        (double Scale, double OffsetX, double OffsetY) transform, float penWidth, int canvasHeight)
    {
        foreach (var entity in entities)
        {
            using var pen = new Pen(GetEntityColor(entity), penWidth);
            RenderEntity(g, entity, transform, pen, canvasHeight);
        }
    }

    private void RenderEntity(Graphics g, EntityObject entity, 
        (double Scale, double OffsetX, double OffsetY) transform, Pen pen, int canvasHeight)
    {
        switch (entity)
        {
            case Line line:
                RenderLine(g, line, transform, pen, canvasHeight);
                break;
            case Circle circle:
                RenderCircle(g, circle, transform, pen, canvasHeight);
                break;
            case Arc arc:
                RenderArc(g, arc, transform, pen, canvasHeight);
                break;
            case Polyline2D polyline2D:
                RenderPolyline2D(g, polyline2D, transform, pen, canvasHeight);
                break;
            case Polyline3D polyline3D:
                RenderPolyline3D(g, polyline3D, transform, pen, canvasHeight);
                break;
            case Spline spline:
                RenderSpline(g, spline, transform, pen, canvasHeight);
                break;
            case Ellipse ellipse:
                RenderEllipse(g, ellipse, transform, pen, canvasHeight);
                break;
            default:
                if (entity.GetType().Name == "Polyline")
                    RenderGenericPolyline(g, entity, transform, pen, canvasHeight);
                break;
        }
    }

    private void RenderLine(Graphics g, Line line, (double Scale, double OffsetX, double OffsetY) transform, Pen pen, int canvasHeight)
    {
        try
        {
            var startPoint = TransformPoint(line.StartPoint, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight);
            var endPoint = TransformPoint(line.EndPoint, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight);
            
            if (!IsValidPoint(startPoint) || !IsValidPoint(endPoint))
                return;
            
            g.DrawLine(pen, startPoint, endPoint);
        }
        catch (OutOfMemoryException)
        {
        }
    }

    private void RenderCircle(Graphics g, Circle circle, (double Scale, double OffsetX, double OffsetY) transform, Pen pen, int canvasHeight)
    {
        try
        {
            var center = TransformPoint(circle.Center, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight);
            var radius = (float)(circle.Radius * transform.Scale);
            
            if (!IsValidPoint(center) || radius < MinGeometrySize)
                return;
            
            g.DrawEllipse(pen, center.X - radius, center.Y - radius, radius * 2, radius * 2);
        }
        catch (OutOfMemoryException)
        {
        }
    }

    private void RenderArc(Graphics g, Arc arc, (double Scale, double OffsetX, double OffsetY) transform, Pen pen, int canvasHeight)
    {
        try
        {
            var center = TransformPoint(arc.Center, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight);
            var radius = (float)(arc.Radius * transform.Scale);
            
            if (!IsValidPoint(center) || radius < MinGeometrySize)
                return;
            
            var startAngle = (float)arc.StartAngle;
            var endAngle = (float)arc.EndAngle;
            var sweepAngle = endAngle - startAngle;
            if (sweepAngle < 0) sweepAngle += 360;
            
            g.DrawArc(pen, center.X - radius, center.Y - radius, radius * 2, radius * 2, -startAngle, -sweepAngle);
        }
        catch (OutOfMemoryException)
        {
        }
    }

    private void RenderPolyline2D(Graphics g, Polyline2D polyline2D, (double Scale, double OffsetX, double OffsetY) transform, Pen pen, int canvasHeight)
    {
        if (polyline2D.SmoothType == PolylineSmoothType.NoSmooth)
            DrawPolyline2D(g, polyline2D, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight, pen);
        else
            DrawSplinePolyline(g, polyline2D, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight, pen);
    }

    private void RenderPolyline3D(Graphics g, Polyline3D polyline3D, (double Scale, double OffsetX, double OffsetY) transform, Pen pen, int canvasHeight)
    {
        if (polyline3D.SmoothType == PolylineSmoothType.NoSmooth)
        {
            var points = polyline3D.Vertexes.Select(v =>
                TransformPoint(new Vector3(v.X, v.Y, v.Z), transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight)).ToArray();
            g.DrawLines(pen, points);
        }
        else
        {
            DrawSplinePolyline(g, polyline3D, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight, pen);
        }
    }

    private void RenderSpline(Graphics g, Spline spline, (double Scale, double OffsetX, double OffsetY) transform, Pen pen, int canvasHeight)
    {
        var splineVertices = InterpolateSpline(spline.ControlPoints.ToList(), spline.Degree, spline.Knots.ToList(), SplineSubdivisions);
        if (splineVertices.Count > 1)
        {
            var points = splineVertices.Select(v => TransformPoint(v, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight)).ToArray();
            g.DrawCurve(pen, points);
        }
    }

    private void RenderEllipse(Graphics g, Ellipse ellipse, (double Scale, double OffsetX, double OffsetY) transform, Pen pen, int canvasHeight)
    {
        try
        {
            var center = TransformPoint(ellipse.Center, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight);
            var majorAxis = (float)(ellipse.MajorAxis * transform.Scale);
            var minorAxis = (float)(ellipse.MinorAxis * transform.Scale);
            
            if (!IsValidPoint(center) || majorAxis < MinGeometrySize || minorAxis < MinGeometrySize)
                return;
            
            var majorAxisVector = ellipse.MajorAxis * ellipse.Normal;
            var rotation = (float)Math.Atan2(majorAxisVector.Y, majorAxisVector.X);

            g.TranslateTransform(center.X, center.Y);
            g.RotateTransform(rotation * 180 / (float)Math.PI);
            g.DrawEllipse(pen, -majorAxis / 2, -minorAxis / 2, majorAxis, minorAxis);
            g.ResetTransform();
        }
        catch (OutOfMemoryException)
        {
            g.ResetTransform();
        }
    }

    private void RenderGenericPolyline(Graphics g, EntityObject entity, (double Scale, double OffsetX, double OffsetY) transform, Pen pen, int canvasHeight)
    {
        var vertexes = GetPolylineVertexes(entity);
        var smoothType = GetPolylineSmoothType(entity);

        if (smoothType == PolylineSmoothType.NoSmooth)
        {
            for (var i = 0; i < vertexes.Count; i++)
            {
                var vertex = vertexes[i];
                var nextVertex = vertexes[(i + 1) % vertexes.Count];
                var startPoint = TransformPoint(new Vector3(vertex.X, vertex.Y, 0), transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight);
                var endPoint = TransformPoint(new Vector3(nextVertex.X, nextVertex.Y, 0), transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight);
                g.DrawLine(pen, startPoint, endPoint);
            }
        }
        else
        {
            DrawSplinePolyline(g, vertexes, transform.Scale, transform.OffsetX, transform.OffsetY, canvasHeight, pen);
        }
    }
    
    private void DrawPolyline2D(Graphics g, Polyline2D polyline2D, double scale, double offsetX, double offsetY,
        int canvasHeight, Pen pen)
    {
        for (var i = 0; i < polyline2D.Vertexes.Count; i++)
        {
            var vertex = polyline2D.Vertexes[i];
            var nextVertex = polyline2D.Vertexes[(i + 1) % polyline2D.Vertexes.Count];
            
            if (vertex.Bulge != 0)
            {
                try
                {
                    var arcSegment = GetArcSegmentFromBulge(vertex, nextVertex, scale, offsetX, offsetY, canvasHeight);
                    if (arcSegment.Radius >= MinGeometrySize && IsValidPoint(arcSegment.StartPoint) && IsValidPoint(arcSegment.EndPoint))
                        g.DrawArc(pen, arcSegment.Rect, arcSegment.StartAngle, arcSegment.SweepAngle);
                }
                catch (OutOfMemoryException)
                {
                }
            }
            else
            {
                DrawPolylineSegment(g, vertex.Position, nextVertex.Position, scale, offsetX, offsetY, canvasHeight, pen);
            }
        }
    }

    private void DrawPolylineSegment(Graphics g, Vector2 startPos, Vector2 endPos, double scale, double offsetX, double offsetY, int canvasHeight, Pen pen)
    {
        try
        {
            var startPoint = TransformPoint(new Vector3(startPos.X, startPos.Y, 0), scale, offsetX, offsetY, canvasHeight);
            var endPoint = TransformPoint(new Vector3(endPos.X, endPos.Y, 0), scale, offsetX, offsetY, canvasHeight);
            
            if (IsValidPoint(startPoint) && IsValidPoint(endPoint))
                g.DrawLine(pen, startPoint, endPoint);
        }
        catch (OutOfMemoryException)
        {
        }
    }

    private void DrawSplinePolyline(Graphics g, Polyline2D polyline2D, double scale, double offsetX, double offsetY,
        int canvasHeight, Pen pen)
    {
        try
        {
            var points = polyline2D.Vertexes.Select(v => TransformPoint(v.Position, scale, offsetX, offsetY, canvasHeight)).ToArray();
            if (points.Length > 1 && points.All(IsValidPoint))
                g.DrawCurve(pen, points);
        }
        catch (OutOfMemoryException)
        {
        }
    }

    private void DrawSplinePolyline(Graphics g, Polyline3D polyline3D, double scale, double offsetX, double offsetY,
        int canvasHeight, Pen pen)
    {
        try
        {
            var points = polyline3D.Vertexes.Select(v => TransformPoint(new Vector3(v.X, v.Y, v.Z), scale, offsetX, offsetY, canvasHeight)).ToArray();
            if (points.Length > 1 && points.All(IsValidPoint))
                g.DrawCurve(pen, points);
        }
        catch (OutOfMemoryException)
        {
        }
    }

    private void DrawSplinePolyline(Graphics g, List<Vector3> polylineVertexes, double scale, double offsetX,
        double offsetY, int canvasHeight, Pen pen)
    {
        try
        {
            var points = polylineVertexes.Select(v => TransformPoint(v, scale, offsetX, offsetY, canvasHeight)).ToArray();
            if (points.Length > 1 && points.All(IsValidPoint))
                g.DrawCurve(pen, points);
        }
        catch (OutOfMemoryException)
        {
        }
    }
    
    private (PointF StartPoint, PointF EndPoint, float Radius, float StartAngle, float SweepAngle, RectangleF Rect)
        GetArcSegmentFromBulge(Polyline2DVertex startVertex, Polyline2DVertex endVertex, double scale, double offsetX,
            double offsetY, int canvasHeight)
    {
        var startPoint = TransformPoint(new Vector3(startVertex.Position.X, startVertex.Position.Y, 0), scale, offsetX,
            offsetY, canvasHeight);
        var endPoint = TransformPoint(new Vector3(endVertex.Position.X, endVertex.Position.Y, 0), scale, offsetX,
            offsetY, canvasHeight);

        var bulge = startVertex.Bulge;
        var chordLength = Math.Sqrt(Math.Pow(endVertex.Position.X - startVertex.Position.X, 2) +
                                    Math.Pow(endVertex.Position.Y - startVertex.Position.Y, 2));
        var sagitta = Math.Abs(bulge) * chordLength / 2;
        var radius = (Math.Pow(chordLength / 2, 2) + Math.Pow(sagitta, 2)) / (2 * sagitta);

        var theta = 4 * Math.Atan(Math.Abs(bulge));
        var gamma = (Math.PI - theta) / 2;
        var phi = Math.Atan2(endVertex.Position.Y - startVertex.Position.Y,
            endVertex.Position.X - startVertex.Position.X);
        var centerAngle = bulge > 0 ? phi + gamma : phi - gamma;

        var centerX = startVertex.Position.X + radius * Math.Cos(centerAngle);
        var centerY = startVertex.Position.Y + radius * Math.Sin(centerAngle);

        var transformedCenter = TransformPoint(new Vector3(centerX, centerY, 0), scale, offsetX, offsetY, canvasHeight);
        var transformedRadius = (float)(radius * scale);

        var startAngle = Math.Atan2(startVertex.Position.Y - centerY, startVertex.Position.X - centerX);
        var endAngle = Math.Atan2(endVertex.Position.Y - centerY, endVertex.Position.X - centerX);
        var sweepAngle = bulge > 0 ? endAngle - startAngle : startAngle - endAngle;
        if (sweepAngle < 0) sweepAngle += 2 * Math.PI;

        startAngle = -startAngle * 180 / Math.PI;
        sweepAngle = -sweepAngle * 180 / Math.PI;

        var startAngleDeg = (float)startAngle;
        var sweepAngleDeg = (float)sweepAngle;

        if (bulge < 0) sweepAngleDeg = -sweepAngleDeg;

        var rect = new RectangleF(
            transformedCenter.X - transformedRadius,
            transformedCenter.Y - transformedRadius,
            2 * transformedRadius,
            2 * transformedRadius
        );

        return (
            startPoint,
            endPoint,
            transformedRadius,
            startAngleDeg,
            sweepAngleDeg,
            rect
        );
    }
    
    private PointF TransformPoint(Vector3 point, double scale, double offsetX, double offsetY, int canvasHeight)
    {
        var x = (float)(point.X * scale + offsetX);
        var y = canvasHeight - (float)(point.Y * scale + offsetY);
        return new PointF(x, y);
    }

    private PointF TransformPoint(Vector2 point, double scale, double offsetX, double offsetY, int canvasHeight)
    {
        return new PointF(
            (float)(point.X * scale + offsetX),
            canvasHeight - (float)(point.Y * scale + offsetY)
        );
    }

    private bool IsValidPoint(PointF point)
    {
        return float.IsFinite(point.X) && float.IsFinite(point.Y) &&
               Math.Abs(point.X) < 10000 && Math.Abs(point.Y) < 10000;
    }
}

public class SvgRenderer : BaseRenderer
{
    public override object Render(DxfDocument dxf, int width, int height)
    {
        var entities = BaseRenderer.CollectEntities(dxf);
        if (entities.Count == 0) 
            return CreateEmptySvg(width, height);

        var bounds = CalculateBoundingBox(entities);
        return Render(entities, bounds, width, height);
    }
    
    public override object Render(List<EntityObject> entities, (double MinX, double MinY, double MaxX, double MaxY) bounds, int width, int height)
    {
        if (bounds.MinX >= bounds.MaxX || bounds.MinY >= bounds.MaxY)
            return CreateEmptySvg(width, height);

        var drawingWidth = bounds.MaxX - bounds.MinX;
        var drawingHeight = bounds.MaxY - bounds.MinY;
        
        var padding = Math.Max(drawingWidth, drawingHeight) * 0.1;
        var viewMinX = bounds.MinX - padding;
        var viewMinY = bounds.MinY - padding; 
        var viewWidth = drawingWidth + 2 * padding;
        var viewHeight = drawingHeight + 2 * padding;

        var avgDimension = (viewWidth + viewHeight) / 2;
        var strokeWidth = avgDimension * 0.0225;
        strokeWidth = Math.Max(avgDimension * 0.015, Math.Min(avgDimension * 0.0375, strokeWidth));

        var culture = CultureInfo.InvariantCulture;
        var svg = new StringBuilder();
        
        svg.AppendLine($"<svg width=\"{width}\" height=\"{height}\" viewBox=\"{viewMinX.ToString("F2", culture)} {viewMinY.ToString("F2", culture)} {viewWidth.ToString("F2", culture)} {viewHeight.ToString("F2", culture)}\" xmlns=\"http://www.w3.org/2000/svg\" shape-rendering=\"geometricPrecision\">");
        svg.AppendLine($"  <rect x=\"{viewMinX.ToString("F2", culture)}\" y=\"{viewMinY.ToString("F2", culture)}\" width=\"{viewWidth.ToString("F2", culture)}\" height=\"{viewHeight.ToString("F2", culture)}\" fill=\"white\"/>");
        svg.AppendLine($"  <g transform=\"scale(1,-1) translate(0,{(-(viewMinY * 2 + viewHeight)).ToString("F2", culture)})\" stroke=\"#000000\" stroke-width=\"{strokeWidth.ToString("F3", culture)}\" fill=\"none\" stroke-linecap=\"round\" stroke-linejoin=\"round\" shape-rendering=\"geometricPrecision\">");

        foreach (var entity in entities)
        {
            var color = GetEntityColorHex(entity);
            RenderEntityToSvg(svg, entity, color, culture);
        }

        svg.AppendLine("  </g>");
        svg.AppendLine("</svg>");

        return svg.ToString();
    }
    
    private string CreateEmptySvg(int width, int height)
    {
        return $@"<svg width=""{width}"" height=""{height}"" viewBox=""0 0 {width} {height}"" xmlns=""http://www.w3.org/2000/svg"">
  <rect width=""100%"" height=""100%"" fill=""white""/>
</svg>";
    }
    
    private void RenderEntityToSvg(StringBuilder svg, EntityObject entity, string color, CultureInfo culture)
    {
        switch (entity)
        {
            case Line line:
                svg.AppendLine($"    <line x1=\"{line.StartPoint.X.ToString("F4", culture)}\" y1=\"{line.StartPoint.Y.ToString("F4", culture)}\" x2=\"{line.EndPoint.X.ToString("F4", culture)}\" y2=\"{line.EndPoint.Y.ToString("F4", culture)}\" stroke=\"{color}\"/>");
                break;
            case Circle circle:
                svg.AppendLine($"    <circle cx=\"{circle.Center.X.ToString("F4", culture)}\" cy=\"{circle.Center.Y.ToString("F4", culture)}\" r=\"{circle.Radius.ToString("F4", culture)}\" stroke=\"{color}\" fill=\"none\"/>");
                break;
            case Arc arc:
                RenderArcToSvg(svg, arc, color, culture);
                break;
            case Polyline2D polyline2D:
                RenderPolyline2DToSvg(svg, polyline2D, color, culture);
                break;
            case Polyline3D polyline3D:
                RenderPolyline3DToSvg(svg, polyline3D, color, culture);
                break;
            case Spline spline:
                RenderSplineToSvg(svg, spline, color, culture);
                break;
            case Ellipse ellipse:
                RenderEllipseToSvg(svg, ellipse, color, culture);
                break;
            default:
                if (entity.GetType().Name == "Polyline")
                    RenderGenericPolylineToSvg(svg, entity, color, culture);
                break;
        }
    }

    private void RenderArcToSvg(StringBuilder svg, Arc arc, string color, CultureInfo culture)
    {
        var startAngleRad = arc.StartAngle * Math.PI / 180;
        var endAngleRad = arc.EndAngle * Math.PI / 180;
        
        var startX = arc.Center.X + arc.Radius * Math.Cos(startAngleRad);
        var startY = arc.Center.Y + arc.Radius * Math.Sin(startAngleRad);
        var endX = arc.Center.X + arc.Radius * Math.Cos(endAngleRad);
        var endY = arc.Center.Y + arc.Radius * Math.Sin(endAngleRad);
        
        var sweepAngle = arc.EndAngle - arc.StartAngle;
        if (sweepAngle < 0) sweepAngle += 360;
        
        var largeArcFlag = sweepAngle > 180 ? 1 : 0;
        var sweepFlag = 1;
        
        svg.AppendLine($"    <path d=\"M {startX.ToString("F4", culture)},{startY.ToString("F4", culture)} A {arc.Radius.ToString("F4", culture)},{arc.Radius.ToString("F4", culture)} 0 {largeArcFlag},{sweepFlag} {endX.ToString("F4", culture)},{endY.ToString("F4", culture)}\" stroke=\"{color}\" fill=\"none\"/>");
    }

    private void RenderPolyline2DToSvg(StringBuilder svg, Polyline2D polyline2D, string color, CultureInfo culture)
    {
        if (polyline2D.Vertexes.Count == 0) return;

        var pathData = new StringBuilder("M ");
        var firstVertex = polyline2D.Vertexes[0];
        pathData.Append($"{firstVertex.Position.X.ToString("F4", culture)},{firstVertex.Position.Y.ToString("F4", culture)}");

        var vertexCount = polyline2D.IsClosed ? polyline2D.Vertexes.Count : polyline2D.Vertexes.Count - 1;
        
        for (var i = 0; i < vertexCount; i++)
        {
            var currentVertex = polyline2D.Vertexes[i];
            var nextIndex = (i + 1) % polyline2D.Vertexes.Count;
            var nextVertex = polyline2D.Vertexes[nextIndex];
            
            if (i == 0 && currentVertex.Position.X == firstVertex.Position.X && 
                currentVertex.Position.Y == firstVertex.Position.Y)
            {
            }
            
            if (Math.Abs(currentVertex.Bulge) > 1e-9)
            {
                var bulge = currentVertex.Bulge;
                var chordLength = Math.Sqrt(Math.Pow(nextVertex.Position.X - currentVertex.Position.X, 2) +
                                           Math.Pow(nextVertex.Position.Y - currentVertex.Position.Y, 2));
                var sagitta = Math.Abs(bulge) * chordLength / 2;
                var radius = (Math.Pow(chordLength / 2, 2) + Math.Pow(sagitta, 2)) / (2 * sagitta);
                
                var theta = 4 * Math.Atan(Math.Abs(bulge));
                var sweepAngle = theta * 180 / Math.PI;
                
                var largeArcFlag = sweepAngle > 180 ? 1 : 0;
                var sweepFlag = bulge > 0 ? 1 : 0;
                
                pathData.Append($" A {radius.ToString("F4", culture)},{radius.ToString("F4", culture)} 0 {largeArcFlag},{sweepFlag} {nextVertex.Position.X.ToString("F4", culture)},{nextVertex.Position.Y.ToString("F4", culture)}");
            }
            else
            {
                pathData.Append($" L {nextVertex.Position.X.ToString("F4", culture)},{nextVertex.Position.Y.ToString("F4", culture)}");
            }
        }

        if (polyline2D.IsClosed)
            pathData.Append(" Z");

        svg.AppendLine($"    <path d=\"{pathData}\" stroke=\"{color}\" fill=\"none\"/>");
    }

    private void RenderPolyline3DToSvg(StringBuilder svg, Polyline3D polyline3D, string color, CultureInfo culture)
    {
        if (polyline3D.Vertexes.Count == 0) return;

        var pathData = new StringBuilder("M ");
        var firstVertex = polyline3D.Vertexes[0];
        pathData.Append($"{firstVertex.X.ToString("F4", culture)},{firstVertex.Y.ToString("F4", culture)}");

        for (var i = 1; i < polyline3D.Vertexes.Count; i++)
        {
            var vertex = polyline3D.Vertexes[i];
            pathData.Append($" L {vertex.X.ToString("F4", culture)},{vertex.Y.ToString("F4", culture)}");
        }

        if (polyline3D.IsClosed)
            pathData.Append(" Z");

        svg.AppendLine($"    <path d=\"{pathData}\" stroke=\"{color}\" fill=\"none\"/>");
    }

    private void RenderSplineToSvg(StringBuilder svg, Spline spline, string color, CultureInfo culture)
    {
        var splineVertices = InterpolateSpline(spline.ControlPoints.ToList(), spline.Degree, spline.Knots.ToList(), SplineSubdivisions);
        if (splineVertices.Count < 2) return;

        var pathData = new StringBuilder("M ");
        var firstVertex = splineVertices[0];
        pathData.Append($"{firstVertex.X.ToString("F4", culture)},{firstVertex.Y.ToString("F4", culture)}");

        for (var i = 1; i < splineVertices.Count; i++)
        {
            var vertex = splineVertices[i];
            pathData.Append($" L {vertex.X.ToString("F4", culture)},{vertex.Y.ToString("F4", culture)}");
        }

        svg.AppendLine($"    <path d=\"{pathData}\" stroke=\"{color}\" fill=\"none\"/>");
    }

    private void RenderGenericPolylineToSvg(StringBuilder svg, EntityObject entity, string color, CultureInfo culture)
    {
        var vertexes = GetPolylineVertexes(entity);
        if (vertexes.Count == 0) return;

        var pathData = new StringBuilder("M ");
        var firstVertex = vertexes[0];
        pathData.Append($"{firstVertex.X.ToString("F4", culture)},{firstVertex.Y.ToString("F4", culture)}");

        for (var i = 1; i < vertexes.Count; i++)
        {
            var vertex = vertexes[i];
            pathData.Append($" L {vertex.X.ToString("F4", culture)},{vertex.Y.ToString("F4", culture)}");
        }

        svg.AppendLine($"    <path d=\"{pathData}\" stroke=\"{color}\" fill=\"none\"/>");
    }

    private void RenderEllipseToSvg(StringBuilder svg, Ellipse ellipse, string color, CultureInfo culture)
    {
        var majorAxis = ellipse.MajorAxis;
        var minorAxis = ellipse.MinorAxis;
        var center = ellipse.Center;
        
        var majorAxisVector = ellipse.MajorAxis * ellipse.Normal;
        var rotationDegrees = Math.Atan2(majorAxisVector.Y, majorAxisVector.X) * 180 / Math.PI;
        
        svg.AppendLine($"    <ellipse cx=\"{center.X.ToString("F4", culture)}\" cy=\"{center.Y.ToString("F4", culture)}\" rx=\"{majorAxis.ToString("F4", culture)}\" ry=\"{minorAxis.ToString("F4", culture)}\" transform=\"rotate({rotationDegrees.ToString("F2", culture)} {center.X.ToString("F4", culture)} {center.Y.ToString("F4", culture)})\" stroke=\"{color}\" fill=\"none\"/>");
    }
}

public class DxfThumbnailGenerator
{
    private readonly BitmapRenderer _bitmapRenderer = new();
    private readonly SvgRenderer _svgRenderer = new();
    
    public Bitmap GenerateBitmap(string filePath, int width = 0, int height = 0)
    {
        try
        {
            var dxf = DxfDocument.Load(filePath);
            var w = width > 0 ? width : BaseRenderer.DefaultThumbnailSize;
            var h = height > 0 ? height : BaseRenderer.DefaultThumbnailSize;
            return (Bitmap)_bitmapRenderer.Render(dxf, w, h);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error loading DXF file for bitmap generation: {ex.Message}", ex);
        }
    }

    public string GenerateSvg(string filePath, int width = 0, int height = 0)
    {
        try
        {
            var dxf = DxfDocument.Load(filePath);
            var w = width > 0 ? width : BaseRenderer.DefaultThumbnailSize;
            var h = height > 0 ? height : BaseRenderer.DefaultThumbnailSize;
            return (string)_svgRenderer.Render(dxf, w, h);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error loading DXF file for SVG generation: {ex.Message}", ex);
        }
    }    
}

