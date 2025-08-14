using System.Drawing.Drawing2D;
using netDxf;
using netDxf.Entities;
using netDxf.Header;

namespace DxfGenerator;

public class DxfThumbnailGenerator
{
    private const int DefaultThumbnailSize = 100;
    private const double BoundsPadding = 0.9;
    private const int SplineSubdivisions = 50;
    private const float DefaultPenWidth = 1f;
    private static readonly double[] CardinalAngles = { 0, 90, 180, 270 };
    private static readonly Color DefaultEntityColor = Color.Black;
    public Bitmap GenerateThumbnail(string filePath)
    {
        try
        {
            var dxf = DxfDocument.Load(filePath);
            return RenderDxfToBitmap(dxf);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error loading DXF file: {ex.Message}", ex);
        }
    }


    private Bitmap RenderDxfToBitmap(DxfDocument dxf)
    {
        var bitmap = new Bitmap(DefaultThumbnailSize, DefaultThumbnailSize);

        using var g = Graphics.FromImage(bitmap);
        ConfigureGraphics(g);

        var entities = CollectEntities(dxf);
        if (entities.Count == 0) return bitmap;

        var bounds = CalculateBoundingBox(entities);
        var transform = CalculateTransform(bounds, DefaultThumbnailSize, DefaultThumbnailSize);

        RenderEntities(g, entities, transform);
        
        return bitmap;
    }

    private static void ConfigureGraphics(Graphics g)
    {
        g.Clear(Color.White);
        g.SmoothingMode = SmoothingMode.AntiAlias;
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

    private (double MinX, double MinY, double MaxX, double MaxY) CalculateBoundingBox(IEnumerable<EntityObject> entities)
    {
        double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
        foreach (var entity in entities)
            UpdateBounds(entity, ref minX, ref minY, ref maxX, ref maxY);
        return (minX, minY, maxX, maxY);
    }

    private (double Scale, double OffsetX, double OffsetY) CalculateTransform(
        (double MinX, double MinY, double MaxX, double MaxY) bounds, int width, int height)
    {
        var scaleX = width / (bounds.MaxX - bounds.MinX);
        var scaleY = height / (bounds.MaxY - bounds.MinY);
        var scale = Math.Min(scaleX, scaleY) * BoundsPadding;
        var offsetX = (width - (bounds.MaxX - bounds.MinX) * scale) / 2 - bounds.MinX * scale;
        var offsetY = (height - (bounds.MaxY - bounds.MinY) * scale) / 2 - bounds.MinY * scale;
        return (scale, offsetX, offsetY);
    }

    private void RenderEntities(Graphics g, IEnumerable<EntityObject> entities, 
        (double Scale, double OffsetX, double OffsetY) transform)
    {
        foreach (var entity in entities)
        {
            using var pen = new Pen(GetEntityColor(entity), DefaultPenWidth);
            RenderEntity(g, entity, transform, pen);
        }
    }

    private void RenderEntity(Graphics g, EntityObject entity, 
        (double Scale, double OffsetX, double OffsetY) transform, Pen pen)
    {
        switch (entity)
        {
            case Line line:
                RenderLine(g, line, transform, pen);
                break;
            case Circle circle:
                RenderCircle(g, circle, transform, pen);
                break;
            case Arc arc:
                RenderArc(g, arc, transform, pen);
                break;
            case Polyline2D polyline2D:
                RenderPolyline2D(g, polyline2D, transform, pen);
                break;
            case Polyline3D polyline3D:
                RenderPolyline3D(g, polyline3D, transform, pen);
                break;
            case Spline spline:
                RenderSpline(g, spline, transform, pen);
                break;
            case Ellipse ellipse:
                RenderEllipse(g, ellipse, transform, pen);
                break;
            default:
                if (entity.GetType().Name == "Polyline")
                    RenderGenericPolyline(g, entity, transform, pen);
                break;
        }
    }

    private void RenderLine(Graphics g, Line line, (double Scale, double OffsetX, double OffsetY) transform, Pen pen)
    {
        var startPoint = TransformPoint(line.StartPoint, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize);
        var endPoint = TransformPoint(line.EndPoint, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize);
        g.DrawLine(pen, startPoint, endPoint);
    }

    private void RenderCircle(Graphics g, Circle circle, (double Scale, double OffsetX, double OffsetY) transform, Pen pen)
    {
        var center = TransformPoint(circle.Center, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize);
        var radius = (float)(circle.Radius * transform.Scale);
        g.DrawEllipse(pen, center.X - radius, center.Y - radius, radius * 2, radius * 2);
    }

    private void RenderArc(Graphics g, Arc arc, (double Scale, double OffsetX, double OffsetY) transform, Pen pen)
    {
        var center = TransformPoint(arc.Center, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize);
        var radius = (float)(arc.Radius * transform.Scale);
        var startAngle = (float)arc.StartAngle;
        var endAngle = (float)arc.EndAngle;
        var sweepAngle = endAngle - startAngle;
        if (sweepAngle < 0) sweepAngle += 360;
        g.DrawArc(pen, center.X - radius, center.Y - radius, radius * 2, radius * 2, -startAngle, -sweepAngle);
    }

    private void RenderPolyline2D(Graphics g, Polyline2D polyline2D, (double Scale, double OffsetX, double OffsetY) transform, Pen pen)
    {
        if (polyline2D.SmoothType == PolylineSmoothType.NoSmooth)
            DrawPolyline2D(g, polyline2D, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize, pen);
        else
            DrawSplinePolyline(g, polyline2D, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize, pen);
    }

    private void RenderPolyline3D(Graphics g, Polyline3D polyline3D, (double Scale, double OffsetX, double OffsetY) transform, Pen pen)
    {
        if (polyline3D.SmoothType == PolylineSmoothType.NoSmooth)
        {
            var points = polyline3D.Vertexes.Select(v =>
                TransformPoint(new Vector3(v.X, v.Y, v.Z), transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize)).ToArray();
            g.DrawLines(pen, points);
        }
        else
        {
            DrawSplinePolyline(g, polyline3D, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize, pen);
        }
    }

    private void RenderSpline(Graphics g, Spline spline, (double Scale, double OffsetX, double OffsetY) transform, Pen pen)
    {
        var splineVertices = InterpolateSpline(spline.ControlPoints.ToList(), spline.Degree, spline.Knots.ToList(), SplineSubdivisions);
        if (splineVertices.Count > 1)
        {
            var points = splineVertices.Select(v => TransformPoint(v, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize)).ToArray();
            g.DrawCurve(pen, points);
        }
    }

    private void RenderEllipse(Graphics g, Ellipse ellipse, (double Scale, double OffsetX, double OffsetY) transform, Pen pen)
    {
        var center = TransformPoint(ellipse.Center, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize);
        var majorAxis = (float)(ellipse.MajorAxis * transform.Scale);
        var minorAxis = (float)(ellipse.MinorAxis * transform.Scale);
        var majorAxisVector = ellipse.MajorAxis * ellipse.Normal;
        var rotation = (float)Math.Atan2(majorAxisVector.Y, majorAxisVector.X);

        g.TranslateTransform(center.X, center.Y);
        g.RotateTransform(rotation * 180 / (float)Math.PI);
        g.DrawEllipse(pen, -majorAxis / 2, -minorAxis / 2, majorAxis, minorAxis);
        g.ResetTransform();
    }

    private void RenderGenericPolyline(Graphics g, EntityObject entity, (double Scale, double OffsetX, double OffsetY) transform, Pen pen)
    {
        var vertexes = GetPolylineVertexes(entity);
        var smoothType = GetPolylineSmoothType(entity);

        if (smoothType == PolylineSmoothType.NoSmooth)
        {
            for (var i = 0; i < vertexes.Count; i++)
            {
                var vertex = vertexes[i];
                var nextVertex = vertexes[(i + 1) % vertexes.Count];
                var startPoint = TransformPoint(new Vector3(vertex.X, vertex.Y, 0), transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize);
                var endPoint = TransformPoint(new Vector3(nextVertex.X, nextVertex.Y, 0), transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize);
                g.DrawLine(pen, startPoint, endPoint);
            }
        }
        else
        {
            DrawSplinePolyline(g, vertexes, transform.Scale, transform.OffsetX, transform.OffsetY, DefaultThumbnailSize, pen);
        }
    }

    private PolylineSmoothType GetPolylineSmoothType(EntityObject entity)
    {
        var smoothTypeProperty = entity.GetType().GetProperty("SmoothType");
        if (smoothTypeProperty?.GetValue(entity) is PolylineSmoothType smoothType)
            return smoothType;
        return PolylineSmoothType.NoSmooth;
    }

    private void UpdateBounds(EntityObject entity, ref double minX, ref double minY, ref double maxX, ref double maxY)
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
            // Вычисляем начальную и конечную точки арки
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
                UpdateBoundsForPoint(new Vector3(vertex.Position.X, vertex.Position.Y, 0), ref minX, ref minY, ref maxX,
                    ref maxY);

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

    private (Vector2 center, double radius, double startAngle, double endAngle)? GetArcGeomFromBulge(
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

    private bool IsAngleBetween(double angle, double startAngle, double endAngle)
    {
        angle = (angle + 360) % 360;
        startAngle = (startAngle + 360) % 360;
        endAngle = (endAngle + 360) % 360;

        if (startAngle < endAngle)
            return startAngle <= angle && angle <= endAngle;
        return startAngle <= angle || angle <= endAngle;
    }

    private void UpdateBoundsForPoint(Vector3 point, ref double minX, ref double minY, ref double maxX, ref double maxY)
    {
        minX = Math.Min(minX, point.X);
        minY = Math.Min(minY, point.Y);
        maxX = Math.Max(maxX, point.X);
        maxY = Math.Max(maxY, point.Y);
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

    private Color GetEntityColor(EntityObject entity)
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

    private void DrawPolyline2D(Graphics g, Polyline2D polyline2D, double scale, double offsetX, double offsetY,
        int canvasHeight, Pen pen)
    {
        for (var i = 0; i < polyline2D.Vertexes.Count; i++)
        {
            var vertex = polyline2D.Vertexes[i];
            var nextVertex = polyline2D.Vertexes[(i + 1) % polyline2D.Vertexes.Count];
            
            if (vertex.Bulge != 0)
            {
                var arcSegment = GetArcSegmentFromBulge(vertex, nextVertex, scale, offsetX, offsetY, canvasHeight);
                g.DrawArc(pen, arcSegment.Rect, arcSegment.StartAngle, arcSegment.SweepAngle);
            }
            else
            {
                DrawPolylineSegment(g, vertex.Position, nextVertex.Position, scale, offsetX, offsetY, canvasHeight, pen);
            }
        }
    }

    private void DrawPolylineSegment(Graphics g, Vector2 startPos, Vector2 endPos, double scale, double offsetX, double offsetY, int canvasHeight, Pen pen)
    {
        var startPoint = TransformPoint(new Vector3(startPos.X, startPos.Y, 0), scale, offsetX, offsetY, canvasHeight);
        var endPoint = TransformPoint(new Vector3(endPos.X, endPos.Y, 0), scale, offsetX, offsetY, canvasHeight);
        g.DrawLine(pen, startPoint, endPoint);
    }

    private void DrawSplinePolyline(Graphics g, Polyline2D polyline2D, double scale, double offsetX, double offsetY,
        int canvasHeight, Pen pen)
    {
        var points = polyline2D.Vertexes.Select(v => TransformPoint(v.Position, scale, offsetX, offsetY, canvasHeight)).ToArray();
        g.DrawCurve(pen, points);
    }

    private void DrawSplinePolyline(Graphics g, Polyline3D polyline3D, double scale, double offsetX, double offsetY,
        int canvasHeight, Pen pen)
    {
        var points = polyline3D.Vertexes.Select(v => TransformPoint(new Vector3(v.X, v.Y, v.Z), scale, offsetX, offsetY, canvasHeight)).ToArray();
        g.DrawCurve(pen, points);
    }

    private void DrawSplinePolyline(Graphics g, List<Vector3> polylineVertexes, double scale, double offsetX,
        double offsetY, int canvasHeight, Pen pen)
    {
        var points = polylineVertexes.Select(v => TransformPoint(v, scale, offsetX, offsetY, canvasHeight)).ToArray();
        g.DrawCurve(pen, points);
    }

    private List<Vector3> GetPolylineVertexes(EntityObject polyline)
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

    private List<Vector3> InterpolateSpline(List<Vector3> controlPoints, int degree, List<double> knotValues,
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

    private Vector3 DeBoor(double t, int degree, List<Vector3> controlPoints, List<double> knotValues)
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
}

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

    public static void OptimizeDxfFile(string dxfFilePath, string acadVersion)
    {
        try
        {
            var dxf = DxfDocument.Load(dxfFilePath);
            ReplaceWithOptimizedVersion(dxf, dxfFilePath, acadVersion);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Ошибка оптимизации DXF: {ex.Message}");
        }
    }

    private static void ReplaceWithOptimizedVersion(DxfDocument dxf, string originalFilePath, string acadVersion)
    {
        try
        {
            if (!AcadVersionToDxfVersion.TryGetValue(acadVersion, out var dxfVersion))
            {
                System.Diagnostics.Debug.WriteLine($"Неподдерживаемая версия AutoCAD: {acadVersion}, оптимизация пропущена");
                return;
            }

            var optimizedDxf = new DxfDocument();
            optimizedDxf.DrawingVariables.AcadVer = dxfVersion;
            
            foreach (var entity in dxf.Entities.All)
                optimizedDxf.Entities.Add((EntityObject)entity.Clone());

            optimizedDxf.Save(originalFilePath);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Ошибка оптимизации {acadVersion}: {ex.Message}");
        }
    }

}