using netDxf;
using netDxf.Entities;

namespace DxfThumbnailGenerator
{
    public class DxfThumbnailGenerator
    {
        public Bitmap GenerateThumbnail(string filePath)
        {
            try
            {
                DxfDocument dxf = DxfDocument.Load(filePath);
                return RenderDxfToBitmap(dxf);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error loading DXF file: {ex.Message}", ex);
            }
        }

        private Bitmap RenderDxfToBitmap(DxfDocument dxf)
        {
            int width = 100;
            int height = 100;
            Bitmap bitmap = new Bitmap(width, height);

            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                List<EntityObject> entities = new List<EntityObject>();
                entities.AddRange(dxf.Entities.Lines);
                entities.AddRange(dxf.Entities.Circles);
                entities.AddRange(dxf.Entities.Arcs);
                entities.AddRange(dxf.Entities.Polylines2D);
                entities.AddRange(dxf.Entities.Polylines3D);
                entities.AddRange(dxf.Entities.Splines);
                entities.AddRange(dxf.Entities.Ellipses);

                if (entities.Count == 0) return bitmap;

                // Calculate bounding box
                double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
                foreach (var entity in entities)
                {
                    UpdateBounds(entity, ref minX, ref minY, ref maxX, ref maxY);
                }

                // Calculate scale and offset
                double scaleX = width / (maxX - minX);
                double scaleY = height / (maxY - minY);
                double scale = Math.Min(scaleX, scaleY) * 0.9; // 90% of the available space
                double offsetX = (width - (maxX - minX) * scale) / 2 - minX * scale;
                double offsetY = (height - (maxY - minY) * scale) / 2 - minY * scale;

                foreach (var entity in entities)
                {
                    Pen pen = new Pen(GetEntityColor(entity), 1);

                    if (entity is Line line)
                    {
                        PointF startPoint = TransformPoint(line.StartPoint, scale, offsetX, offsetY, height);
                        PointF endPoint = TransformPoint(line.EndPoint, scale, offsetX, offsetY, height);
                        g.DrawLine(pen, startPoint, endPoint);
                    }
                    else if (entity is Circle circle)
                    {
                        PointF center = TransformPoint(circle.Center, scale, offsetX, offsetY, height);
                        float radius = (float)(circle.Radius * scale);
                        g.DrawEllipse(pen, center.X - radius, center.Y - radius, radius * 2, radius * 2);
                    }
                    else if (entity is Arc arc)
                    {
                        PointF center = TransformPoint(arc.Center, scale, offsetX, offsetY, height);
                        float radius = (float)(arc.Radius * scale);
                        float startAngle = (float)(arc.StartAngle);
                        float endAngle = (float)(arc.EndAngle);
                        float sweepAngle = endAngle - startAngle;
                        if (sweepAngle < 0) sweepAngle += 360;
                        g.DrawArc(pen, center.X - radius, center.Y - radius, radius * 2, radius * 2, -startAngle, -sweepAngle);
                    }
                    else if (entity is Polyline2D polyline2D)
                    {
                        if (polyline2D.SmoothType == PolylineSmoothType.NoSmooth)
                        {
                            DrawPolyline2D(g, polyline2D, scale, offsetX, offsetY, height, pen);
                        }
                        else
                        {
                            DrawSplinePolyline(g, polyline2D, scale, offsetX, offsetY, height, pen);
                        }
                    }
                    else if (entity is Polyline3D polyline3D)
                    {
                        if (polyline3D.SmoothType == PolylineSmoothType.NoSmooth)
                        {
                            PointF[] points = polyline3D.Vertexes.Select(v =>
                                TransformPoint(new Vector3(v.X, v.Y, v.Z), scale, offsetX, offsetY, height)).ToArray();
                            g.DrawLines(pen, points);
                        }
                        else
                        {
                            DrawSplinePolyline(g, polyline3D, scale, offsetX, offsetY, height, pen);
                        }
                    }
                    else if (entity is Spline spline)
                    {
                        List<Vector3> splineVertices = InterpolateSpline(spline.ControlPoints.ToList(), spline.Degree, spline.Knots.ToList(), subdivisions: 50);

                        if (splineVertices.Count > 1)
                        {
                            PointF[] points = splineVertices.Select(v => TransformPoint(v, scale, offsetX, offsetY, height)).ToArray();
                            g.DrawCurve(pen, points);
                        }
                    }

                    else if (entity is Ellipse ellipse)
                    {
                        PointF center = TransformPoint(ellipse.Center, scale, offsetX, offsetY, height);
                        float majorAxis = (float)(ellipse.MajorAxis * scale);
                        float minorAxis = (float)(ellipse.MinorAxis * scale);

                        Vector3 majorAxisVector = ellipse.MajorAxis * ellipse.Normal;
                        float rotation = (float)Math.Atan2(majorAxisVector.Y, majorAxisVector.X);

                        g.TranslateTransform(center.X, center.Y);
                        g.RotateTransform(rotation * 180 / (float)Math.PI);
                        g.DrawEllipse(pen, -majorAxis / 2, -minorAxis / 2, majorAxis, minorAxis);
                        g.ResetTransform();
                    }
                    else if (entity.GetType().Name == "Polyline")
                    {
                        var vertexes = GetPolylineVertexes(entity);

                        var smoothTypeProperty = entity.GetType().GetProperty("SmoothType");
                        var smoothType = PolylineSmoothType.NoSmooth;
                        if (smoothTypeProperty != null)
                        {
                            smoothType = (PolylineSmoothType)smoothTypeProperty.GetValue(entity);
                        }

                        if (smoothType == PolylineSmoothType.NoSmooth)
                        {
                            for (int i = 0; i < vertexes.Count; i++)
                            {
                                var vertex = vertexes[i];
                                var nextVertex = vertexes[(i + 1) % vertexes.Count];
                                PointF startPoint = TransformPoint(new Vector3(vertex.X, vertex.Y, 0), scale, offsetX, offsetY, height);
                                PointF endPoint = TransformPoint(new Vector3(nextVertex.X, nextVertex.Y, 0), scale, offsetX, offsetY, height);
                                g.DrawLine(pen, startPoint, endPoint);
                            }
                        }
                        else
                        {
                            DrawSplinePolyline(g, vertexes, scale, offsetX, offsetY, height, pen);
                        }
                    }
                }
            }
            return bitmap;
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
                double startX = arc.Center.X + arc.Radius * Math.Cos(arc.StartAngle * Math.PI / 180);
                double startY = arc.Center.Y + arc.Radius * Math.Sin(arc.StartAngle * Math.PI / 180);
                double endX = arc.Center.X + arc.Radius * Math.Cos(arc.EndAngle * Math.PI / 180);
                double endY = arc.Center.Y + arc.Radius * Math.Sin(arc.EndAngle * Math.PI / 180);

                minX = Math.Min(minX, Math.Min(startX, endX));
                minY = Math.Min(minY, Math.Min(startY, endY));
                maxX = Math.Max(maxX, Math.Max(startX, endX));
                maxY = Math.Max(maxY, Math.Max(startY, endY));

                // Проверка углов 0, 90, 180 и 270 градусов, чтобы определить, попадают ли они в арку
                double[] angles = { 0, 90, 180, 270 };
                foreach (double angle in angles)
                {
                    if (IsAngleBetween(angle, arc.StartAngle, arc.EndAngle))
                    {
                        double radian = angle * Math.PI / 180;
                        double x = arc.Center.X + arc.Radius * Math.Cos(radian);
                        double y = arc.Center.Y + arc.Radius * Math.Sin(radian);

                        minX = Math.Min(minX, x);
                        minY = Math.Min(minY, y);
                        maxX = Math.Max(maxX, x);
                        maxY = Math.Max(maxY, y);
                    }
                }
            }
            else if (entity is Polyline2D polyline2D)
            {
                foreach (var vertex in polyline2D.Vertexes)
                {
                    minX = Math.Min(minX, vertex.Position.X);
                    minY = Math.Min(minY, vertex.Position.Y);
                    maxX = Math.Max(maxX, vertex.Position.X);
                    maxY = Math.Max(maxY, vertex.Position.Y);
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
                foreach (var controlPoint in spline.ControlPoints)
                {
                    minX = Math.Min(minX, controlPoint.X);
                    minY = Math.Min(minY, controlPoint.Y);
                    maxX = Math.Max(maxX, controlPoint.X);
                    maxY = Math.Max(maxY, controlPoint.Y);
                }
            }
            else if (entity is Ellipse ellipse)
            {
                double a = ellipse.MajorAxis;
                double b = ellipse.MinorAxis;

                Vector3 majorAxisVector = ellipse.MajorAxis * ellipse.Normal;
                double rotation = Math.Atan2(majorAxisVector.Y, majorAxisVector.X);

                double cosRotation = Math.Cos(rotation);
                double sinRotation = Math.Sin(rotation);

                var points = new[]
                {
            new Vector2(a * cosRotation - b * sinRotation, a * sinRotation + b * cosRotation),
            new Vector2(-a * cosRotation - b * sinRotation, -a * sinRotation + b * cosRotation),
            new Vector2(a * cosRotation + b * sinRotation, a * sinRotation - b * cosRotation),
            new Vector2(-a * cosRotation + b * sinRotation, -a * sinRotation - b * cosRotation)
        };

                foreach (var point in points)
                {
                    UpdateBoundsForPoint(new Vector3(point.X + ellipse.Center.X, point.Y + ellipse.Center.Y, 0), ref minX, ref minY, ref maxX, ref maxY);
                }
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
            float x = (float)(point.X * scale + offsetX);
            float y = canvasHeight - (float)(point.Y * scale + offsetY);
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

            if (entity.Color.IsByLayer)
            {
                if (entity.Layer != null)
                {
                    color = entity.Layer.Color.ToColor();
                }
                else
                {
                    color = Color.Black; // Default color if no layer is assigned
                }
            }
            else
            {
                color = entity.Color.ToColor();
            }

            // Invert white color to black for better visibility on white background
            if (color.ToArgb() == Color.White.ToArgb())
            {
                color = Color.Black;
            }

            return color;
        }

        private (PointF StartPoint, PointF EndPoint, float Radius, float StartAngle, float SweepAngle, RectangleF Rect) GetArcSegmentFromBulge(Polyline2DVertex startVertex, Polyline2DVertex endVertex, double scale, double offsetX, double offsetY, int canvasHeight)
        {
            var startPoint = TransformPoint(new Vector3(startVertex.Position.X, startVertex.Position.Y, 0), scale, offsetX, offsetY, canvasHeight);
            var endPoint = TransformPoint(new Vector3(endVertex.Position.X, endVertex.Position.Y, 0), scale, offsetX, offsetY, canvasHeight);

            double bulge = startVertex.Bulge;
            double chordLength = Math.Sqrt(Math.Pow(endVertex.Position.X - startVertex.Position.X, 2) + Math.Pow(endVertex.Position.Y - startVertex.Position.Y, 2));
            double sagitta = Math.Abs(bulge) * chordLength / 2;
            double radius = (Math.Pow(chordLength / 2, 2) + Math.Pow(sagitta, 2)) / (2 * sagitta);

            double theta = 4 * Math.Atan(Math.Abs(bulge));
            double gamma = (Math.PI - theta) / 2;
            double phi = Math.Atan2(endVertex.Position.Y - startVertex.Position.Y, endVertex.Position.X - startVertex.Position.X);
            double centerAngle = bulge > 0 ? phi + gamma : phi - gamma;

            double centerX = startVertex.Position.X + radius * Math.Cos(centerAngle);
            double centerY = startVertex.Position.Y + radius * Math.Sin(centerAngle);

            var transformedCenter = TransformPoint(new Vector3(centerX, centerY, 0), scale, offsetX, offsetY, canvasHeight);
            float transformedRadius = (float)(radius * scale);

            double startAngle = Math.Atan2(startVertex.Position.Y - centerY, startVertex.Position.X - centerX);
            double endAngle = Math.Atan2(endVertex.Position.Y - centerY, endVertex.Position.X - centerX);
            double sweepAngle = bulge > 0 ? endAngle - startAngle : startAngle - endAngle;
            if (sweepAngle < 0) sweepAngle += 2 * Math.PI;

            startAngle = -startAngle * 180 / Math.PI;
            sweepAngle = -sweepAngle * 180 / Math.PI;

            float startAngleDeg = (float)startAngle;
            float sweepAngleDeg = (float)sweepAngle;

            if (bulge < 0)
            {
                sweepAngleDeg = -sweepAngleDeg;
            }

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

        private void DrawPolyline2D(Graphics g, Polyline2D polyline2D, double scale, double offsetX, double offsetY, int canvasHeight, Pen pen)
        {
            for (int i = 0; i < polyline2D.Vertexes.Count; i++)
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
                    PointF startPoint = TransformPoint(new Vector3(vertex.Position.X, vertex.Position.Y, 0), scale, offsetX, offsetY, canvasHeight);
                    PointF endPoint = TransformPoint(new Vector3(nextVertex.Position.X, nextVertex.Position.Y, 0), scale, offsetX, offsetY, canvasHeight);
                    g.DrawLine(pen, startPoint, endPoint);
                }
            }
        }

        private void DrawSplinePolyline(Graphics g, Polyline2D polyline2D, double scale, double offsetX, double offsetY, int canvasHeight, Pen pen)
        {
            var points = polyline2D.Vertexes.Select(v => TransformPoint(v.Position, scale, offsetX, offsetY, canvasHeight)).ToArray();
            g.DrawCurve(pen, points);
        }

        private void DrawSplinePolyline(Graphics g, Polyline3D polyline3D, double scale, double offsetX, double offsetY, int canvasHeight, Pen pen)
        {
            var points = polyline3D.Vertexes.Select(v => TransformPoint(new Vector3(v.X, v.Y, v.Z), scale, offsetX, offsetY, canvasHeight)).ToArray();
            g.DrawCurve(pen, points);
        }

        private void DrawSplinePolyline(Graphics g, List<Vector3> polylineVertexes, double scale, double offsetX, double offsetY, int canvasHeight, Pen pen)
        {
            var points = polylineVertexes.Select(v => TransformPoint(new Vector3(v.X, v.Y, v.Z), scale, offsetX, offsetY, canvasHeight)).ToArray();
            g.DrawCurve(pen, points);
        }

        private List<Vector3> GetPolylineVertexes(EntityObject polyline)
        {
            List<Vector3> vertexes = new List<Vector3>();
            if (polyline.GetType().Name == "Polyline")
            {
                var polylineType = polyline.GetType();
                var vertexesProperty = polylineType.GetProperty("Vertexes");
                if (vertexesProperty != null)
                {
                    var vertexesList = vertexesProperty.GetValue(polyline) as IEnumerable<object>;
                    if (vertexesList != null)
                    {
                        foreach (var vertex in vertexesList)
                        {
                            var positionProperty = vertex.GetType().GetProperty("Position");
                            if (positionProperty != null)
                            {
                                var position = (Vector3)positionProperty.GetValue(vertex);
                                vertexes.Add(position);
                            }
                        }
                    }
                }
            }
            return vertexes;
        }
        private List<Vector3> InterpolateSpline(List<Vector3> controlPoints, int degree, List<double> knotValues, int subdivisions)
        {
            List<Vector3> vertices = new List<Vector3>();
            double step = 1.0 / subdivisions;

            for (int i = 0; i <= subdivisions; i++)
            {
                double t = i * step * (knotValues.Last() - knotValues.First()) + knotValues.First(); // масштабируем t по диапазону узлов
                Vector3 point = DeBoor(t, degree, controlPoints, knotValues);
                vertices.Add(point);
            }

            return vertices;
        }

        private Vector3 DeBoor(double t, int degree, List<Vector3> controlPoints, List<double> knotValues)
        {
            int n = controlPoints.Count;

            // Поиск сегмента
            int s = -1;
            for (int i = degree; i < knotValues.Count - degree - 1; i++)
            {
                if (t >= knotValues[i] && t < knotValues[i + 1])
                {
                    s = i;
                    break;
                }
            }

            if (s == -1) return controlPoints[controlPoints.Count - 1];

            // Создаем массив контрольных точек
            Vector3[] dPoints = new Vector3[degree + 1];
            for (int j = 0; j <= degree; j++)
            {
                dPoints[j] = controlPoints[s - degree + j];
            }

            // Итеративно вычисляем точки
            for (int r = 1; r <= degree; r++)
            {
                for (int j = degree; j >= r; j--)
                {
                    double alpha = (t - knotValues[s - degree + j]) / (knotValues[s + 1 - r + j] - knotValues[s - degree + j]);
                    dPoints[j] = (1.0 - alpha) * dPoints[j - 1] + alpha * dPoints[j];
                }
            }

            return dPoints[degree];
        }

    }
}
