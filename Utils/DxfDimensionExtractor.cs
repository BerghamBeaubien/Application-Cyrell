using System;
using System.Linq;
using System.Collections.Generic;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using System.Windows.Forms;

public class DxfDimensionExtractor
{
    public static (double width, double height) GetDxfDimensions(string filePath)
    {
        double minX = double.MaxValue;
        double maxX = double.MinValue;
        double minY = double.MaxValue;
        double maxY = double.MinValue;

        try
        {
            using (var stream = System.IO.File.OpenRead(filePath))
            {
                var loader = new DxfReader(stream);
                var document = loader.Read();
                IEnumerable<Entity> entities = GetAllEntities(document);

                foreach (var entity in entities)
                {
                    switch (entity)
                    {
                        case Arc arc:
                            UpdateBoundsForArc(arc, ref minX, ref maxX, ref minY, ref maxY);
                            break;

                        case Line line:
                            UpdateBoundsForLine(line, ref minX, ref maxX, ref minY, ref maxY);
                            break;

                        case Circle circle:
                            UpdateBoundsForCircle(circle, ref minX, ref maxX, ref minY, ref maxY);
                            break;

                        case Spline spline:
                            UpdateBoundsForSpline(spline, ref minX, ref maxX, ref minY, ref maxY);
                            break;

                        case Ellipse ellipse:
                            UpdateBoundsForEllipse(ellipse, ref minX, ref maxX, ref minY, ref maxY);
                            break;
                    }
                }
            }

            if (minX == double.MaxValue || maxX == double.MinValue ||
                minY == double.MaxValue || maxY == double.MinValue)
            {
                throw new Exception("No valid entities found in the DXF file.");
            }

            double width = maxX - minX;
            double height = maxY - minY;

            return (width, height);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing DXF file: {ex.Message}");
            return (0, 0);
        }
    }

    private static void UpdateBoundsForArc(Arc arc, ref double minX, ref double maxX, ref double minY, ref double maxY)
    {
        // Get start and end angles in radians, ensuring they're in the correct order
        double startAngle = arc.StartAngle;
        double endAngle = arc.EndAngle;

        // Normalize angles to 0-2π range
        while (startAngle < 0) startAngle += 2 * Math.PI;
        while (endAngle < 0) endAngle += 2 * Math.PI;
        while (startAngle >= 2 * Math.PI) startAngle -= 2 * Math.PI;
        while (endAngle >= 2 * Math.PI) endAngle -= 2 * Math.PI;

        // If end angle is less than start angle, add 2π to make it larger
        if (endAngle <= startAngle) endAngle += 2 * Math.PI;

        // Calculate start and end points
        double startX = arc.Center.X + arc.Radius * Math.Cos(startAngle);
        double startY = arc.Center.Y + arc.Radius * Math.Sin(startAngle);
        double endX = arc.Center.X + arc.Radius * Math.Cos(endAngle);
        double endY = arc.Center.Y + arc.Radius * Math.Sin(endAngle);

        // Initialize bounds with start and end points
        minX = Math.Min(minX, Math.Min(startX, endX));
        maxX = Math.Max(maxX, Math.Max(startX, endX));
        minY = Math.Min(minY, Math.Min(startY, endY));
        maxY = Math.Max(maxY, Math.Max(startY, endY));

        // Check if arc passes through cardinal points (0, 90, 180, 270 degrees)
        double[] cardinalAngles = { 0, Math.PI / 2, Math.PI, 3 * Math.PI / 2 };

        foreach (double angle in cardinalAngles)
        {
            // Normalize the angle for comparison
            double normalizedAngle = angle;
            while (normalizedAngle < startAngle) normalizedAngle += 2 * Math.PI;

            // Check if this cardinal angle lies within our arc
            if (normalizedAngle <= endAngle)
            {
                double x = arc.Center.X + arc.Radius * Math.Cos(angle);
                double y = arc.Center.Y + arc.Radius * Math.Sin(angle);

                minX = Math.Min(minX, x);
                maxX = Math.Max(maxX, x);
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);
            }
        }
    }

    private static void UpdateBoundsForLine(Line line, ref double minX, ref double maxX, 
        ref double minY, ref double maxY)
    {
        minX = Math.Min(minX, Math.Min(line.StartPoint.X, line.EndPoint.X));
        maxX = Math.Max(maxX, Math.Max(line.StartPoint.X, line.EndPoint.X));
        minY = Math.Min(minY, Math.Min(line.StartPoint.Y, line.EndPoint.Y));
        maxY = Math.Max(maxY, Math.Max(line.StartPoint.Y, line.EndPoint.Y));
    }

    private static void UpdateBoundsForCircle(Circle circle, ref double minX, ref double maxX, 
        ref double minY, ref double maxY)
    {
        minX = Math.Min(minX, circle.Center.X - circle.Radius);
        maxX = Math.Max(maxX, circle.Center.X + circle.Radius);
        minY = Math.Min(minY, circle.Center.Y - circle.Radius);
        maxY = Math.Max(maxY, circle.Center.Y + circle.Radius);
    }

    private static void UpdateBoundsForSpline(Spline spline, ref double minX, ref double maxX, 
        ref double minY, ref double maxY)
    {
        var splineExtents = spline.GetBoundingBox();
        minX = Math.Min(minX, splineExtents.Min.X);
        maxX = Math.Max(maxX, splineExtents.Max.X);
        minY = Math.Min(minY, splineExtents.Min.Y);
        maxY = Math.Max(maxY, splineExtents.Max.Y);
    }

    private static void UpdateBoundsForEllipse(Ellipse ellipse, ref double minX, ref double maxX, 
        ref double minY, ref double maxY)
    {
        minX = Math.Min(minX, ellipse.Center.X - ellipse.MajorAxis);
        maxX = Math.Max(maxX, ellipse.Center.X + ellipse.MajorAxis);
        minY = Math.Min(minY, ellipse.Center.Y - ellipse.MinorAxis);
        maxY = Math.Max(maxY, ellipse.Center.Y + ellipse.MinorAxis);
    }

    public static IEnumerable<Entity> GetAllEntities(CadDocument document)
    {
        var entities = new List<Entity>();
        entities.AddRange(document.Entities);

        foreach (var blockRecord in document.BlockRecords)
        {
            foreach (var entity in blockRecord.Entities)
            {
                //MessageBox.Show(entity.ToString());
                if (entity is Polyline polyline)
                {
                    entities.AddRange(polyline.Explode());
                }
                else if (entity is LwPolyline lwPolyline)
                {
                    entities.AddRange(lwPolyline.Explode());
                }
                else
                {
                    entities.Add(entity);
                }
            }
        }

        return entities;
    }

    public static int GetPartQuadrant(string filePath)
    {
        double minX = double.MaxValue;
        double maxX = double.MinValue;
        double minY = double.MaxValue;
        double maxY = double.MinValue;

        try
        {
            using (var stream = System.IO.File.OpenRead(filePath))
            {
                var loader = new DxfReader(stream);
                var document = loader.Read();
                IEnumerable<Entity> entities = GetAllEntities(document);

                foreach (var entity in entities)
                {
                    switch (entity)
                    {
                        case Arc arc:
                            UpdateBoundsForArc(arc, ref minX, ref maxX, ref minY, ref maxY);
                            break;

                        case Line line:
                            UpdateBoundsForLine(line, ref minX, ref maxX, ref minY, ref maxY);
                            break;

                        case Circle circle:
                            UpdateBoundsForCircle(circle, ref minX, ref maxX, ref minY, ref maxY);
                            break;

                        case Spline spline:
                            UpdateBoundsForSpline(spline, ref minX, ref maxX, ref minY, ref maxY);
                            break;

                        case Ellipse ellipse:
                            UpdateBoundsForEllipse(ellipse, ref minX, ref maxX, ref minY, ref maxY);
                            break;
                    }
                }
            }

            if (minX == double.MaxValue || maxX == double.MinValue ||
                minY == double.MaxValue || maxY == double.MinValue)
            {
                throw new Exception("No valid entities found in the DXF file.");
            }

            // Check if part is symmetrical: Compare distance between 0 and maxY with minY
            double positiveYDistance = maxY - 0;  // distance from origin to maxY
            double negativeYDistance = Math.Abs(minY);  // distance from origin to minY

            bool isSymmetrical = Math.Abs(positiveYDistance - negativeYDistance) < (positiveYDistance * 0.5); // Allowing a small tolerance

            // Compare total distances above and below X-axis
            double aboveXAxis = maxY;
            double belowXAxis = Math.Abs(minY);
            bool mostlyAbove = aboveXAxis > belowXAxis;

            if (isSymmetrical)
            {
                return -1;  // Indiquant la symétrie
            }

            // Pas besoin de tester maxX, il ne change pas le quadrant dans ton code
            return mostlyAbove ? 1 : 4;


        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing DXF file: {ex.Message}");
            return 0;
        }
    }
}