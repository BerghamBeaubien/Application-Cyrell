using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ACadSharp.Entities;

namespace Application_Cyrell.Utils
{
    public class CalloutPlacer
    {
        private const double MIN_SPACING = 0.02; // Minimum space between callout and entities
        private const double GRID_DIVISIONS = 10; // Number of divisions to try for placement

        public class PlacementResult
        {
            public double X { get; set; }
            public double Y { get; set; }
            public double Scale { get; set; }
        }

        public static PlacementResult CalculateOptimalPlacement(double width, double height,
            IEnumerable<Entity> entities, double textWidth, double textHeight)
        {
            // Calculate base scale based on drawing size
            double baseScale = CalculateBaseScale(width, height);

            // Create a grid of potential positions
            var potentialPositions = GeneratePotentialPositions(width, height);

            // Sort positions by distance from center to prefer central placement
            potentialPositions = potentialPositions
                .OrderBy(p => Math.Sqrt(Math.Pow(p.X - width / 2, 2) + Math.Pow(p.Y - height / 2, 2)))
                .ToList();

            foreach (var position in potentialPositions)
            {
                // Calculate callout bounds at this position
                var calloutBounds = new
                {
                    MinX = position.X,
                    MaxX = position.X + textWidth * baseScale,
                    MinY = position.Y,
                    MaxY = position.Y + textHeight * baseScale
                };

                if (IsPositionValid(calloutBounds, entities, width, height))
                {
                    return new PlacementResult
                    {
                        X = position.X,
                        Y = position.Y,
                        Scale = baseScale
                    };
                }
            }

            // If no ideal position found, try with reduced scale
            return FallbackPlacement(width, height, baseScale);
        }

        private static double CalculateBaseScale(double width, double height)
        {
            double drawingSize = Math.Max(width, height);
            if (drawingSize < 10) return 2.0;
            if (drawingSize < 20) return 3.0;
            return 4.0;
        }

        private static List<(double X, double Y)> GeneratePotentialPositions(double width, double height)
        {
            var positions = new List<(double X, double Y)>();
            double stepX = width / GRID_DIVISIONS;
            double stepY = height / GRID_DIVISIONS;

            // Generate grid positions
            for (int i = 0; i <= GRID_DIVISIONS; i++)
            {
                for (int j = 0; j <= GRID_DIVISIONS; j++)
                {
                    positions.Add((
                        X: i * stepX + MIN_SPACING,
                        Y: j * stepY + MIN_SPACING
                    ));
                }
            }

            return positions;
        }

        private static bool IsPositionValid(dynamic calloutBounds, IEnumerable<Entity> entities,
            double width, double height)
        {
            // Check if callout is within drawing bounds
            if (calloutBounds.MinX < 0 || calloutBounds.MaxX > width ||
                calloutBounds.MinY < 0 || calloutBounds.MaxY > height)
            {
                return false;
            }

            // Check for overlap with entities
            foreach (var entity in entities)
            {
                if (DoesOverlap(calloutBounds, entity))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool DoesOverlap(dynamic calloutBounds, Entity entity)
        {
            // Get entity bounds based on type
            var entityBounds = GetEntityBounds(entity);

            // Check for overlap
            return !(calloutBounds.MaxX < entityBounds.MinX ||
                    calloutBounds.MinX > entityBounds.MaxX ||
                    calloutBounds.MaxY < entityBounds.MinY ||
                    calloutBounds.MinY > entityBounds.MaxY);
        }

        private static dynamic GetEntityBounds(Entity entity)
        {
            // Add padding around entity
            const double PADDING = 0.01;

            switch (entity)
            {
                case Line line:
                    return new
                    {
                        MinX = Math.Min(line.StartPoint.X, line.EndPoint.X) - PADDING,
                        MaxX = Math.Max(line.StartPoint.X, line.EndPoint.X) + PADDING,
                        MinY = Math.Min(line.StartPoint.Y, line.EndPoint.Y) - PADDING,
                        MaxY = Math.Max(line.StartPoint.Y, line.EndPoint.Y) + PADDING
                    };
                case Arc arc:
                    // Use the existing UpdateBoundsForArc logic here
                    double minX = double.MaxValue, maxX = double.MinValue;
                    double minY = double.MaxValue, maxY = double.MinValue;
                    UpdateBoundsForArc(arc, ref minX, ref maxX, ref minY, ref maxY);
                    return new
                    {
                        MinX = minX - PADDING,
                        MaxX = maxX + PADDING,
                        MinY = minY - PADDING,
                        MaxY = maxY + PADDING
                    };
                // Add cases for other entity types as needed
                default:
                    return new
                    {
                        MinX = 0.0,
                        MaxX = 0.0,
                        MinY = 0.0,
                        MaxY = 0.0
                    };
            }
        }

        private static PlacementResult FallbackPlacement(double width, double height, double baseScale)
        {
            // If no position works, place in top-left corner with reduced scale
            return new PlacementResult
            {
                X = MIN_SPACING,
                Y = height - MIN_SPACING,
                Scale = baseScale * 0.75 // Reduce scale by 25%
            };
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
    }
}
