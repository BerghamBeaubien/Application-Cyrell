using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application_Cyrell.Utils
{
    class FlatPatternUtils
    {
        public static SolidEdgeGeometry.Face GetFaceFurthestFromCenter(SolidEdgeGeometry.Body body, SolidEdgeGeometry.Faces faces)
        {
            if (faces.Count == 0)
                return null;

            // List to store face areas and IDs
            List<(int Id, double Area, SolidEdgeGeometry.Face Face)> faceInfo = new List<(int, double, SolidEdgeGeometry.Face)>();

            // Initialize bounding box variables
            double[] minPoint = new double[3] { double.MaxValue, double.MaxValue, double.MaxValue };
            double[] maxPoint = new double[3] { double.MinValue, double.MinValue, double.MinValue };

            // Iterate through all faces to find the bounding box
            for (int i = 1; i <= faces.Count; i++)
            {
                SolidEdgeGeometry.Face currentFace = (SolidEdgeGeometry.Face)faces.Item(i);
                Array minRangePoint = Array.CreateInstance(typeof(double), 3);
                Array maxRangePoint = Array.CreateInstance(typeof(double), 3);
                currentFace.GetRange(ref minRangePoint, ref maxRangePoint);
                Array faceRange = Array.CreateInstance(typeof(double), 6);
                minRangePoint.CopyTo(faceRange, 0);
                maxRangePoint.CopyTo(faceRange, 3);

                // Store face area information
                double area = currentFace.Area;
                faceInfo.Add((i, area, currentFace));

                if (faceRange != null && faceRange.Length >= 6)
                {
                    minPoint[0] = Math.Min(minPoint[0], (double)faceRange.GetValue(0));
                    minPoint[1] = Math.Min(minPoint[1], (double)faceRange.GetValue(1));
                    minPoint[2] = Math.Min(minPoint[2], (double)faceRange.GetValue(2));

                    maxPoint[0] = Math.Max(maxPoint[0], (double)faceRange.GetValue(3));
                    maxPoint[1] = Math.Max(maxPoint[1], (double)faceRange.GetValue(4));
                    maxPoint[2] = Math.Max(maxPoint[2], (double)faceRange.GetValue(5));
                }
            }

            // Calculate the center of the bounding box
            double[] boxCenter = new double[3];
            boxCenter[0] = (minPoint[0] + maxPoint[0]) / 2.0;
            boxCenter[1] = (minPoint[1] + maxPoint[1]) / 2.0;
            boxCenter[2] = (minPoint[2] + maxPoint[2]) / 2.0;

            SolidEdgeGeometry.Face furthestFace = null;
            double maxDistance = -1;

            // Define minimum face area threshold (adjust this value as needed)
            const double MIN_FACE_AREA_THRESHOLD = 0.05; // in square meters

            // Loop through all faces to find the one furthest from center
            for (int i = 1; i <= faces.Count; i++)
            {
                SolidEdgeGeometry.Face currentFace = (SolidEdgeGeometry.Face)faces.Item(i);

                // Skip faces with area below the threshold
                if (currentFace.Area < MIN_FACE_AREA_THRESHOLD)
                    continue;

                Array minRangePoint = Array.CreateInstance(typeof(double), 3);
                Array maxRangePoint = Array.CreateInstance(typeof(double), 3);
                currentFace.GetRange(ref minRangePoint, ref maxRangePoint);
                Array faceRange = Array.CreateInstance(typeof(double), 6);
                minRangePoint.CopyTo(faceRange, 0);
                maxRangePoint.CopyTo(faceRange, 3);

                if (faceRange != null && faceRange.Length >= 6)
                {
                    // Calculate face center from its range
                    double[] faceCenter = new double[3];
                    faceCenter[0] = ((double)faceRange.GetValue(0) + (double)faceRange.GetValue(3)) / 2.0;
                    faceCenter[1] = ((double)faceRange.GetValue(1) + (double)faceRange.GetValue(4)) / 2.0;
                    faceCenter[2] = ((double)faceRange.GetValue(2) + (double)faceRange.GetValue(5)) / 2.0;

                    // Calculate distance from body center to face center
                    double distance = Math.Sqrt(
                        Math.Pow(faceCenter[0] - boxCenter[0], 2) +
                        Math.Pow(faceCenter[1] - boxCenter[1], 2) +
                        Math.Pow(faceCenter[2] - boxCenter[2], 2));

                    // Update if this face is further
                    if (distance > maxDistance)
                    {
                        maxDistance = distance;
                        furthestFace = currentFace;
                    }
                }
            }

            // Display information about the selected face
            if (furthestFace != null)
            {
                Console.WriteLine($"Selected furthest face area: {furthestFace.Area * 1550.0031:F6}sqin");
            }

            return furthestFace;
        }

        public static SolidEdgeGeometry.Edge GetEdgeAlignedWithCoordinatesSystem(SolidEdgeGeometry.Face face)
        {
            if (face == null)
                throw new System.Exception("No face selected.");

            SolidEdgeGeometry.Edges edges = (SolidEdgeGeometry.Edges)face.Edges;
            Console.WriteLine($"{edges.Count} edges found on face {face.ID}");

            if (edges.Count == 0)
                throw new System.Exception("Selected face has no edges.");

            const double METERS_TO_INCHES = 39.3701;
            SolidEdgeGeometry.Edge firstEdge = null;
            SolidEdgeGeometry.Edge selectedEdge = null;

            for (int i = 1; i <= edges.Count; i++)
            {
                SolidEdgeGeometry.Edge edge = (SolidEdgeGeometry.Edge)edges.Item(i);
                if (firstEdge == null) firstEdge = edge; // Store first edge found

                SolidEdgeGeometry.Vertex startVertex = (SolidEdgeGeometry.Vertex)edge.StartVertex;
                SolidEdgeGeometry.Vertex endVertex = (SolidEdgeGeometry.Vertex)edge.EndVertex;

                if (startVertex == null || endVertex == null)
                    continue;

                Array startPointArray = Array.CreateInstance(typeof(double), 3);
                Array endPointArray = Array.CreateInstance(typeof(double), 3);

                startVertex.GetPointData(ref startPointArray);
                endVertex.GetPointData(ref endPointArray);

                double[] startPoint = { (double)startPointArray.GetValue(0), (double)startPointArray.GetValue(1), (double)startPointArray.GetValue(2) };
                double[] endPoint = { (double)endPointArray.GetValue(0), (double)endPointArray.GetValue(1), (double)endPointArray.GetValue(2) };

                double startX = startPoint[0] * METERS_TO_INCHES;
                double startY = startPoint[1] * METERS_TO_INCHES;
                double startZ = startPoint[2] * METERS_TO_INCHES;
                double endX = endPoint[0] * METERS_TO_INCHES;
                double endY = endPoint[1] * METERS_TO_INCHES;
                double endZ = endPoint[2] * METERS_TO_INCHES;

                int sameCoordinates = 0;
                if (Math.Abs(startX - endX) < 0.001) sameCoordinates++;
                if (Math.Abs(startY - endY) < 0.001) sameCoordinates++;
                if (Math.Abs(startZ - endZ) < 0.001) sameCoordinates++;

                if (sameCoordinates >= 2)
                {
                    selectedEdge = edge;
                    break;
                }
            }

            return selectedEdge ?? firstEdge;
        }
    }
}
