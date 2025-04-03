using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeCommunity.Extensions;
using SolidEdgeFramework;
using SolidEdgeGeometry;
using System.Windows.Forms;

namespace Application_Cyrell.Utils
{
    class FlatGenerator
    {

        public static void GenerateFlat(SolidEdgeFramework.Application application, SolidEdgeDocument document)
        {
            SolidEdgeFramework.Application seApp = null;

            SolidEdgePart.FlatPatternModel flatPatternModel = null;
            SolidEdgeGeometry.Body body = null;
            SolidEdgeGeometry.Faces faces = null;
            SolidEdgeGeometry.Face face = null;
            SolidEdgeGeometry.Edges edges = null;
            SolidEdgeGeometry.Edge edge = null;
            SolidEdgeGeometry.Vertex vertex = null;

            SolidEdgePart.Models models = null;
            SolidEdgePart.Model model = null;

            try
            {
                // Get the Solid Edge application object
                seApp = application;

                // attribute the document to the Solid Edge application object
                dynamic dynamicDoc = document;

                // Get a reference to the active document.
                dynamicDoc = (SolidEdgeDocument)seApp.ActiveDocument;

                // Get a reference to the active document.
                if (dynamicDoc is SolidEdgePart.PartDocument)
                {
                    dynamicDoc = (SolidEdgeDocument)seApp.GetActiveDocument<SolidEdgePart.PartDocument>(false);
                    Console.WriteLine("PartDocument");
                }
                else if (dynamicDoc is SolidEdgePart.SheetMetalDocument)
                {
                    dynamicDoc = (SolidEdgeDocument)seApp.GetActiveDocument<SolidEdgePart.SheetMetalDocument>(false);
                    Console.WriteLine("SheetMetalDocument");
                }
                else
                {
                    MessageBox.Show("Active document is not a PartDocument or SheetMetalDocument.", "Erreur");
                    return;
                }

                try
                {
                    models = dynamicDoc.Models;
                    model = models.Item(1);

                    Console.WriteLine($"nombre de features : {model.Features.Count.ToString()}");

                    if (model.ConvToSMs.Count == 0 && model.Features.Count == 1)
                    {
                        model.HealAndOptimizeBody(true, true);
                        body = (Body)model.Body;
                        faces = (Faces)body.Faces[FeatureTopologyQueryTypeConstants.igQueryPlane];
                        face = (Face)faces.Item(1);
                        for (int i = 2; i <= faces.Count; i++) // Parcours les faces
                        {
                            SolidEdgeGeometry.Face currentFace = (Face)faces.Item(i);

                            if (currentFace.Area > face.Area) face = currentFace;
                        }
                        Console.WriteLine($"User selected Face {face.ID} - Area: {face.Area * 1550.0031} po²");

                        edges = (SolidEdgeGeometry.Edges)face.Edges;
                        Array edgesArray = Array.CreateInstance(typeof(object), edges.Count);

                        for (int i = 1; i <= edges.Count; i++)
                        {
                            edgesArray.SetValue(edges.Item(i), i - 1); // Note the i-1 since Array is 0-based
                        }

                        model.ConvToSMs.AddEx(face, 0, edgesArray, 0, 0, 0);
                        model.ConvToSMs.Item(1).ShowDimensions = true;
                    }
                    else
                    {
                        Console.WriteLine("Il existe déjà une transformation en Synchronous Sheet Metal.", "Transformation existante");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Erreur lors de la transformation : {ex.Message}");
                    DialogResult result = MessageBox.Show(
                        "Veuillez transformer la pièce en Synchronous Sheet Metal manuellement.\n" +
                        "Quand vous aurez fini, appuyez sur OK pour continuer ou Annuler pour quitter.",
                        "Problème de transformation",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning
                    );

                    if (result == DialogResult.Cancel)
                    {
                        Console.WriteLine("🚪 Programme arrêté par l'utilisateur.");
                        return;
                    }
                }

                if (dynamicDoc.FlatPatternModels.Count == 0)
                {
                    flatPatternModel = dynamicDoc.FlatPatternModels.Add(dynamicDoc.Models.Item(1));
                }
                else
                {
                    flatPatternModel = dynamicDoc.FlatPatternModels.Item(1);
                }

                if (flatPatternModel.FlatPatterns.Count != 0)
                {
                    DialogResult result = MessageBox.Show(
                            "La pièce est déjà dépliée.\n" +
                            "Voulez vous en créer un autre?",
                            "Déplie Existant",
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Warning
                        );
                    if (result != DialogResult.OK) dynamicDoc.Save(); return;
                }

                models = dynamicDoc.Models;
                model = models.Item(1);
                body = (Body)model.Body;
                faces = (Faces)body.Faces[FeatureTopologyQueryTypeConstants.igQueryPlane];
                // Continuer le programme normalement ici
                Console.WriteLine("➡️ Continuation du programme...");

                //Automatic face and edge selections
                face = FlatPatternUtils.GetFaceFurthestFromCenter(body, faces);
                edge = FlatPatternUtils.GetEdgeAlignedWithCoordinatesSystem(face);

                Console.WriteLine($"Edge Choisi: {edge.ID}, Face Choisie: {face.ID}");
                vertex = (SolidEdgeGeometry.Vertex)edge.StartVertex;
                flatPatternModel.FlatPatterns.Add(edge, face, vertex, SolidEdgeConstants.FlattenPatternModelTypeConstants.igFlattenPatternModelTypeFlattenAnything);
                Console.WriteLine("✅ Flat pattern created successfully.");
                dynamicDoc.Save();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du dépliage du fichier {document.Name}\nErreur:  " + ex.Message);
                Console.WriteLine(ex.Message);
            }
        }
    }
}
