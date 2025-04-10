using Application_Cyrell.LogiqueBouttonsSolidEdge;
using Application_Cyrell.Utils;
using SolidEdgeCommunity.Extensions;
using SolidEdgeDraft;
using SolidEdgeFramework;
using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class ProcessPsmCommand : SolidEdgeCommandBase
{
    public ProcessPsmCommand(TextBox textBoxFolderPath, ListBox listBoxDxfFiles)
        : base(textBoxFolderPath, listBoxDxfFiles) { }

    public override void Execute()
    {
        if (_listBoxDxfFiles.SelectedItems.Count == 0)
        {
            MessageBox.Show("Veuillez selectionner au moins un fichier à traiter.");
            return;
        }

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
        bool autoMod = true;
        bool closeDoc = false;

        using (var form = new FlatPatternPromptForm())
        {
            if (form.ShowDialog() == DialogResult.OK)
            {
                if (form.IsAutomatic) autoMod = false;
                if (form.CloseDocument) closeDoc = true; 
            }
            else return;
        }

        try
        {
            // Register with OLE to handle concurrency issues on the current thread.
            SolidEdgeCommunity.OleMessageFilter.Register();

            // Get the Solid Edge application object
            seApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
            seApp.Visible = true;

            foreach (var selectedItem in _listBoxDxfFiles.SelectedItems)
            {
                string selectedFile = (string)selectedItem;
                string fullPath = System.IO.Path.Combine(_textBoxFolderPath.Text, selectedFile);

                // Only process .par or .psm files
                if (!fullPath.EndsWith(".par", StringComparison.OrdinalIgnoreCase) &&
                    !fullPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show($"Le fichier {selectedFile} n'a pas pu etre traité en raison " +
                        "que ce n'est pas un fichier psm ou par", "Erreur d'execution", MessageBoxButtons.OK);
                    continue;
                }

                // Open the selected PSM file
                SolidEdgeFramework.Documents documents = seApp.Documents;
                dynamic dynamicDoc = documents.Open(fullPath);

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
                    dynamicDoc.Close();
                    continue;
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
                        dynamicDoc.Close();
                        continue;
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
                    if (result != DialogResult.OK)
                    {
                        dynamicDoc.Save();
                        continue;
                    }
                }

                models = dynamicDoc.Models;
                model = models.Item(1);
                body = (Body)model.Body;
                faces = (Faces)body.Faces[FeatureTopologyQueryTypeConstants.igQueryPlane];
                // Continuer le programme normalement ici
                Console.WriteLine("➡️ Continuation du programme...");

                if (autoMod)
                {
                    seApp.StartCommand((SolidEdgeCommandConstants)45066);
                    seApp.StartCommand((SolidEdgeCommandConstants)45070);
                    seApp.StartCommand((SolidEdgeCommandConstants)45063);
                    MessageBox.Show(
                            "Veuillez Choisir une Face et une Arête.\n" +
                            "Appuyez sur OK quand vous aurez terminer",
                            "Choisir Face et Arête",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                    dynamicDoc.Save();
                    if (closeDoc)
                    {
                        dynamicDoc.Close();
                    }
                    continue;
                }
                else
                {

                    face = FlatPatternUtils.GetFaceFurthestFromCenter(body, faces);

                    //// Manual face selection remplacer XXXX par l'ID de la face (Pour Débogage)
                    //face = (Face)faces.Item(1);
                    //for (int i =1; i <= faces.Count; i++)
                    //{
                    //    SolidEdgeGeometry.Face currentFace = (Face)faces.Item(i);
                    //    if (currentFace.ID == XXXX) face = currentFace;
                    //}
                    //Console.WriteLine($"User selected Face {face.ID} - Area: {face.Area * 1550.0031} mm²");

                    //Automatic edge selections
                    edge = FlatPatternUtils.GetEdgeAlignedWithCoordinatesSystem(face);
                }

                Console.WriteLine($"Edge Choisi: {edge.ID}, Face Choisie: {face.ID}");
                vertex = (SolidEdgeGeometry.Vertex)edge.StartVertex;
                flatPatternModel.FlatPatterns.Add(edge, face, vertex, SolidEdgeConstants.FlattenPatternModelTypeConstants.igFlattenPatternModelTypeFlattenAnything);
                Console.WriteLine("✅ Flat pattern created successfully.");
                dynamicDoc.Save();
                if (closeDoc)
                {
                    dynamicDoc.Close();
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error opening or processing PSM files in Solid Edge: " + ex.Message);
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (seApp != null)
            {
                Marshal.ReleaseComObject(seApp);
                SolidEdgeCommunity.OleMessageFilter.Unregister();
                seApp = null;
            }
            MessageBox.Show("Traitement Terminé.");
        }
    }

}
