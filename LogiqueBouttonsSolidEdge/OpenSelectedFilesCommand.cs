using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidEdgeCommunity;
using SolidEdgeCommunity.Extensions;
using SolidEdgeConstants;
using SolidEdgeFramework;
using SolidEdgeGeometry;
using SolidEdgePart;

namespace Application_Cyrell.LogiqueBouttonsSolidEdge
{
    public class OpenSelectedFilesCommand : IButtonManager
    {
        private readonly ListBox _listBoxDxfFiles;
        private readonly TextBox _textBoxFolderPath;

        // Paths to templates for STEP files
        private readonly string _assemblyTemplatePath = "P:\\Informatique\\SOLID EDGE\\TEMPLATE\\Normal.asm";
        private readonly string _partTemplatePath = "P:\\Informatique\\SOLID EDGE\\TEMPLATE\\Normal.par";

        public OpenSelectedFilesCommand(ListBox listBoxDxfFiles, TextBox textBoxFolderPath)
        {
            _listBoxDxfFiles = listBoxDxfFiles;
            _textBoxFolderPath = textBoxFolderPath;
        }

        public void Execute()
        {
            if (_listBoxDxfFiles.SelectedItems.Count == 0)
            {
                MessageBox.Show("Veuillez Choisir au moins un fichier à traiter.");
                return;
            }

            SolidEdgeFramework.Application seApp = null;

            try
            {
                // Connect to an existing instance of Solid Edge or start a new one
                seApp = SolidEdgeUtils.Connect(true);
                seApp.Visible = false;

                foreach (var selectedItem in _listBoxDxfFiles.SelectedItems)
                {
                    string selectedFile = (string)selectedItem;
                    string fullPath = Path.Combine(_textBoxFolderPath.Text, selectedFile);

                    // Check if the file is a STEP/STP file
                    if (fullPath.EndsWith(".stp", StringComparison.OrdinalIgnoreCase) ||
                        fullPath.EndsWith(".step", StringComparison.OrdinalIgnoreCase) ||
                        fullPath.EndsWith(".SLDASM", StringComparison.OrdinalIgnoreCase))
                    {
                        ProcessStepFile(seApp, fullPath);
                    }
                    else
                    {
                        // Open the selected file in Solid Edge (for DXF or other files)
                        seApp.Documents.Open(fullPath);
                    }
                }
                seApp.Visible = true;
                MessageBox.Show("Les fichiers ont été ouvert avec succès.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error processing files in Solid Edge: " + ex.Message);
            }
            finally
            {
                if (seApp != null)
                {
                    Marshal.ReleaseComObject(seApp);
                    seApp = null;
                }
            }
        }

        private void ProcessStepFile(SolidEdgeFramework.Application seApp, string fullPath)
        {
            try
            {
                // Ouvrir le fichier STEP avec le template d'assemblage
                SolidEdgeAssembly.AssemblyDocument asmDoc = (SolidEdgeAssembly.AssemblyDocument)seApp.Documents.OpenWithTemplate(fullPath, _assemblyTemplatePath);

                // Vérifier le nombre de composants dans l'assemblage
                SolidEdgeAssembly.Occurrences occurrences = asmDoc.Occurrences;
                int componentCount = occurrences.Count;

                if (componentCount == 1)
                {
                    // Si un seul composant, récupérer son fichier et rouvrir avec le template de pièce
                    SolidEdgeAssembly.Occurrence singleOccurrence = occurrences.Item(1);

                    // Fermer l'assemblage temporaire
                    asmDoc.Close(false);

                    // Ouvrir le fichier en tant que pièce avec le bon template
                    SolidEdgePart.PartDocument partDoc = (SolidEdgePart.PartDocument)seApp.Documents.OpenWithTemplate(fullPath, _partTemplatePath);
                    seApp.StartCommand(PartCommandConstants.PartViewFit);
                    string newPath = Path.Combine(Path.GetDirectoryName(fullPath),
                                     Path.GetFileNameWithoutExtension(fullPath) + ".par");
                    partDoc.SaveAs(newPath);
                }
                else
                {
                    // Garde le nom de l'assemblage si c'est bien un assemblage
                    seApp.StartCommand(AssemblyCommandConstants.AssemblyViewFit);
                    string newPath = Path.Combine(Path.GetDirectoryName(fullPath),
                                     Path.GetFileNameWithoutExtension(fullPath) + ".asm");
                    asmDoc.SaveAs(newPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing {fullPath}: {ex.Message}");
            }
        }

    }
}
