using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidEdgeCommunity;
using SolidEdgeGeometry;
using SolidEdgePart;
using Path = System.IO.Path;

namespace Application_Cyrell.LogiqueBouttonsSolidEdge
{
    public class OpenStepFilesCommand : IButtonManager
    {
        private readonly ListBox _listBoxDxfFiles;
        private readonly TextBox _textBoxFolderPath;

        // Paths to templates
        private readonly string _assemblyTemplatePath = "P:\\Informatique\\SOLID EDGE\\TEMPLATE\\Normal.asm";
        private readonly string _partTemplatePath = "P:\\Informatique\\SOLID EDGE\\TEMPLATE\\Normal.psm";

        public OpenStepFilesCommand(ListBox listBoxDxfFiles, TextBox textBoxFolderPath)
        {
            _listBoxDxfFiles = listBoxDxfFiles;
            _textBoxFolderPath = textBoxFolderPath;
        }

        public void Execute()
        {
            if (_listBoxDxfFiles.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one STEP file to open.");
                return;
            }

            SolidEdgeFramework.Application seApp = null;
            try
            {
                // Connect to an existing instance of Solid Edge or start a new one
                seApp = SolidEdgeUtils.Connect(true);
                seApp.Visible = true;

                foreach (var selectedItem in _listBoxDxfFiles.SelectedItems)
                {
                    string selectedFile = (string)selectedItem;
                    string fullPath = Path.Combine(_textBoxFolderPath.Text, selectedFile);

                    // Only process .stp or .step files
                    if (fullPath.EndsWith(".stp", StringComparison.OrdinalIgnoreCase) ||
                        fullPath.EndsWith(".step", StringComparison.OrdinalIgnoreCase))
                    {
                        ProcessStepFile(seApp, fullPath);
                    }
                }
                seApp.Visible = true;
                MessageBox.Show("Selected STEP files processed successfully.");
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
                // Open part with Normal.par template, keeping the original filename
                SheetMetalDocument parDoc = (SheetMetalDocument)seApp.Documents.OpenWithTemplate(fullPath, _partTemplatePath);

                // Set the document name to match the original STEP file name
                parDoc.Name = Path.GetFileNameWithoutExtension(fullPath);

                MessageBox.Show("Please manually convert the part to sheet metal, then press OK to continue.");

                // Now select plane and edge for flat pattern
                var planeSelection = SelectReference(parDoc, "Please select a reference plane for the flat pattern");
                if (planeSelection != null)
                {
                    var edgeSelection = SelectReference(parDoc, "Please select an edge for the flat pattern orientation");
                    if (edgeSelection != null)
                    {
                        // Create flat pattern using the selected references
                        CreateFlatPattern(parDoc, planeSelection, edgeSelection);
                    }
                    else
                    {
                        MessageBox.Show("No edge selected. Flat pattern creation aborted.");
                    }
                }
                else
                {
                    MessageBox.Show("No plane selected. Flat pattern creation aborted.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing {fullPath}: {ex.Message}");
            }
        }

        private object SelectReference(SheetMetalDocument parDoc, string promptMessage)
        {
            try
            {
                // Activate the Part environment
                parDoc.Activate();

                // Clear any existing selection
                parDoc.SelectSet.RemoveAll();

                MessageBox.Show(promptMessage + "\nPress Enter when done.");

                // Wait for user selection and Enter key
                while (parDoc.SelectSet.Count == 0)
                {
                    System.Threading.Thread.Sleep(100);
                    Application.DoEvents();

                    if (Control.ModifierKeys == Keys.Return)
                        break;
                }

                // Get the selected object
                SolidEdgeFramework.SelectSet selectSet = parDoc.SelectSet;
                if (selectSet.Count > 0)
                {
                    var selection = selectSet.Item(1);
                    return selection;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during selection: {ex.Message}");
            }

            return null;
        }

        private void CreateFlatPattern(SheetMetalDocument parDoc, object refPlane, object refEdge)
        {
            try
            {
                // Create the flat pattern model
                var flatPatternModel = parDoc.FlatPatternModels.Add(parDoc.Models.Item(1));

                // Create the flat pattern using the selected references
                flatPatternModel.FlatPatterns.Add(
                    ReferenceEdge: refEdge,
                    ReferenceFace: refPlane,  // Actually using plane here instead of face
                    ReferenceVertex: null,
                    ModelType: SolidEdgeConstants.FlattenPatternModelTypeConstants.igFlattenPatternModelTypeFlattenAnything
                );

                MessageBox.Show("Flat pattern created successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating flat pattern: {ex.Message}\n\nStack trace: {ex.StackTrace}");
            }
        }
    }
}