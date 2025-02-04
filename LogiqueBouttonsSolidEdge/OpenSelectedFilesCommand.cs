using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidEdgeCommunity;
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
                MessageBox.Show("Please select at least one file to open.");
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
                        fullPath.EndsWith(".step", StringComparison.OrdinalIgnoreCase))
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
                MessageBox.Show("Selected files processed successfully.");
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
                SolidEdgeAssembly.AssemblyDocument asmDoc = (SolidEdgeAssembly.AssemblyDocument)seApp.Documents.OpenWithTemplate(fullPath, _assemblyTemplatePath);

                // Set the document name to match the original STEP file name
                asmDoc.Name = Path.GetFileNameWithoutExtension(fullPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing {fullPath}: {ex.Message}");
            }
        }
    }
}
