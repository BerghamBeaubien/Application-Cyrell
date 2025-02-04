using Application_Cyrell.LogiqueBouttonsSolidEdge;
using SolidEdgeDraft;
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
            MessageBox.Show("Please select at least one PSM file to process.");
            return;
        }

        SolidEdgeFramework.Application seApp = null;

        try
        {
            // Get the Solid Edge application object
            seApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
            seApp.Visible = true;

            foreach (var selectedItem in _listBoxDxfFiles.SelectedItems)
            {
                string selectedFile = (string)selectedItem;
                string fullPath = System.IO.Path.Combine(_textBoxFolderPath.Text, selectedFile);

                // Open the selected PSM file
                SolidEdgeFramework.Documents documents = seApp.Documents;
                var document = documents.Open(fullPath);

                // Check if the opened document is a SheetMetalDocument
                if (document is SolidEdgePart.SheetMetalDocument sheetMetalDoc)
                {
                    // Here you can apply operations specific to SheetMetalDocument, e.g., flatten the sheet metal body
                    //sheetMetalDoc.FlatPatternModels.;

                    MessageBox.Show($"Successfully opened and processed PSM file: {fullPath}");
                }
                else
                {
                    MessageBox.Show($"File {selectedFile} is not a valid Sheet Metal Document (PSM).");
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error opening or processing PSM files in Solid Edge: " + ex.Message);
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

}
