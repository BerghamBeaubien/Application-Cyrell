using ACadSharp.IO;
using Application_Cyrell.LogiqueBouttonsSolidEdge;
using firstCSMacro;
using SolidEdgeAssembly;
using SolidEdgeCommunity.Extensions;
using SolidEdgeConstants;
using SolidEdgeDraft;
using SolidEdgeFramework;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class SaveAndCloseAllCommand : SolidEdgeCommandBase
{
    public SaveAndCloseAllCommand(TextBox textBoxFolderPath, ListBox listBoxDxfFiles)
        : base(textBoxFolderPath, listBoxDxfFiles)
    {
;
    }


    public override void Execute()
    {
        SolidEdgeFramework.Application seApp = null;

        try
        {
            SolidEdgeCommunity.OleMessageFilter.Register();

            // Connect to Solid Edge with minimized UI updates
            seApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
            seApp.Visible = false;
            seApp.DisplayAlerts = false;

            Documents documents = seApp.Documents;
            for (int i = 1; i < documents.Count +1; i++)
            {
                try
                {
                    SolidEdgeDocument document = (SolidEdgeDocument)documents.Item(i); // Solid Edge uses 1-based indexing
                    if (document != null)
                    {
                        Debug.WriteLine($"Document: {document.GetType().ToString()} - Path: {document.Path.ToString()}");
                        //MessageBox.Show($"Voici le chemin du doc : {document.FullName}, {document.Path.ToString()}");
                        // Save the document
                        //document.sav;
                        // Close the document
                        document.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error processing document: " + ex.Message);
                }
            }

            Marshal.ReleaseComObject(documents); // Release documents
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error during operation: " + ex.Message);
        }
        finally
        {
            try
            {
                if (seApp != null)
                {
                    seApp.Visible = true;
                    seApp.DisplayAlerts = true;

                    Marshal.ReleaseComObject(seApp);
                    seApp = null;

                    // Final cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                }
            }
            catch (Exception cleanupEx)
            {
                MessageBox.Show("Error during cleanup: " + cleanupEx.Message);
            }

            MessageBox.Show("All operations are complete.");
        }
    }
}