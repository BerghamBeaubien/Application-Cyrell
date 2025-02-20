using ACadSharp.IO;
using Application_Cyrell.LogiqueBouttonsSolidEdge;
using firstCSMacro;
using SolidEdgeAssembly;
using SolidEdgeDraft;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class ProcessDxfCommand : SolidEdgeCommandBase
{
    private readonly PanelSettings _panelSettings;

    public ProcessDxfCommand(TextBox textBoxFolderPath, ListBox listBoxDxfFiles, PanelSettings pnlSettings)
        : base(textBoxFolderPath, listBoxDxfFiles)
    {
        _panelSettings = pnlSettings;
    }


    public override void Execute()
    {
        if (_listBoxDxfFiles.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select at least one DXF file to process.");
            return;
        }

        bool paramGarderSEOuvert = _panelSettings.paramTag();

        SolidEdgeFramework.Application seApp = null;

        try
        {
            SolidEdgeCommunity.OleMessageFilter.Register();

            // Connect to Solid Edge with minimized UI updates
            seApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
            seApp.Visible = false;
            seApp.DisplayAlerts = false;

            foreach (var selectedItem in _listBoxDxfFiles.SelectedItems)
            {
                string selectedFile = (string)selectedItem;
                string fullPath = System.IO.Path.Combine(_textBoxFolderPath.Text, selectedFile);

                if (selectedFile.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase))
                {
                    SolidEdgeFramework.Documents documents = seApp.Documents;
                    documents.Open(fullPath);

                    if (seApp.ActiveDocument is DraftDocument draftDoc)
                    {
                        AddCalloutAnnotation(draftDoc, fullPath);

                        // Avoid file attribute change and deletion unless necessary
                        if (File.Exists(fullPath))
                        {
                            File.SetAttributes(fullPath, FileAttributes.Normal);
                            File.Delete(fullPath);
                        }

                        draftDoc.SaveAs(fullPath);
                        Marshal.ReleaseComObject(draftDoc); // Release draftDoc
                    }

                    Marshal.ReleaseComObject(documents); // Release documents
                }
            }
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
                    if (!paramGarderSEOuvert)
                    {
                        seApp.Quit();
                    }
                    else
                    {
                        seApp.Visible = true;
                        seApp.DisplayAlerts = true;
                    }

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

    private void AddCalloutAnnotation(DraftDocument draftDoc, String fullPath)
    {
        try
        {
            Sheet sheet = draftDoc.ActiveSheet;
            SolidEdgeFrameworkSupport.Balloons balloons = (SolidEdgeFrameworkSupport.Balloons)sheet.Balloons;
            var (width, height) = DxfDimensionExtractor.GetDxfDimensions(fullPath);

            // Determine the scale and positions based on conditions
            double scale = (width < 10 && height < 10) ? 2.0 : 4.0;

            int quadrant = DxfDimensionExtractor.GetPartQuadrant(fullPath);
            (double x1, double y1) = GetCalloutPosition(quadrant, width, height, draftDoc);

            // Adding a callout annotation
            SolidEdgeFrameworkSupport.Balloon callout = balloons.Add(
                x1: x1,
                y1: y1,
                z1: 0
            );
            callout.BalloonText = draftDoc.Name;
            callout.TextScale = scale;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error adding callout annotation: {ex.Message}");
        }
    }

    private (double x1, double y1) GetCalloutPosition(int quadrant, double width, double height, DraftDocument draftDoc)
    {
        double x1, y1;

        switch (quadrant)
        {
            case 1:
                x1 = (width < 6) ? 0.05 : 0.1;
                y1 = (height < 2.4) ? 0.005 : (height < 4 ? 0.02 : (height <= 6 ? 0.05 : 0.1));
                break;
            case 4:
                x1 = (width < 6) ? 0.05 : 0.1;
                y1 = (height < 2.4) ? -0.005 : (height < 4 ? -0.02 : (height <= 6 ? -0.05 : -0.1));
                break;
            case -1://piece symetrique
                x1 = 0.1;
                y1 = 0;
                break;
            default:
                x1 = 0.1;
                y1 = 0.1;
                break;
        }

        return (x1, y1);
    }
}