using Application_Cyrell.LogiqueBouttonsSolidEdge;
using firstCSMacro;
using SolidEdgeDraft;
using SolidEdgeFramework;
using SolidEdgePart;
using System;
using System.Threading;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Environment = System.Environment;
using System.Diagnostics;
using System.Collections.Concurrent;
using ClosedXML.Excel;
using SolidEdgeCommunity;
using SolidEdgeConstants;
using SolidEdgeCommunity.Extensions;

public class CreateFlatDftCommand : SolidEdgeCommandBase
{
    private readonly string _draftTemplatePath = "P:\\Informatique\\SOLID EDGE\\TEMPLATE\\Normal.dft";
    //private readonly PanelSettings _panelSettings;
    private ConcurrentDictionary<string, HashSet<string>> _globalPartNames = new ConcurrentDictionary<string, HashSet<string>>();

    private bool paramBendTableToggle;
    private bool paramRefVars;
    private bool paramAutoScale;
    private bool paramPartsList;
    private double paramScale;
    private double paramSpaceX;
    private double paramSpaceY;
    List<bool> parametres = new List<bool>();
    List<double> valNum = new List<double>();


    public CreateFlatDftCommand(TextBox textBoxFolderPath, ListBox listBoxDxfFiles, List<bool> parametres, List<double> valNum)
        : base(textBoxFolderPath, listBoxDxfFiles)
    {
        this.parametres = parametres;
        this.valNum = valNum;
    }
    


    public override void Execute()
    {
        this.paramRefVars = parametres[0];
        this.paramBendTableToggle = parametres[1];
        this.paramAutoScale = parametres[2];
        this.paramPartsList = parametres[3];
        this.paramScale = valNum[0];
        this.paramSpaceX = valNum[1] * 0.0254;
        this.paramSpaceY = valNum[2] * 0.0254;

        SolidEdgeFramework.Application seApp = null;
        SolidEdgeFramework.Documents seDocs = null;
        SolidEdgeDraft.DraftDocument seDraftDoc = null;

        try
        {
            //Compteur pour X et Y
            double compteurX = 0.06;
            double compteurY = 0.4;
            int viewCounter = 0;
            int maxViewsPerRow = 3;

            // Get the Solid Edge application object
            OleMessageFilter.Register();
            seApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
            seApp.Visible = true;
            seDocs = seApp.Documents;
            _globalPartNames.Clear();

            // Create single draft document for all sheets
            seDraftDoc = (DraftDocument)seDocs.Add("SolidEdge.DraftDocument", _draftTemplatePath);
            seDraftDoc.Name = "Dessins dft";

            SolidEdgeDraft.Sheets sheets = seDraftDoc.Sheets;
            SolidEdgeDraft.Sheet sheet;

            if (sheets.Count == 1 && string.IsNullOrEmpty(sheets.Item(1).Name))
            {
                // Use the first sheet if it's empty
                sheet = sheets.Item(1);
            }
            else
            {
                // Add a new sheet
                sheet = sheets.Add();
            }

            sheet.Name = "";
            sheet.Activate();

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

                if (paramPartsList && paramRefVars)
                {
                    UpdateDocumentVariables(fullPath, seDocs);
                }

                // Add the model link and create the view
                SolidEdgeDraft.ModelLinks modelLinks = seDraftDoc.ModelLinks;
                SolidEdgeDraft.ModelLink modelLink = modelLinks.Add(fullPath);

                dynamic docActif = modelLink.ModelDocument;

                var scale = 0.1;

                if (paramAutoScale)
                {
                    var flatPatternModels = docActif.FlatPatternModels;
                    if (flatPatternModels.Count < 1)
                    {
                        Debug.WriteLine("No flat models. Using default scale 0.1.");
                    }
                    else
                    {
                        var flatPatternModel = flatPatternModels.Item(1);
                        var flatPattern = flatPatternModel.FlatPatterns.Item(1);

                        double minX, minY, maxX, maxY, minZ, maxZ;
                        try
                        {
                            flatPattern.Range(out minX, out minY, out minZ, out maxX, out maxY, out maxZ);

                            var xRange = maxX - minX;
                            var yRange = maxY - minY;

                            var tallDrawing = (xRange < yRange);

                            var xScale = 16 / (xRange * 39.3701);
                            var yScale = 7 / (yRange * 39.3701);

                            if (tallDrawing)
                            {
                                xScale = 7 / (xRange * 39.3701);
                                yScale = 16 / (yRange * 39.3701);
                            }

                            scale = Math.Round(Math.Min(xScale, yScale) * 0.8, 5);

                            Debug.Print($"{docActif.Name} Scale: {scale}");
                        }
                        catch (NullReferenceException ex)
                        {
                            Debug.WriteLine(ex);
                            Debug.WriteLine("Error calculating range. Using default scale 0.1.");
                        }
                    }
                }
                else
                {
                    scale = paramScale;
                }

                SolidEdgeDraft.DrawingViews dwgViews = sheet.DrawingViews;

                try
                {
                    if (viewCounter % maxViewsPerRow == 0 && viewCounter != 0)
                    {
                        // Move to the next row
                        compteurX = compteurX + paramSpaceX;
                        compteurY = 0.4;
                    }
                    Debug.WriteLine($"Adding view for {selectedFile} at ({compteurX}, {compteurY})");

                    SolidEdgeDraft.DrawingView dwgViewFlat = dwgViews.AddSheetMetalView(
                        From: modelLink,
                        Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView,
                        Scale: scale,
                        x: compteurX,
                        y: compteurY,
                        SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalFlatView
                    );

                    if (paramBendTableToggle)
                    {
                        SolidEdgeDraft.DraftBendTables bendTables = seDraftDoc.DraftBendTables;
                        SolidEdgeDraft.DrawingView bendFlatView = dwgViews.AddSheetMetalView(
                            From: modelLink,
                            Orientation: SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                            Scale: scale,
                            x: 0,
                            y: 0,
                            SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalFlatView
                        );

                        SolidEdgeDraft.DraftBendTable bendTable = bendTables.Add(
                            DrawingView: bendFlatView,
                            SavedSettings: "Normal",
                            AutoBalloon: 1,
                            CreateDraftBendTable: 1
                        );

                        SolidEdgeDraft.DraftBendTable bendTable2 = bendTables.Add(
                            DrawingView: dwgViewFlat,
                            SavedSettings: "Normal",
                            AutoBalloon: 1,
                            CreateDraftBendTable: 1
                        );

                        // Position the bend table relative to the current view position
                        bendTable.SetOrigin(
                            x: compteurX - 0.1, // Offset from the current view position
                            y: compteurY + .05
                        );

                        bendTable2.Delete();
                        bendFlatView.Delete();
                    }

                    // Ensure we're working with the current sheet's view
                    dwgViews = sheet.DrawingViews;
                    dwgViewFlat = dwgViews.Item(dwgViews.Count);

                    // Create parts list using document's PartsLists collection
                    SolidEdgeDraft.PartsLists partsLists = seDraftDoc.PartsLists;

                    // Create the parts list on the active sheet
                    SolidEdgeDraft.PartsList partsList = partsLists.Add(
                        DrawingView: dwgViewFlat,
                        SavedSettings: "ANSI",
                        AutoBalloon: 1,
                        CreatePartsList: 1
                    );

                    //Position the parts list on the current sheet
                    partsList.SetOrigin(compteurX - .1, compteurY + .1);

                    if (!paramPartsList)
                        partsList.Delete();

                    compteurY = compteurY - paramSpaceY;
                    viewCounter++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur durant la création de la table de pliage\n\n" +
                        $"La pièce {sheet.Name} n'est pas dépliée\n\n {ex.Message}",
                        "Erreur d'execution", MessageBoxButtons.OK);
                }
                seApp.StartCommand(DetailCommandConstants.DetailViewFit);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error opening or processing files in Solid Edge: " + ex.Message);
        }
        finally
        {
            if (seApp != null)
            {
                Marshal.ReleaseComObject(seApp);
                seApp = null;
            }

            MessageBox.Show("Traitement Terminé.");
        }
    }

    #region Lancer Macro DenMarForr7
    public void UpdateDocumentVariables(string fullPath, SolidEdgeFramework.Documents seDocs)
    {

        try
        {
            // Open the document
            seDocs.Open(fullPath);
            dynamic docActuel = seDocs.Application.ActiveDocument;
            // Check if it's an assembly or part
            if (seDocs.Application.ActiveDocument is PartDocument || seDocs.Application.ActiveDocument is SheetMetalDocument)
            {
                UpdatePartVariables(false);
            }
            else { UpdatePartVariables(true); }

            docActuel.Save();
            docActuel.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error updating document {fullPath}: {ex.Message}");
        }
    }

    private void UpdatePartVariables(bool assemblage)
    {
        Process appProcess = null;
        try
        {
            // Start the process with timeout protection
            appProcess = Process.Start(@"P:\Informatique\SOLID EDGE\DenMarForr7.exe");

            if (appProcess == null)
            {
                throw new Exception("Failed to start the application process.");
            }

            // Wait for the process to initialize with timeout
            bool initialized = appProcess.WaitForInputIdle(10000); // 10 second timeout
            if (!initialized)
            {
                throw new TimeoutException("Application failed to initialize within expected time.");
            }

            // Additional safeguard to ensure app is fully loaded
            System.Threading.Thread.Sleep(500);

            // Handle TAB based on assemblage flag
            if (!assemblage)
            {
                SendKeys.SendWait("{TAB}");
            }

            // Send initial ENTER
            SendKeys.SendWait("{ENTER}");

            try
            {
                // Wait for confirmation dialog with its own error handling
                WaitForConfirmationDialog();

                System.Threading.Thread.Sleep(1000);

                // Send final ENTER to complete process
                SendKeys.SendWait("{ENTER}");

                // Wait for process to exit with timeout
                bool exited = appProcess.WaitForExit(30000); // 30 second timeout
                if (!exited)
                {
                    throw new TimeoutException("Application did not exit within expected time.");
                }
            }
            catch (TimeoutException tex)
            {
                Console.WriteLine($"Timeout occurred: {tex.Message}");
                // Log the error but continue processing
                // Force process termination if it's still running
                if (!appProcess.HasExited)
                {
                    appProcess.Kill();
                }
            }
        }
        catch (Exception ex)
        {
            // Log the error but don't throw it further to allow continuation
            Console.WriteLine($"Error updating part variables: {ex.Message}");
        }
        finally
        {
            // Ensure process cleanup
            if (appProcess != null && !appProcess.HasExited)
            {
                try
                {
                    appProcess.Kill();
                }
                catch
                {
                    // Ignore errors during cleanup
                }
                finally
                {
                    appProcess.Dispose();
                }
            }
        }
    }

    private void WaitForConfirmationDialog()
    {
        // Maximum wait time (in milliseconds)
        int maxWaitTime = 60000; // 60 seconds for large assemblies
        int waitInterval = 500;  // Check every 500 ms
        int elapsedTime = 0;

        // Keep checking for the confirmation dialog
        while (elapsedTime < maxWaitTime)
        {
            try
            {
                // Check if the confirmation dialog is present
                IntPtr dialogHandle = Win32ApiHelper.FindWindow(null, "Project1");

                if (dialogHandle != IntPtr.Zero)
                {
                    // Window found, add delay for rendering
                    System.Threading.Thread.Sleep(500);
                    Console.WriteLine("Dialog is now visible");
                    return;
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors but continue waiting
                Console.WriteLine($"Error while checking for dialog: {ex.Message}");
            }

            // Wait before checking again
            System.Threading.Thread.Sleep(waitInterval);
            elapsedTime += waitInterval;
        }

        // If we reach here, timeout occurred
        throw new TimeoutException("Confirmation dialog did not appear within the expected time.");
    }

    public static class Win32ApiHelper
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    }
    #endregion
}