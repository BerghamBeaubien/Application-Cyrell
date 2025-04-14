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

public class CreateDftCommand : SolidEdgeCommandBase
{
    private readonly string _draftTemplatePath = "P:\\Informatique\\SOLID EDGE\\TEMPLATE\\Normal.dft";
    private ConcurrentDictionary<string, HashSet<string>> _globalPartNames = new ConcurrentDictionary<string, HashSet<string>>();

    private bool paramParListPieceSolo;
    private bool paramDftIndividuelAssemblage;
    private bool paramIsoView;
    private bool paramFlatView;
    private bool paramBendTableToggle;
    private bool paramRefVars;
    private bool paramCountParts;
    private bool paramAutoScale;
    private double paramScale;

    List<bool> parametres = new List<bool>();
    List<double> valNum = new List<double>();


    public CreateDftCommand(TextBox textBoxFolderPath, ListBox listBoxDxfFiles, List<bool> parametres, List<double> valNum)
        : base(textBoxFolderPath, listBoxDxfFiles)
    {
        this.parametres = parametres;
        this.valNum = valNum;
    }

    public override void Execute()
    {
        this.paramParListPieceSolo = parametres[0];
        this.paramDftIndividuelAssemblage = parametres[1];
        this.paramIsoView = parametres[2];
        this.paramFlatView = parametres[3];
        this.paramBendTableToggle = parametres[4];
        this.paramRefVars = parametres[5];
        this.paramCountParts = parametres[6];
        this.paramAutoScale = parametres[7];
        this.paramScale = valNum[0];

        SolidEdgeFramework.Application seApp = null;
        SolidEdgeFramework.Documents seDocs = null;
        SolidEdgeDraft.DraftDocument seDraftDoc = null;

        try
        {
            // Get the Solid Edge application object
            OleMessageFilter.Register();
            seApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
            seApp.Visible = true;
            seDocs = seApp.Documents;
            _globalPartNames.Clear();

            // Create single draft document for all sheets
            seDraftDoc = (DraftDocument)seDocs.Add("SolidEdge.DraftDocument", _draftTemplatePath);
            seDraftDoc.Name = "Dessins dft";

            foreach (var selectedItem in _listBoxDxfFiles.SelectedItems)
            {
                string selectedFile = (string)selectedItem;
                string fullPath = System.IO.Path.Combine(_textBoxFolderPath.Text, selectedFile);

                // Only process .par or .psm files
                if (!fullPath.EndsWith(".par", StringComparison.OrdinalIgnoreCase) &&
                    !fullPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) &&
                    !fullPath.EndsWith(".asm", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show($"Le fichier {selectedFile} n'a pas pu etre traité en raison " +
                        "que ce n'est pas un fichier psm, par ou asm", "Erreur d'execution", MessageBoxButtons.OK);
                    continue;
                }

                if (paramRefVars)
                {
                    UpdateDocumentVariables(fullPath, seDocs);
                }

                if (paramCountParts)
                {
                    // Open the document
                    dynamic document = seDocs.Open(fullPath);

                    // Collect part names, passing the assembly name if it's an assembly
                    string sourceAssemblyName = fullPath.EndsWith(".asm", StringComparison.OrdinalIgnoreCase)
                        ? Path.GetFileName(fullPath)
                        : null;

                    string trimmedAssemblyName = sourceAssemblyName.Substring(0, sourceAssemblyName.Length - 4);

                    CollectPartNamesFromDocument(document, fullPath, trimmedAssemblyName);
                    document.Close();
                }

                // Create a new sheet for each file
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

                sheet.Name = Path.GetFileNameWithoutExtension(selectedFile);
                sheet.Activate();  // Activate the sheet before creating views

                // Add the model link and create the view
                SolidEdgeDraft.ModelLinks modelLinks = seDraftDoc.ModelLinks;
                SolidEdgeDraft.ModelLink modelLink = modelLinks.Add(fullPath);

                dynamic docActif = modelLink.ModelDocument;
                var scale = 0.1; // Default scale

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
                SolidEdgeDraft.DrawingView dwgView = dwgViews.AddPartView(
                    From: modelLink,
                    Orientation: SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                    Scale: scale,
                    x: 0.203,
                    y: 0.203,
                    ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                );

                SolidEdgeDraft.DrawingView rSideView = dwgViews.AddByFold(
                    From: dwgView,
                    foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                    x: .35,
                    y: .203
                    );

                SolidEdgeDraft.DrawingView bottomView = dwgViews.AddByFold(
                    From: dwgView,
                    foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldDown,
                    x: .203,
                    y: .09525
                    );

                if (paramIsoView)
                {
                    // Create an isometric view
                    SolidEdgeDraft.DrawingView isoView = dwgViews.AddPartView(
                        From: modelLink,
                        Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                        Scale: scale/1.5,
                        x: 0.36195,
                        y: 0.2286,
                        ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                    );
                }

                if ((paramParListPieceSolo && (fullPath.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) ||
                                 fullPath.EndsWith(".par", StringComparison.OrdinalIgnoreCase) ||
                                 fullPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase))))
                {
                    // Ensure we're working with the current sheet's view
                    dwgViews = sheet.DrawingViews;
                    dwgView = dwgViews.Item(1);

                    // Create parts list using document's PartsLists collection
                    SolidEdgeDraft.PartsLists partsLists = seDraftDoc.PartsLists;

                    // Create the parts list on the active sheet
                    SolidEdgeDraft.PartsList partsList = partsLists.Add(
                        DrawingView: dwgView,
                        SavedSettings: "ANSI",
                        AutoBalloon: 1,
                        CreatePartsList: 1
                    );

                    // Position the parts list on the current sheet
                    partsList.SetOrigin(.00635, .27305);
                }

                if ((paramFlatView && fullPath.EndsWith(".par", StringComparison.OrdinalIgnoreCase) ||
                                 fullPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)))
                {
                    try
                    {
                        SolidEdgeDraft.DrawingView dwgViewFlat = dwgViews.AddSheetMetalView(
                            From: modelLink,
                            Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView,
                            Scale: .1,
                            x: .20,
                            y: .05,
                            SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalFlatView
                        );

                        if (paramBendTableToggle)
                        {
                            SolidEdgeDraft.DraftBendTables bendTables = seDraftDoc.DraftBendTables;

                            SolidEdgeDraft.DrawingView bendFlatView = dwgViews.AddSheetMetalView(
                                From: modelLink,
                                Orientation: SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                                Scale: .1,
                                x: .5,
                                y: .5,
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

                            bendTable.SetOrigin(
                                x: .005,
                                y: 0.27559 / 2
                            );

                            bendTable2.Delete();
                            bendFlatView.Delete();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erreur durant la création de la table de pliage\n\n" +
                            $"La pièce {sheet.Name} n'est pas dépliée\n\n {ex.Message}",
                            "Erreur d'execution", MessageBoxButtons.OK);
                    }
                }

                if (paramDftIndividuelAssemblage && fullPath.EndsWith(".asm", StringComparison.OrdinalIgnoreCase))
                {
                    List<SolidEdgeAssembly.Occurrence> assemblyOccurrences = new List<SolidEdgeAssembly.Occurrence>();
                    try
                    {
                        var asmDoc = (SolidEdgeAssembly.AssemblyDocument)seApp.Documents.Open(fullPath);
                        var occurrences = asmDoc.Occurrences;

                        foreach (SolidEdgeAssembly.Occurrence occurrence in occurrences)
                        {
                            try
                            {
                                assemblyOccurrences.Add(occurrence);
                            }
                            catch (Exception)
                            {
                                continue;
                            }
                        }

                        using (Form checkboxForm = new Form())
                        {
                            checkboxForm.Text = $"Components in {Path.GetFileName(fullPath)}";
                            checkboxForm.Size = new Size(400, 300);

                            FlowLayoutPanel panel = new FlowLayoutPanel();
                            panel.Dock = DockStyle.Fill;
                            panel.AutoScroll = true;

                            Dictionary<CheckBox, SolidEdgeAssembly.Occurrence> checkBoxOccurrences =
                                new Dictionary<CheckBox, SolidEdgeAssembly.Occurrence>();

                            // Create a context menu
                            ContextMenuStrip contextMenu = new ContextMenuStrip();

                            foreach (var occurrence in assemblyOccurrences)
                            {
                                CheckBox cb = new CheckBox();
                                cb.Text = Path.GetFileName(occurrence.OccurrenceFileName);
                                cb.AutoSize = true;

                                // Create a context menu for this specific CheckBox
                                ContextMenuStrip cbContextMenu = new ContextMenuStrip();
                                ToolStripMenuItem renameMenuItem = new ToolStripMenuItem("Rename");
                                cbContextMenu.Items.Add(renameMenuItem);

                                cb.ContextMenuStrip = cbContextMenu;

                                // Rename logic
                                renameMenuItem.Click += (sender, e) =>
                                {
                                    string currentName = cb.Text;

                                    using (Form renameForm = new Form())
                                    {
                                        renameForm.Text = "Rename Occurrence";
                                        renameForm.Size = new Size(500, 100);

                                        TextBox renameTextBox = new TextBox
                                        {
                                            Text = currentName,
                                            Dock = DockStyle.Top
                                        };
                                        renameForm.Controls.Add(renameTextBox);

                                        Button okButton = new Button
                                        {
                                            Text = "OK",
                                            Dock = DockStyle.Bottom,
                                            DialogResult = DialogResult.OK // Make it act as the default button
                                        };
                                        renameForm.Controls.Add(okButton);

                                        // Handle the Enter key in the TextBox
                                        renameTextBox.KeyDown += (s, ev) =>
                                        {
                                            if (ev.KeyCode == Keys.Enter)
                                            {
                                                ev.Handled = true; // Prevent the Enter key from doing anything else
                                                ev.SuppressKeyPress = true; // Suppress the default beep sound
                                                okButton.PerformClick(); // Trigger the OK button click
                                            }
                                        };

                                        okButton.Click += (s, ev) =>
                                        {
                                            cb.Text = renameTextBox.Text; // Update CheckBox text
                                            renameForm.DialogResult = DialogResult.OK;
                                            renameForm.Close();
                                        };

                                        renameForm.ShowDialog();
                                    }
                                };


                                checkBoxOccurrences.Add(cb, occurrence);
                                panel.Controls.Add(cb);

                            }

                            Button wakhaButton = new Button();
                            wakhaButton.Text = "OK";
                            wakhaButton.DialogResult = DialogResult.OK;
                            wakhaButton.Dock = DockStyle.Bottom;

                            checkboxForm.Controls.AddRange(new Control[] { panel, wakhaButton });

                            if (checkboxForm.ShowDialog() == DialogResult.OK)
                            {
                                // Process selected occurrences
                                foreach (var kvp in checkBoxOccurrences)
                                {
                                    if (kvp.Key.Checked)
                                    {
                                        string updatedName = kvp.Key.Text; // Get updated name from CheckBox
                                        string occurrencePath = kvp.Value.OccurrenceFileName;

                                        dftIndividual(seDraftDoc, occurrencePath, updatedName, paramFlatView, paramBendTableToggle, paramIsoView);
                                    }
                                }
                            }
                        }

                        asmDoc.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error processing assembly file {selectedFile}: {ex.Message}");
                    }
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
            if (paramCountParts)
            {
                string excelOutputPath = Path.Combine(_textBoxFolderPath.Text, "Part_Names_Summary.xlsx");
                ExportPartNamesToExcel(excelOutputPath);
                MessageBox.Show($"Part names summary exported to {excelOutputPath}");
            }

            if (seApp != null)
            {
                Marshal.ReleaseComObject(seApp);
                seApp = null;
            }

            MessageBox.Show("Traitement Terminé.");
        }
    }

    private void dftIndividual(DraftDocument draftDoc, string fullPath, string sheetName, bool flat, bool bend, bool isoV)
    {
        SolidEdgeDraft.Sheets sheets = draftDoc.Sheets;
        try
        {
            // Create a new sheet if necessary
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

            // Set sheet name and activate
            sheet.Name = sheetName;
            sheet.Activate(); // Activate the sheet before creating views

            // Add the model link
            SolidEdgeDraft.ModelLinks modelLinks = draftDoc.ModelLinks;
            SolidEdgeDraft.ModelLink modelLink = modelLinks.Add(fullPath);
            
            // Create views
            SolidEdgeDraft.DrawingViews dwgViews = sheet.DrawingViews;
            SolidEdgeDraft.DrawingView dwgView = dwgViews.AddPartView(
                From: modelLink,
                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView,
                Scale: 0.1,
                x: 0.203,
                y: 0.1543,
                ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
            );

            SolidEdgeDraft.DrawingView rSideView = dwgViews.AddByFold(
                From: dwgView,
                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                x: 0.3,
                y: 0.1543
            );

            SolidEdgeDraft.DrawingView bottomView = dwgViews.AddByFold(
                From: dwgView,
                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldDown,
                x: 0.203,
                y: 0.1
            );

            if (isoV)
            {
                // Create an isometric view
                SolidEdgeDraft.DrawingView isoView = dwgViews.AddPartView(
                    From: modelLink,
                    Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                    Scale: .1,
                    x: 0.36195,
                    y: 0.2286,
                    ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                );
            }

            if (flat)
            {
                try
                {
                    SolidEdgeDraft.DrawingView dwgViewFlat = dwgViews.AddSheetMetalView(
                        From: modelLink,
                        Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView,
                        Scale: .1,
                        x: .20,
                        y: .05,
                        SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalFlatView
                    );

                    if (bend)
                    {
                        SolidEdgeDraft.DraftBendTables bendTables = draftDoc.DraftBendTables;

                        SolidEdgeDraft.DrawingView bendFlatView = dwgViews.AddSheetMetalView(
                            From: modelLink,
                            Orientation: SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                            Scale: .1,
                            x: .5,
                            y: .5,
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

                        bendTable.SetOrigin(
                            x: .005,
                            y: 0.27559 / 2
                        );

                        bendTable2.Delete();
                        bendFlatView.Delete();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur durant la création de la table de pliage\n\n" +
                        $"La pièce {sheet.Name} n'est pas dépliée\n\n {ex.Message}",
                        "Erreur d'execution", MessageBoxButtons.OK);
                }
            }

            draftDoc.Application.StartCommand(DetailCommandConstants.DetailViewFit);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error processing sheet {sheetName}: {ex.Message}");
        }
    }

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
        try
        {
            // Start the process
            Process appProcess = Process.Start(@"P:\Informatique\SOLID EDGE\DenMarForr7.exe");

            // Wait for the process to fully initialize
            if (appProcess != null)
            {
                // Wait for the main window to be ready
                appProcess.WaitForInputIdle();

                // Additional safeguard to ensure app is fully loaded
                System.Threading.Thread.Sleep(500);

                // Handle TAB based on assemblage flag
                if (!assemblage)
                {
                    SendKeys.SendWait("{TAB}");
                }

                // Send initial ENTER
                SendKeys.SendWait("{ENTER}");

                // Wait for confirmation dialog
                WaitForConfirmationDialog();

                System.Threading.Thread.Sleep(1000);

                // Send final ENTER to complete process
                SendKeys.SendWait("{ENTER}");
            }
            else
            {
                throw new Exception("Failed to start the application process.");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error updating part variables: {ex.Message}");
        }
    }

    private void WaitForConfirmationDialog()
    {
        // Maximum wait time (in milliseconds)
        int maxWaitTime = 30000; // 30 seconds
        int waitInterval = 500;  // Check every 500 ms
        int elapsedTime = 0;

        // Keep checking for the confirmation dialog
        while (elapsedTime < maxWaitTime)
        {
            // Check if the confirmation dialog is present
            // This is a placeholder - you'll need to replace with actual dialog detection
            if (IsConfirmationDialogVisible())
            {
                return; // Dialog found, proceed
            }

            // Wait a short interval
            System.Threading.Thread.Sleep(waitInterval);
            elapsedTime += waitInterval;
        }

        // If we reach here, timeout occurred
        throw new TimeoutException("Confirmation dialog did not appear within the expected time.");
    }

    private bool IsConfirmationDialogVisible()
    {
        // Use Windows API to find a window with the title "Project1"
        IntPtr dialogHandle = Win32ApiHelper.FindWindow(null, "Project1");
        return dialogHandle != IntPtr.Zero;
    }

    // You'll need to add a helper class for Windows API calls if not already present
    public static class Win32ApiHelper
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    }

    private void CollectPartNamesFromDocument(dynamic document, string fullPath, string sourceAssemblyName = null)
    {
        try
        {
            if (document is SolidEdgeAssembly.AssemblyDocument assemblyDoc)
            {
                // Dictionary to track occurrence counts within this assembly
                Dictionary<string, int> localPartCount = new Dictionary<string, int>();

                foreach (SolidEdgeAssembly.Occurrence occurrence in assemblyDoc.Occurrences)
                {
                    string originalFileName = Path.GetFileNameWithoutExtension(occurrence.OccurrenceFileName);
                    string cleanPartName = RemoveDefaultSuffix(originalFileName);

                    // Count unique parts in this assembly
                    if (localPartCount.ContainsKey(cleanPartName))
                    {
                        localPartCount[cleanPartName]++;
                    }
                    else
                    {
                        localPartCount[cleanPartName] = 1;
                    }

                    // Add each unique occurrence to the global dictionary
                    string instanceKey = $"{occurrence.OccurrenceFileName}|{sourceAssemblyName}";

                    _globalPartNames.AddOrUpdate(
                        cleanPartName,
                        new HashSet<string> { instanceKey },
                        (key, existingSet) =>
                        {
                            existingSet.Add(instanceKey);
                            return existingSet;
                        }
                    );
                }

                // Store count information in a special entry
                foreach (var partCount in localPartCount)
                {
                    string countKey = $"__COUNT__|{sourceAssemblyName}";
                    _globalPartNames.AddOrUpdate(
                        partCount.Key,
                        new HashSet<string> { $"{countKey}:{partCount.Value}" },
                        (key, existingSet) =>
                        {
                            existingSet.Add($"{countKey}:{partCount.Value}");
                            return existingSet;
                        }
                    );
                }
            }
            else if (document is SolidEdgePart.PartDocument partDoc)
            {
                string partName = Path.GetFileNameWithoutExtension(fullPath);
                string cleanPartName = RemoveDefaultSuffix(partName);

                _globalPartNames.AddOrUpdate(
                    cleanPartName,
                    new HashSet<string>(),
                    (key, existingSet) => existingSet
                );
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error collecting part names from {fullPath}: {ex.Message}");
        }
    }

    // Method to remove default suffixes
    private string RemoveDefaultSuffix(string originalName)
    {
        int defaultIndex = originalName.IndexOf("_Default_As");
        return defaultIndex != -1 ? originalName.Substring(0, defaultIndex) : originalName;
    }

    // Method to export part names to Excel using ClosedXML
    private void ExportPartNamesToExcel(string outputPath)
    {
        try
        {
            using (var workbook = new XLWorkbook())
            {
                // First worksheet for part names and counts
                var summarySheet = workbook.Worksheets.Add("Parts Summary");
                summarySheet.Cell(1, 1).Value = "Part Name";
                summarySheet.Cell(1, 2).Value = "Total Count";

                // Second worksheet for detailed part information
                var detailsSheet = workbook.Worksheets.Add("Part Details");
                detailsSheet.Cell(1, 1).Value = "Part Name";
                detailsSheet.Cell(1, 2).Value = "File Path";
                detailsSheet.Cell(1, 3).Value = "Source Assembly";
                detailsSheet.Cell(1, 4).Value = "Quantity";

                int summaryRow = 2;
                int detailsRow = 2;

                // Process each part in the dictionary
                foreach (var kvp in _globalPartNames)
                {
                    string partName = kvp.Key;
                    HashSet<string> partFiles = kvp.Value;

                    // Calculate total count from all assemblies
                    int totalCount = 0;
                    Dictionary<string, int> assemblyQuantities = new Dictionary<string, int>();

                    // First pass to collect counts and assembly quantities
                    foreach (var entry in partFiles)
                    {
                        if (entry.StartsWith("__COUNT__"))
                        {
                            string[] countParts = entry.Split('|');
                            if (countParts.Length > 1)
                            {
                                string[] countValueParts = countParts[1].Split(':');
                                if (countValueParts.Length > 1 && int.TryParse(countValueParts[1], out int count))
                                {
                                    string assemblyName = countValueParts[0];
                                    totalCount += count;

                                    if (!assemblyQuantities.ContainsKey(assemblyName))
                                    {
                                        assemblyQuantities[assemblyName] = count;
                                    }
                                    else
                                    {
                                        assemblyQuantities[assemblyName] += count;
                                    }
                                }
                            }
                        }
                    }

                    // If no count entries found, use the number of unique instances
                    if (totalCount == 0)
                    {
                        totalCount = partFiles.Count(entry => !entry.StartsWith("__COUNT__"));
                    }

                    // Add to summary sheet
                    summarySheet.Cell(summaryRow, 1).Value = partName;
                    summarySheet.Cell(summaryRow, 2).Value = totalCount;
                    summaryRow++;

                    // Add to details sheet - only for actual file entries (not count entries)
                    foreach (var fileEntry in partFiles.Where(e => !e.StartsWith("__COUNT__")))
                    {
                        var parts = fileEntry.Split('|');
                        string fileName = Path.GetFileName(parts[0]);
                        string assemblyName = parts.Length > 1 ? parts[1] : string.Empty;

                        // Find quantity for this assembly (if available)
                        int quantity = 1;
                        if (!string.IsNullOrEmpty(assemblyName) && assemblyQuantities.TryGetValue(assemblyName, out int qty))
                        {
                            quantity = qty;
                        }

                        detailsSheet.Cell(detailsRow, 1).Value = partName;
                        detailsSheet.Cell(detailsRow, 2).Value = fileName;
                        detailsSheet.Cell(detailsRow, 3).Value = assemblyName;
                        detailsSheet.Cell(detailsRow, 4).Value = quantity;
                        detailsRow++;
                    }
                }

                // Auto-fit columns for both sheets
                summarySheet.Columns().AdjustToContents();
                detailsSheet.Columns().AdjustToContents();

                // Save the workbook
                workbook.SaveAs(outputPath);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error exporting to Excel: {ex.Message}");
        }
    }

}