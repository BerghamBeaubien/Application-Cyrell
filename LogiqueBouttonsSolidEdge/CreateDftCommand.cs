using Application_Cyrell.LogiqueBouttonsSolidEdge;
using firstCSMacro;
using SolidEdgeDraft;
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Environment = System.Environment;

public class CreateDftCommand : SolidEdgeCommandBase
{
    private readonly string _draftTemplatePath = "P:\\Informatique\\SOLID EDGE\\TEMPLATE\\Normal.dft";
    private readonly PanelSettings _panelSettings;

    public CreateDftCommand(TextBox textBoxFolderPath, ListBox listBoxDxfFiles, PanelSettings pnlSettings)
        : base(textBoxFolderPath, listBoxDxfFiles)
    {
        _panelSettings = pnlSettings;
    }

    public override void Execute()
    {
        if (_listBoxDxfFiles.SelectedItems.Count == 0)
        {
            MessageBox.Show("Please select at least one PSM or PAR file to process.");
            return;
        }

        bool paramParListPieceSolo = _panelSettings.paramDft1();
        bool paramDftIndividuelAssemblage = _panelSettings.paramDft2();
        bool paramIsoView = _panelSettings.paramDft3();
        bool paramFlatView = _panelSettings.paramDft4();
        bool paramBendTableToggle = _panelSettings.paramDft5();

        SolidEdgeFramework.Application seApp = null;
        SolidEdgeFramework.Documents seDocs = null;
        SolidEdgeDraft.DraftDocument seDraftDoc = null;

        try
        {
            // Get the Solid Edge application object
            seApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
            seApp.Visible = true;
            seDocs = seApp.Documents;

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
                    MessageBox.Show($"Le fichier {selectedFile} n'a pas pu etre traite en raison " +
                        "que ce n'est pas un fichier psm, par ou asm", "Erreur d'execution", MessageBoxButtons.OK);
                    continue;
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
                SolidEdgeDraft.DrawingViews dwgViews = sheet.DrawingViews;
                SolidEdgeDraft.DrawingView dwgView = dwgViews.AddPartView(
                    From: modelLink,
                    Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView,
                    Scale: .1,
                    x: 0.203,
                    y: 0.1543,
                    ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                );

                SolidEdgeDraft.DrawingView rSideView = dwgViews.AddByFold(
                    From: dwgView,
                    foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                    x: .3,
                    y: .1543
                    );

                SolidEdgeDraft.DrawingView bottomView = dwgViews.AddByFold(
                    From: dwgView,
                    foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldDown,
                    x: .203,
                    y: .10
                    );

                if (paramIsoView)
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

                // Add dimensions to the views
                try
                {
                    //// Create dimensions collections for each view
                    //SolidEdgeDraft.Dimensions topDimensions = dwgView.Dimensions;
                    //SolidEdgeDraft.Dimensions rightDimensions = rSideView.Dimensions;
                    //SolidEdgeDraft.Dimensions bottomDimensions = bottomView.Dimensions;
                    //SolidEdgeDraft.dim

                    //// Use SmartDimension to automatically place critical dimensions
                    //// For top view
                    //SolidEdgeDraft.SmartDimension topSmartDim = sheet.SmartDimensions;
                    //topSmartDim.Create(
                    //    View: dwgView,
                    //    Type: SolidEdgeDraft.SmartDimensionConstants.igSmartDimensionHorizontal,
                    //    AddBends: false,
                    //    AddAngles: false,
                    //    StaggerOffset: 0.01);

                    //// For right side view
                    //SolidEdgeDraft.SmartDimension rightSmartDim = sheet.SmartDimensions;
                    //rightSmartDim.Create(
                    //    View: rSideView,
                    //    Type: SolidEdgeDraft.SmartDimensionConstants.igSmartDimensionVertical,
                    //    AddBends: false,
                    //    AddAngles: false,
                    //    StaggerOffset: 0.01);

                    //// For bottom view
                    //SolidEdgeDraft.SmartDimension bottomSmartDim = sheet.SmartDimensions;
                    //bottomSmartDim.Create(
                    //    View: bottomView,
                    //    Type: SolidEdgeDraft.SmartDimensionConstants.igSmartDimensionAll,
                    //    AddBends: false,
                    //    AddAngles: false,
                    //    StaggerOffset: 0.01);

                    //// Process the dimensions to ensure they're properly placed
                    //sheet.ProcessAll();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error adding dimensions to views for {sheet.Name}: {ex.Message}",
                        "Dimension Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

                        if(paramBendTableToggle)
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
                                y: 0.27559/2
                            );

                            bendTable2.Delete();
                            bendFlatView.Delete();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erreur durant la création de la table de pliage\n\n" +
                            $"La pièce {sheet.Name} n'est pas dépliée\n\n {ex.Message}",
                            "Erreur d'execution",MessageBoxButtons.OK);
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

                                        dftIndividual(seDraftDoc, occurrencePath, updatedName, paramFlatView, paramBendTableToggle);
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
        }
    }

    private void dftIndividual(DraftDocument draftDoc, string fullPath, string sheetName, bool flat, bool bend)
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

            if (flat)
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
                    SolidEdgeDraft.DraftBendTable bendTable = bendTables.Add(
                        DrawingView: dwgViewFlat,
                        SavedSettings: "Normal",
                        AutoBalloon: 1,
                        CreateDraftBendTable: 1
                    );
                    bendTable.SetOrigin(
                        x: .005,
                        y: .1
                    );
                }
            }

            
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error processing sheet {sheetName}: {ex.Message}");
        }
    }
}