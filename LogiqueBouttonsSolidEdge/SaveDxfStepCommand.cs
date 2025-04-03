using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using Application_Cyrell.Utils;
using SolidEdgeCommunity;
using SolidEdgeDraft;
using SolidEdgeFramework;
using SolidEdgeGeometry;
using SolidEdgePart;
using ListBox = System.Windows.Forms.ListBox;
using TextBox = System.Windows.Forms.TextBox;

namespace Application_Cyrell.LogiqueBouttonsSolidEdge
{

    public class SaveDxfStepCommand : IButtonManager
    {
        private readonly ListBox _listBoxDxfFiles;
        private readonly TextBox _textBoxFolderPath;
        private string _outputFolderPath; // Single folder path for both DXF and STEP
        private bool paramTagDxf;
        private bool paramChangeName;
        private bool paramFabbrica;
        private bool paramMacroDen;
        private bool paramOnlyDxf;
        private bool paramOnlyStep;

        public SaveDxfStepCommand(ListBox listBoxDxfFiles, TextBox textBoxFolderPath)
        {
            _listBoxDxfFiles = listBoxDxfFiles;
            _textBoxFolderPath = textBoxFolderPath;
        }

        private bool PromptForFolder()
        {
            using (var form = new FormulaireDxfStep())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _outputFolderPath = form.OutputPath;
                    paramTagDxf = form.TagDxf;
                    paramChangeName = form.ChangeName;
                    paramFabbrica = form.Fabbrica;
                    paramMacroDen = form.MacroDen;
                    paramOnlyDxf = form.OnlyDxf;
                    paramOnlyStep = form.OnlyStep;
                    if (_outputFolderPath == "")
                    {
                        MessageBox.Show("Choisissez un répértoire de sortie pour continuer");
                        return false;
                    }
                    return true;
                }
                return false;
            }
        }

        public void Execute()
        {
            if (_listBoxDxfFiles.SelectedItems.Count == 0)
            {
                MessageBox.Show("Choisissez au moins un fichier pour continuer");
                return;
            }

            if (!PromptForFolder())
            {
                MessageBox.Show("Sélection du répertoire annulée");
                return;
            }

            SolidEdgeFramework.Application seApp = null;
            try
            {
                seApp = SolidEdgeUtils.Connect(true);
                seApp.DisplayAlerts = false;
                seApp.Visible = false;

                foreach (var selectedItem in _listBoxDxfFiles.SelectedItems)
                {
                    string selectedFile = (string)selectedItem;
                    string fullPath = Path.Combine(_textBoxFolderPath.Text, selectedFile);

                    if (fullPath.EndsWith(".par", StringComparison.OrdinalIgnoreCase) ||
                        fullPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase))
                    {
                        SaveDxfnStep(seApp, fullPath);
                    }
                    else
                    {
                        MessageBox.Show($"Le fichier {selectedFile} n'est pas un fichier PAR ou PSM");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors du traitement des fichiers dans Solid Edge: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (seApp != null)
                    {
                        DialogResult result = MessageBox.Show(
                            "Voulez-vous voir les dxf généres dans Solid Edge",
                            "Solid Edge Document Management",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question);

                        if (result == DialogResult.No)
                        {
                            seApp.Quit();
                        }
                        else
                        {
                            seApp.Visible = true;
                            seApp.DisplayAlerts = true;
                        }
                    }

                    Marshal.ReleaseComObject(seApp);
                    seApp = null;

                    // Final cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception cleanupEx)
                {
                    MessageBox.Show("Error during cleanup: " + cleanupEx.Message);
                }

                MessageBox.Show("All operations are complete.");

            }
        }

        private void SaveDxfnStep(SolidEdgeFramework.Application seApp, string fullPath)
        {
            var documents = seApp.Documents;
            documents.Open(fullPath);

            if (seApp.ActiveDocument is SolidEdgePart.PartDocument ||
                seApp.ActiveDocument is SolidEdgePart.SheetMetalDocument)
            {
                timeToShine(seApp.ActiveDocument, seApp);
            }
        }

        private void timeToShine(dynamic activeDocument, SolidEdgeFramework.Application seApp)
        {
            try
            {
                // Définition du nom par défaut du document
                string docName = Path.GetFileNameWithoutExtension(activeDocument.FullName);

                if (paramFabbrica)
                {
                    // Remplacer chaque 'B' par 'P' et supprimer tout ce qui suit "_Default_As Machined"
                    docName = docName.Replace("B", "P");
                    int index = docName.IndexOf("_Default_As Machined");
                    if (index != -1)
                    {
                        docName = docName.Substring(0, index);
                    }
                }

                if (paramChangeName)
                {
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "DXF files (*.dxf)|*.dxf|STEP files (*.stp)|*.stp";
                        saveFileDialog.Title = "Enregistrer sous";
                        saveFileDialog.FileName = docName;

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            docName = Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                        }
                    }
                }

                // Use the same folder path for both DXF and STEP files
                string activeDxfPath = Path.Combine(_outputFolderPath, $"{docName}.dxf");
                string activeStepPath = Path.Combine(_outputFolderPath, $"{docName}.stp");

                // Macro Den update if required
                if (paramMacroDen)
                {
                    seApp.Visible = true;
                    UpdatePartVariables();
                    seApp.Visible = false;
                }
                // Handling export based on parameters
                if (paramOnlyStep && !paramOnlyDxf)
                {
                    // Only STEP export - no need to check flat pattern
                    activeDocument.SaveAs(activeStepPath);
                    activeDocument.Close();
                    return;
                }

                // For DXF, check flat pattern
                SolidEdgePart.Models models = activeDocument.Models;
                SolidEdgePart.FlatPatternModels flatPatternModels = activeDocument.FlatPatternModels;

                if (flatPatternModels.Count == 0)
                {
                    DialogResult result = MessageBox.Show(
                        $"Le document {docName} n'est pas aplati. Impossible de générer un DXF.\nVoulez-vous en créer un ?",
                        "Confirmation",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );

                    if (result == DialogResult.Yes)
                    {
                        FlatGenerator.GenerateFlat(seApp, activeDocument);
                    }
                    else
                    {
                        if (!paramOnlyDxf)
                        {
                            activeDocument.SaveAs(activeStepPath);
                        }

                        activeDocument.Close();
                        return;
                    }                    
                }

                // Check if flat pattern is up to date
                bool flatPatternIsUpToDate = false;
                for (int i = 1; i <= flatPatternModels.Count; i++)
                {
                    var flatPatternModel = flatPatternModels.Item(i);
                    if (flatPatternModel.IsUpToDate)
                    {
                        flatPatternIsUpToDate = true;
                        break;
                    }
                }

                if (!flatPatternIsUpToDate)
                {
                    MessageBox.Show($"Le Flat Pattern de la piece {docName} existe mais n'est pas à jour. Impossible de générer un DXF.");

                    // If DXF can't be created, save STEP if not paramOnlyDxf
                    if (!paramOnlyDxf)
                    {
                        activeDocument.SaveAs(activeStepPath);
                    }

                    activeDocument.Close();
                    return;
                }

                // Export logic
                if (paramOnlyDxf)
                {
                    // Only DXF export
                    models.SaveAsFlatDXFEx(activeDxfPath, null, null, null, true);
                    LesAffaires(seApp, activeDxfPath);
                    activeDocument.Close();
                }
                else
                {
                    // Both DXF and STEP export
                    models.SaveAsFlatDXFEx(activeDxfPath, null, null, null, true);
                    activeDocument.SaveAs(activeStepPath);
                    activeDocument.Close();
                    LesAffaires(seApp, activeDxfPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur du traitement du fichier: {activeDocument.FullName}\nErreur: {ex.ToString()}");
            }
        }

        private void LesAffaires(dynamic seApp,string activeDxfPath)
        {
            // Open the saved DXF to add callout annotation
            var draftDoc = seApp.Documents.Open(activeDxfPath) as DraftDocument;
            if (draftDoc != null && paramTagDxf == true)
            {
                try
                {
                    // Add callout annotation or any other modifications
                    AddCalloutAnnotation(draftDoc, activeDxfPath);

                    // Delete the existing file if it exists
                    if (File.Exists(activeDxfPath))
                    {
                        File.SetAttributes(activeDxfPath, FileAttributes.Normal);
                        File.Delete(activeDxfPath);
                    }

                    // Save the document
                    draftDoc.SaveAs(activeDxfPath);
                }
                finally
                {
                    // Release the COM object
                    Marshal.ReleaseComObject(draftDoc);
                }
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

        private void UpdatePartVariables()
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

                    SendKeys.SendWait("{TAB}");

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

        public static class Win32ApiHelper
        {
            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        }
    }
}