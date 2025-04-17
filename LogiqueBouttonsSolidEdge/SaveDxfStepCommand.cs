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
using SolidEdgeConstants;
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
                    this._outputFolderPath = form.OutputPath;
                    this.paramTagDxf = form.TagDxf;
                    this.paramChangeName = form.ChangeName;
                    this.paramFabbrica = form.Fabbrica;
                    this.paramMacroDen = form.MacroDen;
                    this.paramOnlyDxf = form.OnlyDxf;
                    this.paramOnlyStep = form.OnlyStep;
                    if (_outputFolderPath == "")
                    {
                        MessageBox.Show("Choisissez un r�p�rtoire de sortie pour continuer");
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
                MessageBox.Show("S�lection du r�pertoire annul�e");
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
                            "Voulez-vous voir les dxf g�n�res dans Solid Edge",
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

                MessageBox.Show("Traitement Termin�.");

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
                // D�finition du nom par d�faut du document
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
                        $"Le document {docName} n'est pas aplati. Impossible de g�n�rer un DXF.\nVoulez-vous en cr�er un ?",
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
                    MessageBox.Show($"Le Flat Pattern de la piece {docName} existe mais n'est pas � jour. Impossible de g�n�rer un DXF.");

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
                    seApp.StartCommand(DetailCommandConstants.DetailViewFit);

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

        #region Lancer Macro DenMarForr7
        private void UpdatePartVariables()
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

                SendKeys.SendWait("{TAB}");

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
}