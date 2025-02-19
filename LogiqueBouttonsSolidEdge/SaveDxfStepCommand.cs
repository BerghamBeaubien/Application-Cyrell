using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
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
        private string _dxfFolderPath;
        private string _stepFolderPath;
        private bool paramTagDxf;
        private bool paramChangeName;
        private bool paramFabbrica;

        public SaveDxfStepCommand(ListBox listBoxDxfFiles, TextBox textBoxFolderPath)
        {
            _listBoxDxfFiles = listBoxDxfFiles;
            _textBoxFolderPath = textBoxFolderPath;
        }

        private bool PromptForFolders()
        {
            using (var form = new FolderSelectionForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _dxfFolderPath = form.DxfPath;
                    _stepFolderPath = form.StepPath;
                    paramTagDxf = form.TagDxf;
                    paramChangeName = form.ChangeName;
                    paramFabbrica = form.Fabbrica;
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

            if (!PromptForFolders())
            {
                MessageBox.Show("Sélection des répertoires annulée");
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
                timeToShine(seApp.ActiveDocument,seApp);
            }
        }

        private void timeToShine(dynamic activeDocument, SolidEdgeFramework.Application seApp)
        {
            try
            {
                SolidEdgePart.Models models = activeDocument.Models;
                SolidEdgePart.Model model = models.Item(1);
                SolidEdgePart.FlatPatternModels flatPatternModels = activeDocument.FlatPatternModels;

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
                        saveFileDialog.FileName = docName; // Utilise le nom modifié si `paramFabbrica` est vrai

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            docName = Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                        }
                    }
                }


                string activeDxfPath = Path.Combine(_dxfFolderPath, $"{docName}.dxf");
                string activeStepPath = Path.Combine(_stepFolderPath, $"{docName}.stp");

                if (flatPatternModels.Count == 0)
                {
                    MessageBox.Show($"Le document {docName} n'est pas aplati. Impossible de générer un DXF.");
                    activeDocument.SaveAs(activeStepPath);
                    activeDocument.Close();
                    return;
                }

                if (flatPatternModels.Count > 0)
                {
                    SolidEdgePart.FlatPatternModel flatPatternModel = null;
                    bool flatPatternIsUpToDate = false;

                    for (int i = 1; i <= flatPatternModels.Count; i++)
                    {
                        flatPatternModel = flatPatternModels.Item(i);
                        if (flatPatternModel.IsUpToDate)
                        {
                            flatPatternIsUpToDate = true;
                            break;
                        }
                    }

                    if (!flatPatternIsUpToDate)
                    {
                        MessageBox.Show($"Le Flat Pattern de la piece {docName} existe mais n'est pas à jour. Impossible de générer un DXF.");
                        return;
                    }

                    // Save as DXF
                    models.SaveAsFlatDXFEx(activeDxfPath, null, null, null, true);
                    activeDocument.SaveAs(activeStepPath);
                    activeDocument.Close();

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
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur du traitement du fichier: {activeDocument.FullName}\nErreur: {ex.ToString()}");
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
                callout.BalloonText = "%{File Name (no extension)}";
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

    public class FolderSelectionForm : Form
    {
        private TextBox txtDxfPath;
        private TextBox txtStepPath;
        private CheckBox chkTagDxf;
        private CheckBox chkChangeName;
        private CheckBox chkFabbrica;
        private Button btnBrowseDxf;
        private Button btnBrowseStep;
        private Button btnContinue;
        private Button btnCancel;

        public string DxfPath => txtDxfPath.Text;
        public string StepPath => txtStepPath.Text;
        public bool TagDxf => chkTagDxf.Checked;
        public bool ChangeName => chkChangeName.Checked;
        public bool Fabbrica => chkFabbrica.Checked;

        public FolderSelectionForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Sélection des répertoires";
            this.Size = new Size(500, 200);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;

            // DXF Path Controls
            Label lblDxf = new Label
            {
                Text = "Répertoire DXF:",
                Location = new System.Drawing.Point(10, 15),
                AutoSize = true
            };

            txtDxfPath = new TextBox
            {
                Location = new System.Drawing.Point(10, 35),
                Width = 380,
                ReadOnly = true
            };

            btnBrowseDxf = new Button
            {
                Text = "...",
                Location = new System.Drawing.Point(400, 34),
                Width = 30
            };
            btnBrowseDxf.Click += (s, e) => BrowseFolder(txtDxfPath);

            // STEP Path Controls
            Label lblStep = new Label
            {
                Text = "Répertoire STEP:",
                Location = new System.Drawing.Point(10, 65),
                AutoSize = true
            };

            txtStepPath = new TextBox
            {
                Location = new System.Drawing.Point(10, 85),
                Width = 380,
                ReadOnly = true
            };

            btnBrowseStep = new Button
            {
                Text = "...",
                Location = new System.Drawing.Point(400, 84),
                Width = 30
            };
            btnBrowseStep.Click += (s, e) => BrowseFolder(txtStepPath);

            // Checkbox
            chkTagDxf = new CheckBox
            {
                Text = "Tag DXF",
                Location = new System.Drawing.Point(10, 115),
                AutoSize = true
            };

            chkChangeName = new CheckBox
            {
                Text = "Changer le nom",
                Location = new System.Drawing.Point(80, 115),
                AutoSize = true
            };

            chkFabbrica = new CheckBox
            {
                Text = "Fabbrica",
                Location = new System.Drawing.Point(180, 115),
                AutoSize = true
            };

            // Buttons
            btnContinue = new Button
            {
                Text = "Continuer",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(280, 115),
                Width = 80
            };

            btnCancel = new Button
            {
                Text = "Annuler",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(370, 115),
                Width = 80
            };

            this.Controls.AddRange(new Control[] {
            lblDxf, txtDxfPath, btnBrowseDxf,
            lblStep, txtStepPath, btnBrowseStep,
            chkTagDxf, btnContinue, btnCancel,chkChangeName,chkFabbrica
        });

            this.AcceptButton = btnContinue;
            this.CancelButton = btnCancel;
        }

        private void BrowseFolder(TextBox textBox)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.CheckFileExists = false;
                dialog.CheckPathExists = true;
                dialog.ValidateNames = false;
                dialog.FileName = "Folder Selection."; // Placeholder text

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = Path.GetDirectoryName(dialog.FileName);
                    textBox.Text = selectedPath;
                }
            }
        }
    }
}