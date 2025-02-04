using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using firstCSMacro;
using Microsoft.Office.Interop.Excel;
using SolidEdgeCommunity;
using SolidEdgeFramework;
using SolidEdgePart;
using ListBox = System.Windows.Forms.ListBox;

namespace Application_Cyrell.LogiqueBouttonsSolidEdge
{
    public class ExtracteurDimDepliCommand : IButtonManager
    {
        private readonly ListBox _listBoxDxfFiles;
        private readonly System.Windows.Forms.TextBox _textBoxFolderPath;
        private readonly PanelSettings _panelSettings;

        public ExtracteurDimDepliCommand(ListBox listBoxDxfFiles, System.Windows.Forms.TextBox textBoxFolderPath, PanelSettings pnlSettings)
        {
            _listBoxDxfFiles = listBoxDxfFiles;
            _textBoxFolderPath = textBoxFolderPath;
            _panelSettings = pnlSettings;
        }

        public void Execute()
        {
            if (_listBoxDxfFiles.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one file.");
                return;
            }

            SolidEdgeFramework.Application seApp = null;
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            bool paramGarderSEOuvert = _panelSettings.paramDim2();

            try
            {
                // Initialize Excel
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                if (excelApp == null)
                {
                    MessageBox.Show("Excel is not properly installed on your system.");
                    return;
                }

                // Initialize Solid Edge for PAR/PSM files
                bool needsSolidEdge = _listBoxDxfFiles.SelectedItems.Cast<string>()
                    .Any(file => file.EndsWith(".par", StringComparison.OrdinalIgnoreCase) ||
                                file.EndsWith(".psm", StringComparison.OrdinalIgnoreCase));

                if (needsSolidEdge)
                {
                    OleMessageFilter.Register();
                    seApp = SolidEdgeUtils.Connect(true);
                    seApp.Visible = false;
                    seApp.DisplayAlerts = false;
                }

                workbook = excelApp.Workbooks.Add();
                worksheet = (Worksheet)workbook.Worksheets[1];
                worksheet.Name = "Combined Dimensions";

                // Write Excel headers
                worksheet.Cells[1, 1] = "File Name";
                worksheet.Cells[1, 2] = "Width (inches)";
                worksheet.Cells[1, 3] = "Height (inches)";
                worksheet.Cells[1, 4] = "Thickness (inches)";

                int row = 2;

                foreach (var selectedItem in _listBoxDxfFiles.SelectedItems)
                {
                    string selectedFile = (string)selectedItem;
                    string fullPath = Path.Combine(_textBoxFolderPath.Text, selectedFile);
                    string extension = Path.GetExtension(fullPath).ToLower();

                    switch (extension)
                    {
                        case ".dxf":
                            ProcessDxfFile(fullPath, worksheet, ref row, selectedFile);
                            break;

                        case ".par":
                        case ".psm":
                            ProcessSolidEdgeFile(seApp, fullPath, worksheet, ref row, selectedFile);
                            break;

                        default:
                            MessageBox.Show($"Unsupported file type: {selectedFile}");
                            break;
                    }
                }

                // Save Excel file
                string savePath = Path.Combine(_textBoxFolderPath.Text, "Dimensions-Deplie.xlsx");
                savePath = GetUniqueFileName(savePath);

                workbook.SaveAs(savePath);
                MessageBox.Show($"Dimensions exported successfully to {savePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Cleanup Excel
                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                // Cleanup Solid Edge
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
        }

        private void ProcessDxfFile(string fullPath, Worksheet worksheet, ref int row, string selectedFile)
        {
            var (width, height) = DxfDimensionExtractor.GetDxfDimensions(fullPath);
            worksheet.Cells[row, 1] = Path.GetFileNameWithoutExtension(selectedFile);
            worksheet.Cells[row, 2] = Math.Round(width, 3);
            worksheet.Cells[row, 3] = Math.Round(height, 3);
            row++;
            worksheet.Columns.AutoFit();
        }

        private void ProcessSolidEdgeFile(SolidEdgeFramework.Application seApp, string fullPath,
            Worksheet worksheet, ref int row, string selectedFile)
        {
            var documents = seApp.Documents;
            documents.Open(fullPath);

            try
            {
                if (seApp.ActiveDocument is PartDocument partDoc)
                {
                    ProcessSolidEdgeDocument(partDoc, worksheet, ref row, selectedFile);
                }
                else if (seApp.ActiveDocument is SheetMetalDocument psmDoc)
                {
                    ProcessSolidEdgeDocument(psmDoc, worksheet, ref row, selectedFile);
                }
            }
            catch (Exception ex) { MessageBox.Show($"Error processing file {selectedFile}: {ex.Message}"); }

        }

        private void ProcessSolidEdgeDocument(dynamic document, Worksheet worksheet, ref int row, string selectedFile)
        {
            try
            {
                double valueInInchesX = 0;
                double valueInInchesY = 0;
                double valueInInchesZ = 0;
                bool paramAfficherMsgNoFPM = _panelSettings.paramDim1();

                if (document.FlatPatternModels.Count != 0)
                {
                    var variables = document.Variables;
                    var listDim = (VariableList)variables.Query(
                        pFindCriterium: "*",
                        NamedBy: SolidEdgeConstants.VariableNameBy.seVariableNameByBoth,
                        VarType: SolidEdgeConstants.VariableVarType.SeVariableVarTypeDimension
                    );

                    var flatX = listDim.Item("Flat_Pattern_Model_CutSizeX");
                    var flatY = listDim.Item("Flat_Pattern_Model_CutSizeY");

                    var listVar = (VariableList)variables.Query(
                        pFindCriterium: "*",
                        NamedBy: SolidEdgeConstants.VariableNameBy.seVariableNameByBoth,
                        VarType: SolidEdgeConstants.VariableVarType.SeVariableVarTypeVariable
                    );
                    var flatZ = listVar.Item("MaterialThickness");

                    valueInInchesX = ConvertDimensionToInches(flatX);
                    valueInInchesY = ConvertDimensionToInches(flatY);
                    valueInInchesZ = ConvertDimensionToInches(flatZ);
                }
                else
                {
                    if (paramAfficherMsgNoFPM == true) { MessageBox.Show($"La pièce {selectedFile} n'est pas déplié"); }
                }

                worksheet.Cells[row, 1] = Path.GetFileNameWithoutExtension(selectedFile);
                worksheet.Cells[row, 2] = Math.Round(valueInInchesX, 3);
                worksheet.Cells[row, 3] = Math.Round(valueInInchesY, 3);
                worksheet.Cells[row, 4] = Math.Round(valueInInchesZ, 3);
                row++;
                worksheet.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing file {selectedFile}: {ex.Message}");
            }
        }

        private double ConvertDimensionToInches(object dimension)
        {
            object value = dimension.GetType().InvokeMember("Value",
                System.Reflection.BindingFlags.GetProperty,
                null,
                dimension,
                null);
            return Convert.ToDouble(value) * 39.3701;
        }
        private string GetUniqueFileName(string path)
        {
            string directory = Path.GetDirectoryName(path);
            string fileName = Path.GetFileNameWithoutExtension(path);
            string extension = Path.GetExtension(path);
            int counter = 1;

            while (File.Exists(path))
            {
                path = Path.Combine(directory, $"{fileName} ({counter}){extension}");
                counter++;
            }

            return path;
        }
    }
}