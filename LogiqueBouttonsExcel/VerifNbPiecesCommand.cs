using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Application_Cyrell.LogiqueBouttonsSolidEdge;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using Path = System.IO.Path;

namespace Application_Cyrell.LogiqueBouttonsExcel
{
    public class ValidationResultQte
    {
        public HashSet<string> ExcelTags { get; set; } = new HashSet<string>();
        public HashSet<string> DxfFiles { get; set; } = new HashSet<string>();
        public HashSet<string> StepFiles { get; set; } = new HashSet<string>();
        public int TotalCount { get; set; }
        public HashSet<string> MissingDxf { get; set; } = new HashSet<string>();
        public HashSet<string> MissingStep { get; set; } = new HashSet<string>();
        public HashSet<string> ExtraDxf { get; set; } = new HashSet<string>();
        public HashSet<string> ExtraStep { get; set; } = new HashSet<string>();
        public Dictionary<string, int> TagQuantities { get; set; } = new Dictionary<string, int>();
        public bool QcPass { get; set; }
    }

    public class DetailedComparisonFormQte : Form
    {
        public DetailedComparisonFormQte(ValidationResultQte result)
        {
            InitializeComponents(result);
        }

        private void InitializeComponents(ValidationResultQte result)
        {
            Text = "Detailed Comparison";
            Size = new Size(800, 600);

            var listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                OwnerDraw = true,
            };

            listView.Columns.Add("TAG", 200);
            listView.Columns.Add("Quantity", 70);
            listView.Columns.Add("DXF", 70);
            listView.Columns.Add("STEP", 70);
            listView.Columns.Add("Status", 200);

            // Combine all unique tags
            var allTags = new HashSet<string>();
            allTags.UnionWith(result.ExcelTags);
            allTags.UnionWith(result.DxfFiles.Select(f => Path.GetFileNameWithoutExtension(f)));
            allTags.UnionWith(result.StepFiles.Select(f => Path.GetFileNameWithoutExtension(f)));

            foreach (var tag in allTags.OrderBy(t => t))
            {
                var item = new ListViewItem(tag);

                // Quantity
                result.TagQuantities.TryGetValue(tag, out int qty);
                item.SubItems.Add(qty.ToString());

                // DXF Status
                bool hasDxf = result.DxfFiles.Contains($"{NormalizeFileName(tag)}.dxf", StringComparer.OrdinalIgnoreCase);
                item.SubItems.Add(hasDxf ? "✓" : "✗");

                // STEP Status
                bool hasStep = result.StepFiles.Any(f =>
                    f.Equals($"{NormalizeFileName(tag)}.stp", StringComparison.OrdinalIgnoreCase) ||
                    f.Equals($"{NormalizeFileName(tag)}.step", StringComparison.OrdinalIgnoreCase));
                item.SubItems.Add(hasStep ? "✓" : "✗");

                // Overall Status
                string status = "";
                if (!result.ExcelTags.Contains(tag))
                    status = "❌ Not in Excel";
                else if (!hasDxf && !hasStep)
                    status = "❌ Missing All Files";
                else if (!hasDxf)
                    status = "❌ Missing DXF";
                else if (!hasStep)
                    status = "❌ Missing STEP";
                else if (qty <= 0)
                    status = "⚠️ Invalid Quantity";
                else
                    status = "✅ OK";

                item.SubItems.Add(status);

                // Add item to the list view
                listView.Items.Add(item);
            }

            // Handle custom drawing for subitems
            listView.DrawSubItem += (s, e) =>
            {
                // Default drawing for items and subitems
                if (e.ColumnIndex == 4) // Status column
                {
                    if (e.Item.SubItems[e.ColumnIndex].Text != "✅ OK")
                        e.Graphics.FillRectangle(Brushes.MistyRose, e.Bounds);
                    else
                        e.Graphics.FillRectangle(SystemBrushes.Window, e.Bounds);
                }
                else if (e.ColumnIndex == 2 || e.ColumnIndex == 3) // DXF or STEP columns
                {
                    if (e.Item.SubItems[e.ColumnIndex].Text == "✗")
                        e.Graphics.FillRectangle(Brushes.LightSalmon, e.Bounds);
                    else
                        e.Graphics.FillRectangle(SystemBrushes.Window, e.Bounds);
                }
                else
                {
                    e.Graphics.FillRectangle(SystemBrushes.Window, e.Bounds);
                }

                // Draw the text
                e.Graphics.DrawString(
                    e.SubItem.Text,
                    listView.Font,
                    SystemBrushes.ControlText,
                    e.Bounds);
            };

            // Draw column headers
            listView.DrawColumnHeader += (s, e) =>
            {
                e.DrawBackground();
                e.Graphics.DrawString(
                    e.Header.Text,
                    listView.Font,
                    SystemBrushes.ControlText,
                    e.Bounds);
            };

            Controls.Add(listView);
        }

        private string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return fileName;

            // Remove any spaces before the extension
            int extensionIndex = fileName.LastIndexOf('.');
            if (extensionIndex > 0)
            {
                string nameWithoutExtension = fileName.Substring(0, extensionIndex).TrimEnd();
                string extension = fileName.Substring(extensionIndex);
                return nameWithoutExtension + extension;
            }

            return fileName;
        }
    }

    public class VerifNbPiecesCommand : IButtonManager
    {
        private string pathXl, pathDxf, pathStep;
        private ValidationResultQte lastValidationResultQte;
        //private const string xlpaath = "P:\\CYRAMP\\39981 FABBRICA 121 BROADWAY PO 4500023612 62  PANNES.xlsm";
        //private const string dxfpaath = "P:\\DESSINS\\FABBRICA USA\\121 BROADWAY\\BACKPAN\\1. PRODUCTION\\39981 PO 4500023612 Feb 12th\\DXF GALV 20GA";
        //private const string steppaath = "P:\\DESSINS\\FABBRICA USA\\121 BROADWAY\\BACKPAN\\1. PRODUCTION\\39981 PO 4500023612 Feb 12th\\STEP";

        public VerifNbPiecesCommand(TextBox txtBoxXl,
            TextBox txtBoxDxf, TextBox txtBoxStep)
        {
            pathXl = txtBoxXl.Text;
            pathDxf = txtBoxDxf.Text;
            pathStep = txtBoxStep.Text;
            //pathXl = xlpaath;
            //pathDxf = dxfpaath;
            //pathStep = steppaath;
        }

        public void Execute()
        {
            if (string.IsNullOrEmpty(pathXl))
            {
                MessageBox.Show("Veuillez choisir un fichier Excel à analyser");
                return;
            }

            if (!File.Exists(pathXl))
            {
                MessageBox.Show("Le fichier Excel spécifié n'existe pas.");
                return;
            }

            try
            {
                lastValidationResultQte = ValidateFiles();
                ShowValidationReport(lastValidationResultQte);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'exécution: {ex.Message}");
            }
        }

        private ValidationResultQte ValidateFiles()
        {
            var result = new ValidationResultQte();

            // Read Excel values
            var (excelTags, totalCount, tagQuantities) = ReadColumnValuesFromProjetSheet(pathXl);
            result.ExcelTags = excelTags;
            result.TotalCount = totalCount;
            result.TagQuantities = tagQuantities;

            // Get DXF files and normalize filenames
            result.DxfFiles = GetFilesByExtension(pathDxf, "dxf")
                .Select(NormalizeFileName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Get Step files and normalize filenames
            result.StepFiles = GetFilesByExtension(pathStep, "stp")
                .Concat(GetFilesByExtension(pathStep, "step"))
                .Select(NormalizeFileName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Find missing files
            result.MissingDxf = new HashSet<string>(
                result.ExcelTags.Where(tag => !result.DxfFiles.Contains($"{tag}.dxf", StringComparer.OrdinalIgnoreCase)));

            result.MissingStep = new HashSet<string>(
                result.ExcelTags.Where(tag =>
                    !result.StepFiles.Any(file =>
                        file.Equals($"{tag}.stp", StringComparison.OrdinalIgnoreCase) ||
                        file.Equals($"{tag}.step", StringComparison.OrdinalIgnoreCase))));

            // Find extra files
            result.ExtraDxf = new HashSet<string>(
                result.DxfFiles.Select(f => Path.GetFileNameWithoutExtension(f))
                    .Where(tag => !result.ExcelTags.Contains(tag, StringComparer.OrdinalIgnoreCase)));

            result.ExtraStep = new HashSet<string>(
                result.StepFiles.Select(f => Path.GetFileNameWithoutExtension(f))
                    .Where(tag => !result.ExcelTags.Contains(tag, StringComparer.OrdinalIgnoreCase)));

            // Determine QC Pass/Fail
            result.QcPass = !result.MissingDxf.Any() &&
                           !result.MissingStep.Any() &&
                           !result.ExtraDxf.Any() &&
                           !result.ExtraStep.Any() &&
                           result.TagQuantities.All(kv => kv.Value > 0);

            return result;
        }

        private void ShowValidationReport(ValidationResultQte result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== Validation Report ===");
            sb.AppendLine($"Total pieces count: {result.TotalCount}");
            sb.AppendLine($"Unique references in Excel: {result.ExcelTags.Count}");
            sb.AppendLine($"DXF files found: {result.DxfFiles.Count}");
            sb.AppendLine($"STEP files found: {result.StepFiles.Count}");
            sb.AppendLine();

            if (result.TagQuantities.Any(kv => kv.Value <= 0))
            {
                sb.AppendLine("Tags with invalid quantities:");
                foreach (var kv in result.TagQuantities.Where(kv => kv.Value <= 0))
                {
                    sb.AppendLine($"- {kv.Key}: {kv.Value}");
                }
                sb.AppendLine();
            }

            if (result.MissingDxf.Any())
            {
                sb.AppendLine("Missing DXF files:");
                foreach (var tag in result.MissingDxf.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            if (result.MissingStep.Any())
            {
                sb.AppendLine("Missing STEP files:");
                foreach (var tag in result.MissingStep.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            if (result.ExtraDxf.Any())
            {
                sb.AppendLine("Extra DXF files (not in Excel):");
                foreach (var tag in result.ExtraDxf.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            if (result.ExtraStep.Any())
            {
                sb.AppendLine("Extra STEP files (not in Excel):");
                foreach (var tag in result.ExtraStep.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            sb.AppendLine($"QC Check Result: {(result.QcPass ? "PASS" : "FAIL")}");

            var reportForm = new Form()
            {
                Text = "Validation Report",
                Size = new Size(200, 400),
                StartPosition = FormStartPosition.CenterScreen
            };

            var textBox = new TextBox()
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Text = sb.ToString()
            };

            var buttonPanel = new Panel()
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };

            var okButton = new Button()
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Left = 10,
                Top = 8
            };

            var detailsButton = new Button()
            {
                Text = "Show Details",
                Width = 90,
                Left = 90,
                Top = 8
            };

            detailsButton.Click += (s, e) =>
            {
                var detailsForm = new DetailedComparisonFormQte(lastValidationResultQte);
                detailsForm.Show();
            };

            buttonPanel.Controls.AddRange(new Control[] { okButton, detailsButton });
            reportForm.Controls.AddRange(new Control[] { textBox, buttonPanel });

            reportForm.ShowDialog();
        }

        private HashSet<string> GetFilesByExtension(string directoryPath, string extension)
        {
            var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (!Directory.Exists(directoryPath))
                return files;

            string normalizedExtension = extension.StartsWith(".", StringComparison.Ordinal)
                ? extension
                : "." + extension;

            try
            {
                var fileNames = Directory.EnumerateFiles(directoryPath, $"*.{extension}", SearchOption.AllDirectories)
                    .Select(Path.GetFileName)
                    .Select(NormalizeFileName);

                files.UnionWith(fileNames);
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException || ex is PathTooLongException)
            {
                // Handle exceptions silently but return empty set
            }

            return files;
        }

        private (HashSet<string> uniqueValues, int totalCount, Dictionary<string, int> tagQuantities)
            ReadColumnValuesFromProjetSheet(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            HashSet<string> uniqueValues = new HashSet<string>();
            Dictionary<string, int> tagQuantities = new Dictionary<string, int>();
            int totalCount = 0;

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets["PROJET"];
                    if (worksheet == null)
                    {
                        throw new Exception("La feuille 'PROJET' n'a pas été trouvée.");
                    }

                    // Determine header location
                    int startRow, column;
                    var headerD26 = worksheet.Cells[26, 4].Value?.ToString()?.Trim();
                    var headerC17 = worksheet.Cells[17, 3].Value?.ToString()?.Trim();

                    if (headerD26 == "TAG #")
                    {
                        startRow = 28;
                        column = 4;
                    }
                    else if (headerC17 == "TAG #")
                    {
                        startRow = 19;
                        column = 3;
                    }
                    else
                    {
                        throw new Exception("En-tête 'TAG #' introuvable dans D26 ou C17");
                    }

                    int endRow = worksheet.Dimension?.End.Row ?? 1000;
                    int actualLastRow = startRow;

                    // Find last row with data
                    for (int row = startRow; row <= endRow; row++)
                    {
                        var cell = worksheet.Cells[row, column].Value;
                        if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                        {
                            actualLastRow = row;
                        }
                    }

                    // Process data
                    for (int row = startRow; row <= actualLastRow; row++)
                    {
                        if (row == 501) continue;

                        var mainCell = worksheet.Cells[row, column].Value;
                        var quantityCell = worksheet.Cells[row, column + 1].Value;

                        if (mainCell != null && !string.IsNullOrWhiteSpace(mainCell.ToString()))
                        {
                            string value = mainCell.ToString().Trim();
                            if (value == "0") continue;
                            int quantity = 0;

                            // Parse quantity
                            if (quantityCell != null)
                            {
                                if (int.TryParse(quantityCell.ToString(), out int parsedQuantity))
                                {
                                    quantity = parsedQuantity;
                                }
                                else if (quantityCell.ToString().Contains("*"))
                                {
                                    string[] factors = quantityCell.ToString().Split('*');
                                    quantity = factors.All(f => int.TryParse(f, out int n))
                                        ? factors.Select(int.Parse).Aggregate((a, b) => a * b)
                                        : 0;
                                }
                            }

                            uniqueValues.Add(value);
                            tagQuantities[value] = quantity;
                            totalCount += quantity;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erreur lors de la lecture du fichier Excel: {ex.Message}");
            }

            return (uniqueValues, totalCount, tagQuantities);
        }

        private string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return fileName;

            // Remove any spaces before the extension
            int extensionIndex = fileName.LastIndexOf('.');
            if (extensionIndex > 0)
            {
                string nameWithoutExtension = fileName.Substring(0, extensionIndex).TrimEnd();
                string extension = fileName.Substring(extensionIndex);
                return nameWithoutExtension + extension;
            }

            return fileName;
        }
    }
}