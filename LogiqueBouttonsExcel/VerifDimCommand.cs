using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
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
    public class ValidationResultDimension
    {
        public HashSet<string> ExcelTags { get; set; } = new HashSet<string>();
        public HashSet<string> DxfFiles { get; set; } = new HashSet<string>();
        public int TotalCount { get; set; }
        public HashSet<string> MissingDxf { get; set; } = new HashSet<string>();
        public HashSet<string> ExtraDxf { get; set; } = new HashSet<string>();
        public Dictionary<string, int> TagQuantities { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, (double Width, double Height, int Row)> TagDimensions { get; set; } = new Dictionary<string, (double Width, double Height, int Row)>();
        public HashSet<string> DimensionMismatches { get; set; } = new HashSet<string>();  // Added to track mismatches
        public bool QcPass { get; set; }
        public Dictionary<string, bool> SwappedDimensions { get; set; } = new Dictionary<string, bool>();
    }

    public class DetailedComparisonFormDimension : Form
    {
        private readonly string _pathDxf;
        public DetailedComparisonFormDimension(ValidationResultDimension result, string pathDxf)
        {
            _pathDxf = pathDxf;
            InitializeComponents(result);
        }

        private ListView listView;
        private MenuStrip menuStrip;

        private void InitializeComponents(ValidationResultDimension result)
        {
            Text = "Detailed Comparison";
            Size = new Size(1000, 600);

            listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                OwnerDraw = true
            };

            menuStrip = new MenuStrip();

            var containerPanel = new Panel
            {
                Dock = DockStyle.Fill
            };
            menuStrip.Dock = DockStyle.Top;
            Controls.Add(menuStrip);

            listView.Dock = DockStyle.Fill;
            containerPanel.Controls.Add(listView);
            var sortMenu = new ToolStripMenuItem("Sort");

            var sortByRowAsc = new ToolStripMenuItem("Sort by Row (Ascending)");
            sortByRowAsc.Click += (s, e) =>
            {
                listView.ListViewItemSorter = new ListViewItemComparer(0, true);
                listView.Sort();
            };

            var sortByTagAsc = new ToolStripMenuItem("Sort by TAG (A-Z)");
            sortByTagAsc.Click += (s, e) =>
            {
                listView.ListViewItemSorter = new ListViewItemComparer(1, true);
                listView.Sort();
            };

            sortMenu.DropDownItems.Add(sortByRowAsc);
            sortMenu.DropDownItems.Add(sortByTagAsc);
            menuStrip.Items.Add(sortMenu);

            Controls.Add(menuStrip);
            listView.Dock = DockStyle.Fill;

            listView.Columns.Add("Row", 50);
            listView.Columns.Add("TAG", 170);
            listView.Columns.Add("Qty", 40);
            listView.Columns.Add("DXF", 40);
            listView.Columns.Add("Width Match", 40);
            listView.Columns.Add("Height Match", 40);
            listView.Columns.Add("Mismatched Width", 140); // Only show if width doesn't match
            listView.Columns.Add("Mismatched Height", 140); // Only show if height doesn't match
            listView.Columns.Add("Status", 200);

            // Combine all unique tags
            var allTags = new HashSet<string>();
            allTags.UnionWith(result.ExcelTags);
            allTags.UnionWith(result.DxfFiles.Select(f => Path.GetFileNameWithoutExtension(f)));

            foreach (var tag in allTags.OrderBy(t => t))
            {
                var item = new ListViewItem("");
                // Add row number as first subitem
                int rowNumber = result.TagDimensions.TryGetValue(tag, out var dim) ? dim.Row : 0;
                item.Text = rowNumber.ToString(); // First column is Row

                item.SubItems.Add(tag);

                // Quantity
                result.TagQuantities.TryGetValue(tag, out int qty);
                item.SubItems.Add(qty.ToString());

                // DXF Status
                bool hasDxf = result.DxfFiles.Contains($"{NormalizeFileName(tag)}.dxf", StringComparer.OrdinalIgnoreCase);
                item.SubItems.Add(hasDxf ? "✓" : "✗");

                // Excel Dimensions
                double excelWidth = 0;
                double excelHeight = 0;
                if (result.TagDimensions.TryGetValue(tag, out var excelDimensions))
                {
                    excelWidth = excelDimensions.Width;
                    excelHeight = excelDimensions.Height;
                }

                // DXF Dimensions
                double dxfWidth = 0;
                double dxfHeight = 0;
                if (hasDxf)
                {
                    var dxfDimensions = DxfDimensionExtractor.GetDxfDimensions(Path.Combine(_pathDxf, $"{NormalizeFileName(tag)}.dxf"));
                    dxfWidth = dxfDimensions.width;
                    dxfHeight = dxfDimensions.height;
                }

                // Comparaison Moins Stricte
                bool widthMatch = Math.Truncate(excelWidth) == Math.Truncate(dxfWidth);
                bool heightMatch = Math.Truncate(excelHeight) == Math.Truncate(dxfHeight);
                bool swappedMatch = false;

                if (!widthMatch || !heightMatch)
                {
                    bool swappedWidthMatch = Math.Truncate(excelWidth) == Math.Truncate(dxfHeight);
                    bool swappedHeightMatch = Math.Truncate(excelHeight) == Math.Truncate(dxfWidth);
                    swappedMatch = swappedWidthMatch && swappedHeightMatch;

                    if (swappedMatch)
                    {
                        widthMatch = true;
                        heightMatch = true;
                    }
                }

                // Width Match Column
                item.SubItems.Add(widthMatch ? "✓" : "✗");

                // Height Match Column
                item.SubItems.Add(heightMatch ? "✓" : "✗");

                // Mismatched Width Column (only show if width doesn't match)
                item.SubItems.Add(widthMatch ? "" : $"XL {excelWidth:F3} | DXF {dxfWidth:F3}");

                // Mismatched Height Column (only show if height doesn't match)
                item.SubItems.Add(heightMatch ? "" : $"XL {excelHeight:F3} | DXF {dxfHeight:F3}");

                string status = "";
                bool isSwapped = result.SwappedDimensions.TryGetValue(tag, out bool swapped) && swapped;

                // Overall Status
                if (!result.ExcelTags.Contains(tag))
                    status = "❌ Not in Excel";
                else if (!hasDxf)
                    status = "❌ Missing DXF";
                else if (!widthMatch || !heightMatch)
                {
                    if (!swappedMatch)
                        status = "❌ Dimension Mismatch";
                    else
                        status = "✅ OK (Dimensions Inversées)";
                }
                else if (qty <= 0)
                    status = "⚠️ Invalid Quantity";
                else
                    status = "✅ OK";

                double compareDxfWidth = isSwapped ? dxfHeight : dxfWidth;
                double compareDxfHeight = isSwapped ? dxfWidth : dxfHeight;

                //item.SubItems.Add(widthMatch ? "" : $"XL {excelWidth:F3} | DXF {compareDxfWidth:F3}");
                //item.SubItems.Add(heightMatch ? "" : $"XL {excelHeight:F3} | DXF {compareDxfHeight:F3}");

                item.SubItems.Add(status);

                // Add item to the list view
                listView.Items.Add(item);
            }

            listView.DrawSubItem += (s, e) =>
            {
                // Default to window background
                Color backgroundColor = SystemColors.Window;

                // Color status column
                if (e.ColumnIndex == 8) // Status column
                {
                    string status = e.Item.SubItems[8].Text;
                    if (status.StartsWith("❌"))
                        backgroundColor = Color.MistyRose;
                    else if (status.StartsWith("⚠️"))
                        backgroundColor = Color.LightYellow;
                    else if (status.StartsWith("✅"))
                        backgroundColor = Color.LightGreen;
                }
                // Color DXF, Width Match, Height Match columns
                else if (e.ColumnIndex == 3 || e.ColumnIndex == 4 || e.ColumnIndex == 5)
                {
                    if (e.Item.SubItems[e.ColumnIndex].Text == "✗")
                        backgroundColor = Color.LightSalmon;
                }

                using (var brush = new SolidBrush(backgroundColor))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds);
                }

                // Draw text
                TextRenderer.DrawText(e.Graphics, e.SubItem.Text, listView.Font, e.Bounds,
                    SystemColors.ControlText, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
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

            Controls.Add(containerPanel);
        }
        private void SortListView(int column, bool ascending)
        {
            listView.ListViewItemSorter = new ListViewItemComparer(column, ascending);
            listView.Sort();
        }

        private class ListViewItemComparer : IComparer
        {
            private readonly int column;
            private readonly bool ascending;

            public ListViewItemComparer(int column, bool ascending)
            {
                this.column = column;
                this.ascending = ascending;
            }

            public int Compare(object x, object y)
            {
                var itemX = (ListViewItem)x;
                var itemY = (ListViewItem)y;

                // Parse as numbers if possible
                if (int.TryParse(itemX.SubItems[column].Text, out int numX) &&
                    int.TryParse(itemY.SubItems[column].Text, out int numY))
                {
                    return ascending ? numX.CompareTo(numY) : numY.CompareTo(numX);
                }

                // Fall back to string comparison
                return ascending
                    ? string.Compare(itemX.SubItems[column].Text, itemY.SubItems[column].Text)
                    : string.Compare(itemY.SubItems[column].Text, itemX.SubItems[column].Text);
            }
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

    public class VerifDimCommand : IButtonManager
    {
        private string pathXl, pathDxf;
        private ValidationResultDimension lastValidationResult;

        public VerifDimCommand(TextBox txtBoxXl, TextBox txtBoxDxf)
        {
            pathXl = txtBoxXl.Text;
            pathDxf = txtBoxDxf.Text;
        }

        public void Execute()
        {
            if (string.IsNullOrEmpty(pathXl))
            {
                MessageBox.Show("Veuillez choisir un fichier Excel à analyser");
                return;
            }

            if (string.IsNullOrEmpty(pathDxf))
            {
                MessageBox.Show("Veuillez choisir un repertoir contenant des DXF");
                return;
            }

            if (!File.Exists(pathXl))
            {
                MessageBox.Show("Le fichier Excel spécifié n'existe pas.");
                return;
            }

            try
            {
                lastValidationResult = ValidateFiles();
                ShowValidationReport(lastValidationResult);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'exécution: {ex.Message}");
            }
        }

        private ValidationResultDimension ValidateFiles()
        {
            var result = new ValidationResultDimension();

            // Read Excel values
            var (excelTags, totalCount, tagQuantities, tagDimensions) = ReadColumnValuesFromProjetSheet(pathXl);
            result.ExcelTags = excelTags;
            result.TotalCount = totalCount;
            result.TagQuantities = tagQuantities;
            result.TagDimensions = tagDimensions;
            result.DimensionMismatches = new HashSet<string>();

            // Get DXF files and normalize filenames
            result.DxfFiles = GetFilesByExtension(pathDxf, "dxf")
                .Select(NormalizeFileName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Find missing files
            result.MissingDxf = new HashSet<string>(
                result.ExcelTags.Where(tag => !result.DxfFiles.Contains($"{tag}.dxf", StringComparer.OrdinalIgnoreCase)));

            // Find extra files
            result.ExtraDxf = new HashSet<string>(
                result.DxfFiles.Select(f => Path.GetFileNameWithoutExtension(f))
                    .Where(tag => !result.ExcelTags.Contains(tag, StringComparer.OrdinalIgnoreCase)));

            // Check dimensions for each tag
            foreach (var tag in result.ExcelTags)
            {
                if (result.DxfFiles.Contains($"{NormalizeFileName(tag)}.dxf"))
                {
                    var dxfDimensions = DxfDimensionExtractor.GetDxfDimensions(Path.Combine(pathDxf, $"{NormalizeFileName(tag)}.dxf"));

                    if (result.TagDimensions.TryGetValue(tag, out var excelDimensions))
                    {
                        bool widthMatch = Math.Truncate(excelDimensions.Width) == Math.Truncate(dxfDimensions.width);
                        bool heightMatch = Math.Truncate(excelDimensions.Height) == Math.Truncate(dxfDimensions.height);

                        if (!widthMatch || !heightMatch)
                        {
                            // Try swapped dimensions
                            bool swappedWidthMatch = Math.Truncate(excelDimensions.Width) == Math.Truncate(dxfDimensions.height);
                            bool swappedHeightMatch = Math.Truncate(excelDimensions.Height) == Math.Truncate(dxfDimensions.width);

                            if (swappedWidthMatch && swappedHeightMatch)
                            {
                                result.SwappedDimensions[tag] = true;
                            }
                            else
                            {
                                result.DimensionMismatches.Add(tag);
                            }
                        }
                    }
                }
            }

            result.QcPass = !result.MissingDxf.Any() &&
                    !result.ExtraDxf.Any() &&
                    !result.DimensionMismatches.Any() &&  // Added check for dimension mismatches
                    result.TagQuantities.All(kv => kv.Value > 0);

            return result;
        }

        private void ShowValidationReport(ValidationResultDimension result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== Validation Report ===");
            sb.AppendLine($"Total pieces count: {result.TotalCount}");
            sb.AppendLine($"Unique references in Excel: {result.ExcelTags.Count}");
            sb.AppendLine($"DXF files found: {result.DxfFiles.Count}");
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

            if (result.ExtraDxf.Any())
            {
                sb.AppendLine("Extra DXF files (not in Excel):");
                foreach (var tag in result.ExtraDxf.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            if (result.DimensionMismatches.Any())
            {
                sb.AppendLine("Files with dimension mismatches:");
                foreach (var tag in result.DimensionMismatches.OrderBy(x => x))
                {
                    var excelDim = result.TagDimensions[tag];
                    var dxfDim = DxfDimensionExtractor.GetDxfDimensions(Path.Combine(pathDxf, $"{NormalizeFileName(tag)}.dxf"));
                    sb.AppendLine($"- {tag}:");
                    sb.AppendLine($"  Excel: W={excelDim.Width:F2}, H={excelDim.Height:F2}");
                    sb.AppendLine($"  DXF:   W={dxfDim.width:F2}, H={dxfDim.height:F2}");
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
                var detailsForm = new DetailedComparisonFormDimension(lastValidationResult, pathDxf);
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

        private (HashSet<string> uniqueValues, int totalCount, Dictionary<string, int> tagQuantities, Dictionary<string, (double Width, double Height, int Row)> tagDimensions)
    ReadColumnValuesFromProjetSheet(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            HashSet<string> uniqueValues = new HashSet<string>();
            Dictionary<string, int> tagQuantities = new Dictionary<string, int>();
            Dictionary<string, (double Width, double Height, int Row)> tagDimensions = new Dictionary<string, (double Width, double Height, int Row)>();
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
                        var widthCell = worksheet.Cells[row, column + 2].Value;
                        var heightCell = worksheet.Cells[row, column + 3].Value;

                        if (mainCell != null && !string.IsNullOrWhiteSpace(mainCell.ToString()))
                        {
                            string value = mainCell.ToString().Trim();
                            if (value == "0") continue;
                            int quantity = 0;
                            double width = 0;
                            double height = 0;

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

                            // Parse width and height
                            if (widthCell != null && double.TryParse(widthCell.ToString(), out double parsedWidth))
                            {
                                width = parsedWidth;
                            }

                            if (heightCell != null && double.TryParse(heightCell.ToString(), out double parsedHeight))
                            {
                                height = parsedHeight;
                            }

                            uniqueValues.Add(value);
                            tagQuantities[value] = quantity;
                            tagDimensions[value] = (width, height, row); // Include the row number in the tuple
                            totalCount += quantity;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erreur lors de la lecture du fichier Excel: {ex.Message}");
            }

            return (uniqueValues, totalCount, tagQuantities, tagDimensions);
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