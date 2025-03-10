using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Application_Cyrell.LogiqueBouttonsSolidEdge;
using ExcelDataReader;
using Path = System.IO.Path;

namespace Application_Cyrell.LogiqueBouttonsExcel
{
    // ValidationResultDimension class remains unchanged
    public class ValidationResultDimension
    {
        public HashSet<string> ExcelTags { get; set; } = new HashSet<string>();
        public HashSet<string> DxfFiles { get; set; } = new HashSet<string>();
        public int TotalCount { get; set; }
        public HashSet<string> MissingDxf { get; set; } = new HashSet<string>();
        public HashSet<string> ExtraDxf { get; set; } = new HashSet<string>();
        public Dictionary<string, int> TagQuantities { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, (double Width, double Height, int Row)> TagDimensions { get; set; } = new Dictionary<string, (double Width, double Height, int Row)>();
        public HashSet<string> DimensionMismatches { get; set; } = new HashSet<string>();  // Pour suivre les différences de dimensions
        public bool QcPass { get; set; }
        public Dictionary<string, bool> SwappedDimensions { get; set; } = new Dictionary<string, bool>();
    }

    // DetailedComparisonFormDimension class remains unchanged
    public class DetailedComparisonFormDimension : Form
    {
        private readonly string _pathDxf;
        private ListView listView;
        private MenuStrip menuStrip;

        public DetailedComparisonFormDimension(ValidationResultDimension result, string pathDxf)
        {
            _pathDxf = pathDxf;
            InitializeComponents(result);
        }

        private void InitializeComponents(ValidationResultDimension result)
        {
            // Configuration de la fenêtre
            Text = "Comparaison Détaillée";
            Size = new Size(1000, 600);

            // Création de la ListView
            listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                OwnerDraw = true
            };
            Controls.Add(listView);

            // Création du menu
            menuStrip = new MenuStrip();
            Controls.Add(menuStrip);
            menuStrip.Dock = DockStyle.Top;

            // Configuration des options de tri
            var sortMenu = new ToolStripMenuItem("Trier");

            var sortByRowAsc = new ToolStripMenuItem("Trier par Rangée (Ordre Excel)");
            sortByRowAsc.Click += (s, e) => SortListView(0, true);

            var sortByTagAsc = new ToolStripMenuItem("Trier par TAG (A-Z)");
            sortByTagAsc.Click += (s, e) => SortListView(1, true);

            sortMenu.DropDownItems.Add(sortByRowAsc);
            sortMenu.DropDownItems.Add(sortByTagAsc);
            menuStrip.Items.Add(sortMenu);

            // Configuration des colonnes de la ListView
            listView.Columns.Add("Rangée", 50);
            listView.Columns.Add("TAG", 170);
            listView.Columns.Add("Qté", 40);
            listView.Columns.Add("DXF", 40);
            listView.Columns.Add("Largeur", 45);
            listView.Columns.Add("Hauteur", 45);
            listView.Columns.Add("Problème Largeur", 140);
            listView.Columns.Add("Problème Hauteur", 140);
            listView.Columns.Add("Statut", 200);

            // Combiner tous les tags uniques
            var allTags = new HashSet<string>();
            allTags.UnionWith(result.ExcelTags);
            allTags.UnionWith(result.DxfFiles.Select(f => Path.GetFileNameWithoutExtension(f)));

            // Remplir la ListView avec les données
            PopulateListView(allTags, result);

            // Configuration des événements de dessin
            ConfigureDrawingEvents();
        }

        // Méthode pour remplir la ListView avec les données
        private void PopulateListView(HashSet<string> allTags, ValidationResultDimension result)
        {
            foreach (var tag in allTags.OrderBy(t => t))
            {
                var item = new ListViewItem("");

                // Ajouter le numéro de ligne
                int rowNumber = result.TagDimensions.TryGetValue(tag, out var dim) ? dim.Row : 0;
                item.Text = rowNumber.ToString();

                item.SubItems.Add(tag);

                // Quantité
                result.TagQuantities.TryGetValue(tag, out int qty);
                item.SubItems.Add(qty.ToString());

                // Statut DXF
                bool hasDxf = result.DxfFiles.Contains($"{NormalizeFileName(tag)}.dxf", StringComparer.OrdinalIgnoreCase);
                item.SubItems.Add(hasDxf ? "✓" : "✗");

                // Dimensions Excel
                double excelWidth = 0;
                double excelHeight = 0;
                if (result.TagDimensions.TryGetValue(tag, out var excelDimensions))
                {
                    excelWidth = excelDimensions.Width;
                    excelHeight = excelDimensions.Height;
                }

                // Dimensions DXF
                double dxfWidth = 0;
                double dxfHeight = 0;
                if (hasDxf)
                {
                    var dxfDimensions = DxfDimensionExtractor.GetDxfDimensions(Path.Combine(_pathDxf, $"{NormalizeFileName(tag)}.dxf"));
                    dxfWidth = dxfDimensions.width;
                    dxfHeight = dxfDimensions.height;
                }

                // Comparaison moins stricte (valeurs tronquées)
                bool widthMatch = Math.Truncate(excelWidth) == Math.Truncate(dxfWidth);
                bool heightMatch = Math.Truncate(excelHeight) == Math.Truncate(dxfHeight);
                bool swappedMatch = false;

                // Vérifier si les dimensions sont inversées
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

                // Colonnes de correspondance
                item.SubItems.Add(widthMatch ? "✓" : "✗");
                item.SubItems.Add(heightMatch ? "✓" : "✗");

                // Colonnes de différences
                item.SubItems.Add(widthMatch ? "" : $"XL {excelWidth:F3} | DXF {dxfWidth:F3}");
                item.SubItems.Add(heightMatch ? "" : $"XL {excelHeight:F3} | DXF {dxfHeight:F3}");

                // Déterminer le statut global
                string status = DetermineStatus(tag, hasDxf, widthMatch, heightMatch, swappedMatch, qty, result);
                item.SubItems.Add(status);

                // Ajouter l'élément à la ListView
                listView.Items.Add(item);
            }
        }

        // Déterminer le statut global d'un élément
        private string DetermineStatus(string tag, bool hasDxf, bool widthMatch, bool heightMatch,
                                     bool swappedMatch, int qty, ValidationResultDimension result)
        {
            if (!result.ExcelTags.Contains(tag))
                return "❌ Pas dans Excel";
            else if (!hasDxf)
                return "❌ DXF Manquant";
            else if (!widthMatch || !heightMatch)
            {
                if (!swappedMatch)
                    return "❌ Dimensions Incompatibles";
                else
                    return "✅ OK (Dimensions Inversées)";
            }
            else if (qty <= 0)
                return "⚠️ Quantité Invalide";
            else
                return "✅ OK";
        }

        // Configuration des événements de dessin
        private void ConfigureDrawingEvents()
        {
            // Événement de dessin des sous-éléments
            listView.DrawSubItem += (s, e) =>
            {
                // Couleur de fond par défaut
                Color backgroundColor = SystemColors.Window;

                // Colorer la colonne de statut
                if (e.ColumnIndex == 8) // Colonne de statut
                {
                    string status = e.Item.SubItems[8].Text;
                    if (status.StartsWith("❌"))
                        backgroundColor = Color.MistyRose;
                    else if (status.StartsWith("⚠️"))
                        backgroundColor = Color.LightYellow;
                    else if (status.StartsWith("✅"))
                        backgroundColor = Color.LightGreen;
                }
                // Colorer les colonnes DXF, Largeur, Hauteur
                else if (e.ColumnIndex == 3 || e.ColumnIndex == 4 || e.ColumnIndex == 5)
                {
                    if (e.Item.SubItems[e.ColumnIndex].Text == "✗")
                        backgroundColor = Color.LightSalmon;
                }

                // Dessiner le fond
                using (var brush = new SolidBrush(backgroundColor))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds);
                }

                // Dessiner le texte
                TextRenderer.DrawText(e.Graphics, e.SubItem.Text, listView.Font, e.Bounds,
                    SystemColors.ControlText, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            };

            // Événement de dessin des en-têtes de colonne
            listView.DrawColumnHeader += (s, e) =>
            {
                e.DrawBackground();
                e.Graphics.DrawString(
                    e.Header.Text,
                    listView.Font,
                    SystemBrushes.ControlText,
                    e.Bounds);
            };
        }

        // Méthode pour trier la ListView
        private void SortListView(int column, bool ascending)
        {
            listView.ListViewItemSorter = new ListViewItemComparer(column, ascending);
            listView.Sort();
        }

        // Classe pour comparer les éléments de la ListView
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

                // Essayer de parser comme nombres si possible
                if (int.TryParse(itemX.SubItems[column].Text, out int numX) &&
                    int.TryParse(itemY.SubItems[column].Text, out int numY))
                {
                    return ascending ? numX.CompareTo(numY) : numY.CompareTo(numX);
                }

                // Sinon, comparer comme chaînes
                return ascending
                    ? string.Compare(itemX.SubItems[column].Text, itemY.SubItems[column].Text)
                    : string.Compare(itemY.SubItems[column].Text, itemX.SubItems[column].Text);
            }
        }

        // Normaliser le nom de fichier (supprimer les espaces avant l'extension)
        private string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return fileName;

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
            // Vérifier que les fichiers/dossiers sont spécifiés
            if (string.IsNullOrEmpty(pathXl))
            {
                MessageBox.Show("Veuillez choisir un fichier Excel à analyser");
                return;
            }

            if (string.IsNullOrEmpty(pathDxf))
            {
                MessageBox.Show("Veuillez choisir un répertoire contenant des DXF");
                return;
            }

            if (!File.Exists(pathXl))
            {
                MessageBox.Show("Le fichier Excel spécifié n'existe pas.");
                return;
            }

            try
            {
                // Effectuer la validation et afficher le rapport
                lastValidationResult = ValidateFiles();
                ShowValidationReport(lastValidationResult);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'exécution: {ex.Message}");
            }
        }

        // Valider les fichiers Excel et DXF
        private ValidationResultDimension ValidateFiles()
        {
            var result = new ValidationResultDimension();

            // Lire les valeurs d'Excel
            var (excelTags, totalCount, tagQuantities, tagDimensions) = ReadColumnValuesFromProjetSheet(pathXl);
            result.ExcelTags = excelTags;
            result.TotalCount = totalCount;
            result.TagQuantities = tagQuantities;
            result.TagDimensions = tagDimensions;
            result.DimensionMismatches = new HashSet<string>();

            // Obtenir les fichiers DXF et normaliser les noms de fichier
            result.DxfFiles = GetFilesByExtension(pathDxf, "dxf")
                .Select(NormalizeFileName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Trouver les fichiers manquants
            result.MissingDxf = new HashSet<string>(
                result.ExcelTags.Where(tag => !result.DxfFiles.Contains($"{tag}.dxf", StringComparer.OrdinalIgnoreCase)));

            // Trouver les fichiers supplémentaires
            result.ExtraDxf = new HashSet<string>(
                result.DxfFiles.Select(f => Path.GetFileNameWithoutExtension(f))
                    .Where(tag => !result.ExcelTags.Contains(tag, StringComparer.OrdinalIgnoreCase)));

            // Vérifier les dimensions pour chaque tag
            CheckDimensions(result);

            // Déterminer si le QC passe
            result.QcPass = !result.MissingDxf.Any() &&
                    !result.ExtraDxf.Any() &&
                    !result.DimensionMismatches.Any() &&
                    result.TagQuantities.All(kv => kv.Value > 0);

            return result;
        }

        // Vérifier les dimensions pour tous les tags
        private void CheckDimensions(ValidationResultDimension result)
        {
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
                            // Essayer avec les dimensions inversées
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
        }

        // Afficher le rapport de validation
        private void ShowValidationReport(ValidationResultDimension result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== Rapport de Validation ===");
            sb.AppendLine($"Nombre total de pièces: {result.TotalCount}");
            sb.AppendLine($"Références uniques dans Excel: {result.ExcelTags.Count}");
            sb.AppendLine($"Fichiers DXF trouvés: {result.DxfFiles.Count}");
            sb.AppendLine();

            // Vérifier les quantités invalides
            if (result.TagQuantities.Any(kv => kv.Value <= 0))
            {
                sb.AppendLine("Tags avec quantités invalides:");
                foreach (var kv in result.TagQuantities.Where(kv => kv.Value <= 0))
                {
                    sb.AppendLine($"- {kv.Key}: {kv.Value}");
                }
                sb.AppendLine();
            }

            // Vérifier les DXF manquants
            if (result.MissingDxf.Any())
            {
                sb.AppendLine("Fichiers DXF manquants:");
                foreach (var tag in result.MissingDxf.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            // Vérifier les DXF supplémentaires
            if (result.ExtraDxf.Any())
            {
                sb.AppendLine("Fichiers DXF supplémentaires (pas dans Excel):");
                foreach (var tag in result.ExtraDxf.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            // Vérifier les différences de dimensions
            if (result.DimensionMismatches.Any())
            {
                sb.AppendLine("Fichiers avec différences de dimensions:");
                foreach (var tag in result.DimensionMismatches.OrderBy(x => x))
                {
                    var excelDim = result.TagDimensions[tag];
                    var dxfDim = DxfDimensionExtractor.GetDxfDimensions(Path.Combine(pathDxf, $"{NormalizeFileName(tag)}.dxf"));
                    sb.AppendLine($"- {tag}:");
                    sb.AppendLine($"  Excel: L={excelDim.Width:F2}, H={excelDim.Height:F2}");
                    sb.AppendLine($"  DXF:   L={dxfDim.width:F2}, H={dxfDim.height:F2}");
                }
                sb.AppendLine();
            }

            sb.AppendLine($"Résultat du Contrôle Qualité: {(result.QcPass ? "RÉUSSI" : "ÉCHOUÉ")}");

            // Créer et afficher le formulaire de rapport
            CreateAndShowReportForm(sb.ToString());
        }

        // Créer et afficher le formulaire de rapport
        private void CreateAndShowReportForm(string reportText)
        {
            var reportForm = new Form()
            {
                Text = "Rapport de Validation",
                Size = new Size(600, 400),
                StartPosition = FormStartPosition.CenterScreen
            };

            var textBox = new TextBox()
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Text = reportText
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
                Text = "Afficher Détails",
                Width = 120,
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

        // Obtenir les fichiers par extension
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
                // Gérer les exceptions silencieusement mais retourner un ensemble vide
            }

            return files;
        }

        // Lire les valeurs de colonnes de la feuille PROJET - Converti pour ClosedXML
        private (HashSet<string> uniqueValues, int totalCount, Dictionary<string, int> tagQuantities, Dictionary<string, (double Width, double Height, int Row)> tagDimensions)
        ReadColumnValuesFromProjetSheet(string filePath)
        {
            HashSet<string> uniqueValues = new HashSet<string>();
            Dictionary<string, int> tagQuantities = new Dictionary<string, int>();
            Dictionary<string, (double Width, double Height, int Row)> tagDimensions = new Dictionary<string, (double Width, double Height, int Row)>();
            int totalCount = 0;

            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = false }
                        });

                        // Chercher la feuille PROJET
                        DataTable worksheet = null;
                        foreach (DataTable table in result.Tables)
                        {
                            if (table.TableName.Equals("PROJET", StringComparison.OrdinalIgnoreCase))
                            {
                                worksheet = table;
                                break;
                            }
                        }

                        if (worksheet == null)
                        {
                            throw new Exception("La feuille 'PROJET' n'a pas été trouvée.");
                        }

                        // Déterminer l'emplacement de l'en-tête
                        int startRow, column;
                        var headerD26 = worksheet.Rows[25][3]?.ToString().Trim() ?? string.Empty;
                        var headerC17 = worksheet.Rows[16][2]?.ToString().Trim() ?? string.Empty;

                        if (headerD26 == "TAG #")
                        {
                            startRow = 27; // 0-based indexing
                            column = 3;  // 0-based indexing
                        }
                        else if (headerC17 == "TAG #")
                        {
                            startRow = 18; // 0-based indexing
                            column = 2;  // 0-based indexing
                        }
                        else
                        {
                            throw new Exception("En-tête 'TAG #' introuvable dans D26 ou C17");
                        }

                        // Trouver la dernière ligne avec des données
                        int lastRow = worksheet.Rows.Count - 1;
                        int actualLastRow = startRow;

                        for (int row = startRow; row <= lastRow; row++)
                        {
                            if (row >= worksheet.Rows.Count) break;

                            var cell = worksheet.Rows[row][column];
                            if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                            {
                                actualLastRow = row;
                            }
                        }

                        // Traiter les données
                        ProcessExcelData(worksheet, startRow, actualLastRow, column, uniqueValues, tagQuantities, tagDimensions, ref totalCount);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erreur lors de la lecture du fichier Excel: {ex.Message}");
            }

            return (uniqueValues, totalCount, tagQuantities, tagDimensions);
        }

        // Traiter les données du fichier Excel - Adapté pour ExcelDataReader
        private void ProcessExcelData(DataTable worksheet, int startRow, int actualLastRow, int column,
                                    HashSet<string> uniqueValues, Dictionary<string, int> tagQuantities,
                                    Dictionary<string, (double Width, double Height, int Row)> tagDimensions, ref int totalCount)
        {
            for (int row = startRow; row <= actualLastRow; row++)
            {
                if (row == 500) continue; // Ligne à ignorer (500 pour 0-based indexing)
                if (row >= worksheet.Rows.Count) break;

                var mainCell = worksheet.Rows[row][column];
                var quantityCell = worksheet.Rows[row][column + 1];
                var widthCell = worksheet.Rows[row][column + 2];
                var heightCell = worksheet.Rows[row][column + 3];

                if (mainCell != null && !string.IsNullOrWhiteSpace(mainCell.ToString()))
                {
                    string value = mainCell.ToString().Trim();
                    if (value == "0") continue;

                    // Analyser la quantité
                    int quantity = ParseQuantity(quantityCell);

                    // Analyser la largeur et la hauteur
                    double width = 0, height = 0;
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
                    tagDimensions[value] = (width, height, row + 1); // Convertir en 1-based pour correspondre à l'indexation Excel
                    totalCount += quantity;
                }
            }
        }

        // Méthode auxiliaire pour analyser la quantité
        private int ParseQuantity(object cellValue)
        {
            if (cellValue == null) return 0;

            var valueStr = cellValue.ToString();
            if (string.IsNullOrWhiteSpace(valueStr)) return 0;

            if (int.TryParse(valueStr, out int quantity))
            {
                return quantity;
            }
            return 0;
        }

        // Normaliser le nom de fichier
        private string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return fileName;

            // Supprimer les espaces avant l'extension
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