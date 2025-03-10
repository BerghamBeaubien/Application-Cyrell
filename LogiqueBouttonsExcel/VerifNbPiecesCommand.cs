// Ce fichier contient la logique de validation des quantités entre un fichier Excel et des fichiers DXF/STEP
// Il définit une classe pour stocker les résultats de validation, un formulaire détaillé pour afficher 
// les comparaisons, et une commande principale pour exécuter la vérification

using Application_Cyrell.LogiqueBouttonsSolidEdge;
using ExcelDataReader;
using static System.Net.Mime.MediaTypeNames;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System;
using System.Data;

namespace Application_Cyrell.LogiqueBouttonsExcel
{
    // Cette classe stocke les résultats de la validation entre les fichiers Excel et DXF/STEP
    // Elle conserve des ensembles d'étiquettes Excel, fichiers DXF, fichiers STEP, et les éléments manquants ou supplémentaires
    public class ValidationResultQte
    {
        public HashSet<string> ExcelTags { get; set; } = new HashSet<string>(); // Tags trouvés dans Excel
        public HashSet<string> DxfFiles { get; set; } = new HashSet<string>(); // Fichiers DXF trouvés
        public HashSet<string> StepFiles { get; set; } = new HashSet<string>(); // Fichiers STEP trouvés
        public int TotalCount { get; set; } // Nombre total de pièces
        public HashSet<string> MissingDxf { get; set; } = new HashSet<string>(); // Tags sans fichier DXF correspondant
        public HashSet<string> MissingStep { get; set; } = new HashSet<string>(); // Tags sans fichier STEP correspondant
        public HashSet<string> ExtraDxf { get; set; } = new HashSet<string>(); // Fichiers DXF sans tag correspondant
        public HashSet<string> ExtraStep { get; set; } = new HashSet<string>(); // Fichiers STEP sans tag correspondant
        public Dictionary<string, int> TagQuantities { get; set; } = new Dictionary<string, int>(); // Quantités par tag
        public bool QcPass { get; set; } // Indique si la validation est réussie
    }

    // Formulaire pour afficher une comparaison détaillée des résultats de validation
    // Montre chaque tag avec sa quantité et statut de présence des fichiers DXF/STEP
    public class DetailedComparisonFormQte : Form
    {
        public DetailedComparisonFormQte(ValidationResultQte result)
        {
            InitializeComponents(result);
        }

        // Initialise l'interface graphique du formulaire de comparaison détaillée
        private void InitializeComponents(ValidationResultQte result)
        {
            Text = "Comparaison Détaillée";
            Size = new Size(800, 600);

            // Création d'une vue en liste pour afficher les données
            var listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                OwnerDraw = true,
            };

            // Ajout des colonnes à la liste
            listView.Columns.Add("TAG", 200);
            listView.Columns.Add("Quantité", 70);
            listView.Columns.Add("DXF", 70);
            listView.Columns.Add("STEP", 70);
            listView.Columns.Add("Statut", 200);

            // Combinaison de tous les tags uniques pour l'affichage
            var allTags = new HashSet<string>();
            allTags.UnionWith(result.ExcelTags);
            allTags.UnionWith(result.DxfFiles.Select(f => Path.GetFileNameWithoutExtension(f)));
            allTags.UnionWith(result.StepFiles.Select(f => Path.GetFileNameWithoutExtension(f)));

            // Création des éléments de la liste pour chaque tag
            foreach (var tag in allTags.OrderBy(t => t))
            {
                var item = new ListViewItem(tag);

                // Affichage de la quantité
                result.TagQuantities.TryGetValue(tag, out int qty);
                item.SubItems.Add(qty.ToString());

                // Statut du fichier DXF (présent ou non)
                bool hasDxf = result.DxfFiles.Contains($"{NormalizeFileName(tag)}.dxf", StringComparer.OrdinalIgnoreCase);
                item.SubItems.Add(hasDxf ? "✓" : "✗");

                // Statut du fichier STEP (présent ou non)
                bool hasStep = result.StepFiles.Any(f =>
                    f.Equals($"{NormalizeFileName(tag)}.stp", StringComparison.OrdinalIgnoreCase) ||
                    f.Equals($"{NormalizeFileName(tag)}.step", StringComparison.OrdinalIgnoreCase));
                item.SubItems.Add(hasStep ? "✓" : "✗");

                // Statut global de la validation pour ce tag
                string status = "";
                if (!result.ExcelTags.Contains(tag))
                    status = "❌ Absent d'Excel";
                else if (!hasDxf && !hasStep)
                    status = "❌ Tous Fichiers Manquants";
                else if (!hasDxf)
                    status = "❌ DXF Manquant";
                else if (!hasStep)
                    status = "❌ STEP Manquant";
                else if (qty <= 0)
                    status = "⚠️ Quantité Invalide";
                else
                    status = "✅ OK";

                item.SubItems.Add(status);

                // Ajout de l'élément à la liste
                listView.Items.Add(item);
            }

            // Gestion du dessin personnalisé pour les sous-éléments
            listView.DrawSubItem += (s, e) =>
            {
                // Dessin par défaut pour les éléments et sous-éléments
                if (e.ColumnIndex == 4) // Colonne de statut
                {
                    if (e.Item.SubItems[e.ColumnIndex].Text != "✅ OK")
                        e.Graphics.FillRectangle(Brushes.MistyRose, e.Bounds);
                    else
                        e.Graphics.FillRectangle(SystemBrushes.Window, e.Bounds);
                }
                else if (e.ColumnIndex == 2 || e.ColumnIndex == 3) // Colonnes DXF ou STEP
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

                // Dessin du texte
                e.Graphics.DrawString(
                    e.SubItem.Text,
                    listView.Font,
                    SystemBrushes.ControlText,
                    e.Bounds);
            };

            // Dessin des en-têtes de colonnes
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

        // Normalise un nom de fichier en supprimant les espaces avant l'extension
        private string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return fileName;

            // Suppression des espaces avant l'extension
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

    // Classe principale qui implémente la vérification du nombre de pièces
    // Elle compare les données d'un fichier Excel avec les fichiers DXF et STEP dans les dossiers spécifiés
    public class VerifNbPiecesCommand : IButtonManager
    {
        private string pathXl, pathDxf, pathStep; // Chemins vers les fichiers et dossiers à vérifier
        private ValidationResultQte lastValidationResultQte; // Stocke le dernier résultat de validation
        //private const string xlpaath = "P:\\CYRAMP\\39981 FABBRICA 121 BROADWAY PO 4500023612 62  PANNES.xlsm";
        //private const string dxfpaath = "P:\\DESSINS\\FABBRICA USA\\121 BROADWAY\\BACKPAN\\1. PRODUCTION\\39981 PO 4500023612 Feb 12th\\DXF GALV 20GA";
        //private const string steppaath = "P:\\DESSINS\\FABBRICA USA\\121 BROADWAY\\BACKPAN\\1. PRODUCTION\\39981 PO 4500023612 Feb 12th\\STEP";

        // Constructeur qui prend les chemins des fichiers à partir des zones de texte
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

        // Exécute la commande de vérification
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

        // Valide les fichiers et retourne un objet de résultat de validation
        private ValidationResultQte ValidateFiles()
        {
            var result = new ValidationResultQte();

            // Lecture des valeurs Excel
            var (excelTags, totalCount, tagQuantities) = ReadColumnValuesFromProjetSheet(pathXl);
            result.ExcelTags = excelTags;
            result.TotalCount = totalCount;
            result.TagQuantities = tagQuantities;

            // Récupération des fichiers DXF et normalisation des noms
            result.DxfFiles = GetFilesByExtension(pathDxf, "dxf")
                .Select(NormalizeFileName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Récupération des fichiers Step et normalisation des noms
            result.StepFiles = GetFilesByExtension(pathStep, "stp")
                .Concat(GetFilesByExtension(pathStep, "step"))
                .Select(NormalizeFileName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Recherche des fichiers manquants
            result.MissingDxf = new HashSet<string>(
                result.ExcelTags.Where(tag => !result.DxfFiles.Contains($"{tag}.dxf", StringComparer.OrdinalIgnoreCase)));

            result.MissingStep = new HashSet<string>(
                result.ExcelTags.Where(tag =>
                    !result.StepFiles.Any(file =>
                        file.Equals($"{tag}.stp", StringComparison.OrdinalIgnoreCase) ||
                        file.Equals($"{tag}.step", StringComparison.OrdinalIgnoreCase))));

            // Recherche des fichiers supplémentaires
            result.ExtraDxf = new HashSet<string>(
                result.DxfFiles.Select(f => Path.GetFileNameWithoutExtension(f))
                    .Where(tag => !result.ExcelTags.Contains(tag, StringComparer.OrdinalIgnoreCase)));

            result.ExtraStep = new HashSet<string>(
                result.StepFiles.Select(f => Path.GetFileNameWithoutExtension(f))
                    .Where(tag => !result.ExcelTags.Contains(tag, StringComparer.OrdinalIgnoreCase)));

            // Détermination du résultat du contrôle qualité
            result.QcPass = !result.MissingDxf.Any() &&
                           !result.MissingStep.Any() &&
                           !result.ExtraDxf.Any() &&
                           !result.ExtraStep.Any() &&
                           result.TagQuantities.All(kv => kv.Value > 0);

            return result;
        }

        // Affiche un rapport de validation dans une fenêtre
        private void ShowValidationReport(ValidationResultQte result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== Rapport de Validation ===");
            sb.AppendLine($"Nombre total de pièces: {result.TotalCount}");
            sb.AppendLine($"Références uniques dans Excel: {result.ExcelTags.Count}");
            sb.AppendLine($"Fichiers DXF trouvés: {result.DxfFiles.Count}");
            sb.AppendLine($"Fichiers STEP trouvés: {result.StepFiles.Count}");
            sb.AppendLine();

            // Affichage des tags avec des quantités invalides
            if (result.TagQuantities.Any(kv => kv.Value <= 0))
            {
                sb.AppendLine("Tags avec quantités invalides:");
                foreach (var kv in result.TagQuantities.Where(kv => kv.Value <= 0))
                {
                    sb.AppendLine($"- {kv.Key}: {kv.Value}");
                }
                sb.AppendLine();
            }

            // Affichage des fichiers DXF manquants
            if (result.MissingDxf.Any())
            {
                sb.AppendLine("Fichiers DXF manquants:");
                foreach (var tag in result.MissingDxf.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            // Affichage des fichiers STEP manquants
            if (result.MissingStep.Any())
            {
                sb.AppendLine("Fichiers STEP manquants:");
                foreach (var tag in result.MissingStep.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            // Affichage des fichiers DXF supplémentaires
            if (result.ExtraDxf.Any())
            {
                sb.AppendLine("Fichiers DXF supplémentaires (non présents dans Excel):");
                foreach (var tag in result.ExtraDxf.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            // Affichage des fichiers STEP supplémentaires
            if (result.ExtraStep.Any())
            {
                sb.AppendLine("Fichiers STEP supplémentaires (non présents dans Excel):");
                foreach (var tag in result.ExtraStep.OrderBy(x => x))
                {
                    sb.AppendLine($"- {tag}");
                }
                sb.AppendLine();
            }

            sb.AppendLine($"Résultat du Contrôle Qualité: {(result.QcPass ? "RÉUSSI" : "ÉCHEC")}");

            // Création du formulaire de rapport
            var reportForm = new Form()
            {
                Text = "Rapport de Validation",
                Size = new Size(200, 400),
                StartPosition = FormStartPosition.CenterScreen
            };

            // Zone de texte pour afficher le rapport
            var textBox = new TextBox()
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Text = sb.ToString()
            };

            // Panneau pour les boutons
            var buttonPanel = new Panel()
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };

            // Bouton OK
            var okButton = new Button()
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Left = 10,
                Top = 8
            };

            // Bouton pour afficher les détails
            var detailsButton = new Button()
            {
                Text = "Afficher Détails",
                Width = 90,
                Left = 90,
                Top = 8
            };

            // Action du bouton de détails : ouvre le formulaire détaillé
            detailsButton.Click += (s, e) =>
            {
                var detailsForm = new DetailedComparisonFormQte(lastValidationResultQte);
                detailsForm.Show();
            };

            // Ajout des contrôles aux panneaux et formulaires
            buttonPanel.Controls.AddRange(new Control[] { okButton, detailsButton });
            reportForm.Controls.AddRange(new Control[] { textBox, buttonPanel });

            reportForm.ShowDialog();
        }

        // Récupère les fichiers par extension dans un répertoire
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
                // Gestion silencieuse des exceptions mais retour d'un ensemble vide
            }

            return files;
        }

        // Lit les valeurs des colonnes de la feuille PROJET du fichier Excel
        private (HashSet<string> uniqueValues, int totalCount, Dictionary<string, int> tagQuantities)
        ReadColumnValuesFromProjetSheet(string filePath)
        {
            HashSet<string> uniqueValues = new HashSet<string>();
            Dictionary<string, int> tagQuantities = new Dictionary<string, int>();
            int totalCount = 0;

            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // Détection automatique du format, prend en charge .xls, .xlsx, .xlsm
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Utilisation de l'objet ExcelDataSetConfiguration pour plus de contrôle
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = false
                            }
                        });

                        // Recherche de la feuille PROJET
                        DataTable projetSheet = null;

                        foreach (DataTable table in result.Tables)
                        {
                            if (table.TableName.Equals("PROJET", StringComparison.OrdinalIgnoreCase))
                            {
                                projetSheet = table;
                                break;
                            }
                        }

                        if (projetSheet == null)
                        {
                            throw new Exception("La feuille 'PROJET' n'a pas été trouvée.");
                        }

                        // Détermination de l'emplacement de l'en-tête
                        int startRow, column;
                        string headerD26 = GetCellValue(projetSheet, 25, 3); // D26 en index 0-based
                        string headerC17 = GetCellValue(projetSheet, 16, 2); // C17 en index 0-based

                        if (headerD26 == "TAG #")
                        {
                            startRow = 27; // 28 en index 1-based
                            column = 3;  // Colonne D en index 0-based
                        }
                        else if (headerC17 == "TAG #")
                        {
                            startRow = 18; // 19 en index 1-based
                            column = 2;  // Colonne C en index 0-based
                        }
                        else
                        {
                            throw new Exception("En-tête 'TAG #' introuvable dans D26 ou C17");
                        }

                        int endRow = projetSheet.Rows.Count;
                        int actualLastRow = startRow;

                        // Recherche de la dernière ligne avec des données
                        for (int row = startRow; row < endRow; row++)
                        {
                            string cellValue = GetCellValue(projetSheet, row, column);
                            if (!string.IsNullOrWhiteSpace(cellValue))
                            {
                                actualLastRow = row;
                            }
                        }

                        // Traitement des données
                        for (int row = startRow; row <= actualLastRow; row++)
                        {
                            if (row == 500) continue; // Ignorer la ligne 501 (1-based) qui est 500 en index 0-based

                            string mainCellValue = GetCellValue(projetSheet, row, column);
                            if (string.IsNullOrWhiteSpace(mainCellValue)) continue;

                            string value = mainCellValue.Trim();
                            if (value == "0") continue;

                            int quantity = 0;

                            // Analyse de la quantité
                            string quantityCellValue = GetCellValue(projetSheet, row, column + 1);
                            if (!string.IsNullOrWhiteSpace(quantityCellValue))
                            {
                                if (int.TryParse(quantityCellValue, out int parsedQuantity))
                                {
                                    quantity = parsedQuantity;
                                }
                                else if (quantityCellValue.Contains("*"))
                                {
                                    string[] factors = quantityCellValue.Split('*');
                                    quantity = factors.All(f => int.TryParse(f.Trim(), out int n))
                                        ? factors.Select(f => int.Parse(f.Trim())).Aggregate((a, b) => a * b)
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
                throw new Exception($"Erreur lors de la lecture du fichier Excel: {ex.Message}", ex);
            }

            return (uniqueValues, totalCount, tagQuantities);
        }

        // Méthode auxiliaire pour obtenir en toute sécurité les valeurs des cellules
        private string GetCellValue(DataTable sheet, int row, int column)
        {
            try
            {
                if (row < 0 || column < 0 || row >= sheet.Rows.Count || column >= sheet.Columns.Count)
                {
                    return string.Empty;
                }

                var cellValue = sheet.Rows[row][column];
                return cellValue == null || cellValue == DBNull.Value ? string.Empty : cellValue.ToString().Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        // Normalise un nom de fichier en supprimant les espaces avant l'extension
        private string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return fileName;

            // Suppression des espaces avant l'extension
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