using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Application_Cyrell;
using Application_Cyrell.LogiqueBouttonsSolidEdge;
using System.Drawing;
using Application_Cyrell.Properties;

namespace firstCSMacro
{
    //static class Program
    //{
    //    [STAThread]
    //    static void Main()
    //    {
    //        System.Windows.Forms.Application.EnableVisualStyles();
    //        System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
    //        System.Windows.Forms.Application.Run(new MainForm());
    //    }
    //}

    public partial class PanelSE : Form
    {
        private TextBox textBoxFolderPath;
        public ListBox listBoxDxfFiles;
        private Button button10;
        private Button buttonKillSe;
        private FlowLayoutPanel filterPanel;
        private PanelSettings _panelSettings;
        public bool paramFermerSe;
        private Button btnSettings;
        private Button btnBrowseSe;
        private Button buttonTagDxf;
        private Button buttonOuvrirFichiers;
        private Button buttonExportDim;
        private Button buttonSaveDxfStep;
        private Button buttonGenererDFT;
        private CancellationTokenSource cancelTokenTag;
        private CustomTooltipForm customTooltipTag;
        private CancellationTokenSource cancelTokenDim;
        private CustomTooltipForm customTooltipDimensions;
        private CancellationTokenSource cancelTokenDft;
        private CustomTooltipForm customTooltipDft;
        private TextBox txtBoxUnlock;
        private PictureBox picBoxArrow;
        private Dictionary<string, bool> extensionFilters = new Dictionary<string, bool>()
        {
            { ".asm", true },
            { ".dxf", true },
            { ".stp", true },
            { ".step", true },
            { ".par", true },
            { ".psm", true }
        };

        public PanelSE()
        {
            InitializeComponent();
            InitializeFilterControls();
        }

        private void InitializeComponent()
        {
            this.listBoxDxfFiles = new System.Windows.Forms.ListBox();
            this.textBoxFolderPath = new System.Windows.Forms.TextBox();
            this.button10 = new System.Windows.Forms.Button();
            this.buttonKillSe = new System.Windows.Forms.Button();
            this.filterPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnSettings = new System.Windows.Forms.Button();
            this.btnBrowseSe = new System.Windows.Forms.Button();
            this.buttonTagDxf = new System.Windows.Forms.Button();
            this.buttonOuvrirFichiers = new System.Windows.Forms.Button();
            this.buttonExportDim = new System.Windows.Forms.Button();
            this.buttonSaveDxfStep = new System.Windows.Forms.Button();
            this.buttonGenererDFT = new System.Windows.Forms.Button();
            this.txtBoxUnlock = new System.Windows.Forms.TextBox();
            this.picBoxArrow = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxArrow)).BeginInit();
            this.SuspendLayout();
            // 
            // listBoxDxfFiles
            // 
            this.listBoxDxfFiles.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.listBoxDxfFiles.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.listBoxDxfFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBoxDxfFiles.ForeColor = System.Drawing.Color.White;
            this.listBoxDxfFiles.FormattingEnabled = true;
            this.listBoxDxfFiles.ItemHeight = 36;
            this.listBoxDxfFiles.Location = new System.Drawing.Point(84, 184);
            this.listBoxDxfFiles.Name = "listBoxDxfFiles";
            this.listBoxDxfFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBoxDxfFiles.Size = new System.Drawing.Size(902, 468);
            this.listBoxDxfFiles.TabIndex = 4;
            // 
            // textBoxFolderPath
            // 
            this.textBoxFolderPath.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.textBoxFolderPath.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBoxFolderPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxFolderPath.ForeColor = System.Drawing.Color.White;
            this.textBoxFolderPath.Location = new System.Drawing.Point(84, 134);
            this.textBoxFolderPath.Name = "textBoxFolderPath";
            this.textBoxFolderPath.Size = new System.Drawing.Size(902, 28);
            this.textBoxFolderPath.TabIndex = 5;
            // 
            // button10
            // 
            this.button10.BackColor = System.Drawing.Color.DarkOrange;
            this.button10.FlatAppearance.BorderSize = 0;
            this.button10.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button10.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button10.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button10.Location = new System.Drawing.Point(532, 700);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(129, 40);
            this.button10.TabIndex = 7;
            this.button10.Text = "Select All";
            this.button10.UseVisualStyleBackColor = false;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // buttonKillSe
            // 
            this.buttonKillSe.FlatAppearance.BorderSize = 0;
            this.buttonKillSe.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonKillSe.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.buttonKillSe.ForeColor = System.Drawing.Color.OrangeRed;
            this.buttonKillSe.Location = new System.Drawing.Point(726, 68);
            this.buttonKillSe.Name = "buttonKillSe";
            this.buttonKillSe.Size = new System.Drawing.Size(220, 40);
            this.buttonKillSe.TabIndex = 8;
            this.buttonKillSe.Text = "FERMER SOLID EDGE";
            this.buttonKillSe.UseVisualStyleBackColor = true;
            this.buttonKillSe.Click += new System.EventHandler(this.buttonKillSe_Click);
            // 
            // filterPanel
            // 
            this.filterPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.filterPanel.Location = new System.Drawing.Point(84, 706);
            this.filterPanel.Name = "filterPanel";
            this.filterPanel.Size = new System.Drawing.Size(420, 25);
            this.filterPanel.TabIndex = 0;
            // 
            // btnSettings
            // 
            this.btnSettings.BackColor = System.Drawing.Color.Transparent;
            this.btnSettings.BackgroundImage = global::Application_Cyrell.Properties.Resources.logoParam;
            this.btnSettings.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnSettings.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSettings.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.btnSettings.FlatAppearance.BorderSize = 0;
            this.btnSettings.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.btnSettings.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSettings.Location = new System.Drawing.Point(998, 685);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnSettings.Size = new System.Drawing.Size(74, 74);
            this.btnSettings.TabIndex = 1;
            this.btnSettings.UseVisualStyleBackColor = false;
            this.btnSettings.Click += new System.EventHandler(this.btnSettings_Click);
            // 
            // btnBrowseSe
            // 
            this.btnBrowseSe.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnBrowseSe.BackgroundImage = global::Application_Cyrell.Properties.Resources.search_in_folder;
            this.btnBrowseSe.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBrowseSe.FlatAppearance.BorderSize = 0;
            this.btnBrowseSe.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowseSe.Location = new System.Drawing.Point(998, 109);
            this.btnBrowseSe.Name = "btnBrowseSe";
            this.btnBrowseSe.Size = new System.Drawing.Size(65, 65);
            this.btnBrowseSe.TabIndex = 13;
            this.btnBrowseSe.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnBrowseSe.UseVisualStyleBackColor = true;
            this.btnBrowseSe.Click += new System.EventHandler(this.btnBrowseSe_Click);
            // 
            // buttonTagDxf
            // 
            this.buttonTagDxf.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonTagDxf.FlatAppearance.BorderSize = 0;
            this.buttonTagDxf.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonTagDxf.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonTagDxf.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.buttonTagDxf.Location = new System.Drawing.Point(179, 12);
            this.buttonTagDxf.Name = "buttonTagDxf";
            this.buttonTagDxf.Size = new System.Drawing.Size(234, 40);
            this.buttonTagDxf.TabIndex = 14;
            this.buttonTagDxf.Text = "Annoter DXF (Tag)";
            this.buttonTagDxf.UseVisualStyleBackColor = false;
            this.buttonTagDxf.Visible = false;
            this.buttonTagDxf.Click += new System.EventHandler(this.buttonTagDxf_Click);
            this.buttonTagDxf.MouseEnter += new System.EventHandler(this.btnTaguerDxf_MouseEnter);
            this.buttonTagDxf.MouseLeave += new System.EventHandler(this.btnTaguerDxf_MouseLeave);
            // 
            // buttonOuvrirFichiers
            // 
            this.buttonOuvrirFichiers.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonOuvrirFichiers.FlatAppearance.BorderSize = 0;
            this.buttonOuvrirFichiers.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOuvrirFichiers.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonOuvrirFichiers.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.buttonOuvrirFichiers.Location = new System.Drawing.Point(439, 12);
            this.buttonOuvrirFichiers.Name = "buttonOuvrirFichiers";
            this.buttonOuvrirFichiers.Size = new System.Drawing.Size(234, 40);
            this.buttonOuvrirFichiers.TabIndex = 15;
            this.buttonOuvrirFichiers.Text = "Ouvrir Fichiers Choisis";
            this.buttonOuvrirFichiers.UseVisualStyleBackColor = false;
            this.buttonOuvrirFichiers.Visible = false;
            this.buttonOuvrirFichiers.Click += new System.EventHandler(this.buttonOuvrirFichiers_Click);
            // 
            // buttonExportDim
            // 
            this.buttonExportDim.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonExportDim.FlatAppearance.BorderSize = 0;
            this.buttonExportDim.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonExportDim.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonExportDim.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.buttonExportDim.Location = new System.Drawing.Point(726, 12);
            this.buttonExportDim.Name = "buttonExportDim";
            this.buttonExportDim.Size = new System.Drawing.Size(220, 40);
            this.buttonExportDim.TabIndex = 16;
            this.buttonExportDim.Text = "Exporter Dimensions";
            this.buttonExportDim.UseVisualStyleBackColor = false;
            this.buttonExportDim.Visible = false;
            this.buttonExportDim.Click += new System.EventHandler(this.buttonExportDim_Click);
            this.buttonExportDim.MouseEnter += new System.EventHandler(this.btnExporterDim_MouseEnter);
            this.buttonExportDim.MouseLeave += new System.EventHandler(this.btnExporterDim_MouseLeave);
            // 
            // buttonSaveDxfStep
            // 
            this.buttonSaveDxfStep.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonSaveDxfStep.FlatAppearance.BorderSize = 0;
            this.buttonSaveDxfStep.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSaveDxfStep.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSaveDxfStep.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.buttonSaveDxfStep.Location = new System.Drawing.Point(179, 68);
            this.buttonSaveDxfStep.Name = "buttonSaveDxfStep";
            this.buttonSaveDxfStep.Size = new System.Drawing.Size(234, 40);
            this.buttonSaveDxfStep.TabIndex = 17;
            this.buttonSaveDxfStep.Text = "Sauvegarder DXF && Step";
            this.buttonSaveDxfStep.UseVisualStyleBackColor = false;
            this.buttonSaveDxfStep.Visible = false;
            this.buttonSaveDxfStep.Click += new System.EventHandler(this.buttonSaveDxfStep_Click);
            // 
            // buttonGenererDFT
            // 
            this.buttonGenererDFT.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonGenererDFT.FlatAppearance.BorderSize = 0;
            this.buttonGenererDFT.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonGenererDFT.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonGenererDFT.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.buttonGenererDFT.Location = new System.Drawing.Point(439, 68);
            this.buttonGenererDFT.Name = "buttonGenererDFT";
            this.buttonGenererDFT.Size = new System.Drawing.Size(234, 40);
            this.buttonGenererDFT.TabIndex = 18;
            this.buttonGenererDFT.Text = "Générer Dessins (DFT)";
            this.buttonGenererDFT.UseVisualStyleBackColor = false;
            this.buttonGenererDFT.Visible = false;
            this.buttonGenererDFT.Click += new System.EventHandler(this.buttonGenererDFT_Click);
            this.buttonGenererDFT.MouseEnter += new System.EventHandler(this.btnGenererDft_MouseEnter);
            this.buttonGenererDFT.MouseLeave += new System.EventHandler(this.btnGenererDft_MouseLeave);
            // 
            // txtBoxUnlock
            // 
            this.txtBoxUnlock.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.txtBoxUnlock.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtBoxUnlock.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxUnlock.ForeColor = System.Drawing.Color.LightCyan;
            this.txtBoxUnlock.Location = new System.Drawing.Point(51, 7);
            this.txtBoxUnlock.Name = "txtBoxUnlock";
            this.txtBoxUnlock.Size = new System.Drawing.Size(950, 82);
            this.txtBoxUnlock.TabIndex = 19;
            this.txtBoxUnlock.Text = "Veiullez Choisir un répértoire pour continuer";
            // 
            // picBoxArrow
            // 
            this.picBoxArrow.BackgroundImage = global::Application_Cyrell.Properties.Resources.logoArrow;
            this.picBoxArrow.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.picBoxArrow.Location = new System.Drawing.Point(998, 25);
            this.picBoxArrow.Name = "picBoxArrow";
            this.picBoxArrow.Size = new System.Drawing.Size(65, 63);
            this.picBoxArrow.TabIndex = 20;
            this.picBoxArrow.TabStop = false;
            // 
            // PanelSE
            // 
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.ClientSize = new System.Drawing.Size(1123, 798);
            this.Controls.Add(this.picBoxArrow);
            this.Controls.Add(this.txtBoxUnlock);
            this.Controls.Add(this.buttonGenererDFT);
            this.Controls.Add(this.buttonSaveDxfStep);
            this.Controls.Add(this.buttonExportDim);
            this.Controls.Add(this.buttonOuvrirFichiers);
            this.Controls.Add(this.buttonTagDxf);
            this.Controls.Add(this.btnBrowseSe);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.filterPanel);
            this.Controls.Add(this.buttonKillSe);
            this.Controls.Add(this.button10);
            this.Controls.Add(this.textBoxFolderPath);
            this.Controls.Add(this.listBoxDxfFiles);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "PanelSE";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            ((System.ComponentModel.ISupportInitialize)(this.picBoxArrow)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void InitializeFilterControls()
        {
            foreach (var ext in extensionFilters.Keys)
            {
                var cb = new CheckBox
                {
                    Text = ext,
                    Checked = true,
                    ForeColor = System.Drawing.Color.White,
                    AutoSize = true,
                    Margin = new Padding(5, 3, 5, 3)
                };
                cb.CheckedChanged += FilterCheckBox_CheckedChanged;
                cb.MouseUp += FilterCheckBox_MouseUp; // Add MouseUp event
                filterPanel.Controls.Add(cb);
            }
        }

        private void FilterCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            var cb = sender as CheckBox;
            extensionFilters[cb.Text] = cb.Checked;
            RefreshFileList();
        }

        private void FilterCheckBox_MouseUp(object sender, MouseEventArgs e)
        {
            var cb = sender as CheckBox;
            if (e.Button == MouseButtons.Right) // Detect right-click
            {
                bool anyChecked = filterPanel.Controls.OfType<CheckBox>().Any(c => c.Checked);

                if (!anyChecked || filterPanel.Controls.OfType<CheckBox>().Count(c => c.Checked) == 1 && cb.Checked)
                {
                    // If no checkboxes are checked or only the clicked checkbox is checked, check all checkboxes
                    foreach (Control control in filterPanel.Controls)
                    {
                        if (control is CheckBox otherCb)
                        {
                            otherCb.Checked = true;
                        }
                    }
                }
                else
                {
                    // Uncheck all other checkboxes, but keep the clicked one checked
                    foreach (Control control in filterPanel.Controls)
                    {
                        if (control is CheckBox otherCb && otherCb != cb)
                        {
                            otherCb.Checked = false;
                        }
                    }
                    cb.Checked = true; // Keep the clicked checkbox checked
                }
            }
        }

        private void RefreshFileList()
        {
            if (string.IsNullOrEmpty(textBoxFolderPath.Text)) return;

            var activeExtensions = extensionFilters
                .Where(kv => kv.Value)
                .Select(kv => kv.Key)
                .ToList();

            listBoxDxfFiles.Items.Clear();
            string[] allFiles = Directory.GetFiles(textBoxFolderPath.Text, "*.*")
                .Where(file => activeExtensions.Any(ext =>
                    file.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                .ToArray();

            Array.Sort(allFiles, FileSorter.CompareFileNames);
            foreach (string file in allFiles)
            {
                listBoxDxfFiles.Items.Add(Path.GetFileName(file));
            }
        }



        private void button10_Click(object sender, EventArgs e)
        {
            var SelectAllCommand = new SelectAllCommand(textBoxFolderPath, listBoxDxfFiles);
            SelectAllCommand.Execute();
        }

        private void buttonKillSe_Click(object sender, EventArgs e)
        {
            var KillSECommand = new KillSECommand(textBoxFolderPath, listBoxDxfFiles);
            KillSECommand.Execute();
        }

        internal void InitializeSettings(PanelSettings pnlSettings)
        {
            //recuperer tt les methodes importantes
            _panelSettings = pnlSettings;
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            if (_panelSettings != null)
            {
                var mainForm = this.ParentForm as MainForm;
                mainForm?.OpenChildForm(() => _panelSettings);
            }
        }

        private void btnBrowseSe_Click(object sender, EventArgs e)
        {
            var browseCommand = new BrowseCommand(listBoxDxfFiles, textBoxFolderPath, extensionFilters);
            browseCommand.Execute();
            if (buttonExportDim.Visible != true)
            {
                txtBoxUnlock.Visible = false;
                picBoxArrow.Visible = false;
                buttonExportDim.Visible = true;
                buttonGenererDFT.Visible = true;
                buttonOuvrirFichiers.Visible = true;
                buttonSaveDxfStep.Visible = true;
                buttonTagDxf.Visible = true;
            }
        }

        private void buttonTagDxf_Click(object sender, EventArgs e)
        {
            var processDxfCommand = new ProcessDxfCommand(textBoxFolderPath, listBoxDxfFiles, _panelSettings);
            DialogResult dialogResult = MessageBox.Show("Voulez-vous Vraiment taguer les fichiers selectionnes \n" +
                "Attention, Cette action est irreversible", "Confirmation", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes) { processDxfCommand.Execute(); }
        }

        private void buttonOuvrirFichiers_Click(object sender, EventArgs e)
        {
            var openSelectedFilesCommand = new OpenSelectedFilesCommand(listBoxDxfFiles, textBoxFolderPath);
            openSelectedFilesCommand.Execute();
        }

        private void buttonExportDim_Click(object sender, EventArgs e)
        {
            var exportDim = new ExtracteurDimDepliCommand(listBoxDxfFiles, textBoxFolderPath, _panelSettings);
            exportDim.Execute();
        }

        private void buttonSaveDxfStep_Click(object sender, EventArgs e)
        {
            var openStepCmd = new SaveDxfStepCommand(listBoxDxfFiles, textBoxFolderPath);
            openStepCmd.Execute();
        }

        private void buttonGenererDFT_Click(object sender, EventArgs e)
        {
            var createDft = new CreateDftCommand(textBoxFolderPath, listBoxDxfFiles, _panelSettings);
            createDft.Execute();
        }

        private async void btnTaguerDxf_MouseEnter(object sender, EventArgs e)
        {
            // Cancel any existing delay if the mouse enters again before the previous delay is complete
            cancelTokenTag?.Cancel();
            cancelTokenTag = new CancellationTokenSource();
            var token = cancelTokenTag.Token;

            // Show the custom tooltip
            string title = "Utilisation Fonction Tagger Dxf";
            string overlayText = "Ce Boutton va Ouvrir les fichiers dxf, poser le nom de la pièce et sauvegarder.\r" +
            "\nVoir Paramètres si vous ne souhaitez pas garder les documents\r\nmodifiés ouverts" +
            " pour verifier ";
            Image tooltipImage = Resources.gifTagDxf;

            try
            {
                await Task.Delay(1000, token); // 1000 milliseconds = 1 second
                customTooltipTag = new CustomTooltipForm(title, overlayText, tooltipImage)
                {
                    Location = Cursor.Position
                };
                customTooltipTag.Show();
            }
            catch (TaskCanceledException)
            {
                // The delay was canceled, so the form doesn't show
            }
        }

        private void btnTaguerDxf_MouseLeave(object sender, EventArgs e)
        {
            // Hide the custom tooltip for Taguer Dxf
            customTooltipTag?.Close();
            customTooltipTag = null;
            cancelTokenTag?.Cancel();
        }

        private async void btnExporterDim_MouseEnter(object sender, EventArgs e)
        {
            // Cancel any existing delay if the mouse enters again before the previous delay is complete
            cancelTokenDim?.Cancel();
            cancelTokenDim = new CancellationTokenSource();
            var token = cancelTokenDim.Token;

            // Show the custom tooltip for Exporter Dimensions
            string title = "Utilisation Fonction Exporter Dimensions";
            string overlayText = "Ce Boutton va récupérer les dimensions du déplié des pièces séléctionnées " +
                "et va ensuite les placer dans un fichier excel\n Si une des pièces choisies n'est pas dépliée elle sera identifiée " +
                "et ses valeurs dans excel seront 0";
            Image tooltipImage = Resources.gifDimension;

            try
            {
                await Task.Delay(1000, token); // 1000 milliseconds = 1 second
                customTooltipDimensions = new CustomTooltipForm(title, overlayText, tooltipImage)
                {
                    Location = Cursor.Position
                };
                customTooltipDimensions.Show();
            }
            catch (TaskCanceledException)
            {
                // The delay was canceled, so the form doesn't show
            }
        }

        private void btnExporterDim_MouseLeave(object sender, EventArgs e)
        {
            // Hide the custom tooltip for Exporter Dimensions
            customTooltipDimensions?.Close();
            customTooltipDimensions = null;
            cancelTokenDim?.Cancel();
        }

        private async void btnGenererDft_MouseEnter(object sender, EventArgs e)
        {
            // Cancel any existing delay if the mouse enters again before the previous delay is complete
            cancelTokenDft?.Cancel();
            cancelTokenDft = new CancellationTokenSource();
            var token = cancelTokenDft.Token;

            // Show the custom tooltip for Exporter Dimensions
            string title = "Utilisation Fonction Générer Dessins";
            string overlayText = "Ce Boutton va créer un fichier DFT et va ensuite créer un onglet pour chaque pièce et" +
                "assemblage chosi \n" +
                "Pour les assemblages, une page indiquant toutes les composantes de l'assemlage permet de choisir" +
                "les composantes que vous souhaitez faire un dessin \n " +
                "Voir Paramètres si vous ne voulez pas de partsList pour les pièces individuelles";
            Image tooltipImage = Resources.gifDft;

            try
            {
                await Task.Delay(1000, token); // 1000 milliseconds = 1 second
                customTooltipDft = new CustomTooltipForm(title, overlayText, tooltipImage)
                {
                    Location = Cursor.Position
                }; 
                customTooltipDft.Show();
            }
            catch (TaskCanceledException)
            {
                // The delay was canceled, so the form doesn't show
            }
        }

        private void btnGenererDft_MouseLeave(object sender, EventArgs e)
        {
            // Hide the custom tooltip for Exporter Dimensions
            customTooltipDft?.Close();
            customTooltipDft = null;
            cancelTokenDft?.Cancel();
        }
    }   
}