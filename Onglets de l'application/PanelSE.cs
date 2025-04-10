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
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Control = System.Windows.Forms.Control;
using System.Diagnostics;

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
        private Button buttonSelectAll;
        public ListBox listBoxDxfFiles;
        private Button buttonKillSe;
        private FlowLayoutPanel filterPanel;
        private Panel sidebarPanel;
        private PanelSettings _panelSettings;
        public bool paramFermerSe;
        private Button btnSettings;
        private Button btnBrowseSe;
        private Button buttonTagDxf;
        private Button buttonOuvrirFichiers;
        private Button buttonExportDim;
        private Button buttonFlatPatterns;
        private Button buttonSaveDxfStep;
        private Button buttonGenererDFT;
        private CancellationTokenSource cancelTokenTag;
        private CustomTooltipForm customTooltipTag;
        private CancellationTokenSource cancelTokenDim;
        private CustomTooltipForm customTooltipDimensions;
        private CancellationTokenSource cancelTokenDft;
        private CustomTooltipForm customTooltipDft;
        private PictureBox picBoxArrow;
        private Label labelUnlock;
        private Label labelSelectedFiles;
        private Label labelSelectedFilesCount;
        private Button themeSwitchButton;
        private Dictionary<string, bool> extensionFilters = new Dictionary<string, bool>()
        {
            { ".asm", true },
            { ".dxf", true },
            { ".stp", true },
            { ".step", true },
            { ".par", true },
            { ".psm", true },
            { ".SLDASM", true }
        };

        public PanelSE()
        {
            InitializeComponent();
            InitializeFilterControls();
            listBoxDxfFiles.SelectedIndexChanged += ListBoxDxfFiles_SelectedIndexChanged;
        }

        private void InitializeComponent()
        {
            this.listBoxDxfFiles = new System.Windows.Forms.ListBox();
            this.textBoxFolderPath = new System.Windows.Forms.TextBox();
            this.buttonSelectAll = new System.Windows.Forms.Button();
            this.buttonKillSe = new System.Windows.Forms.Button();
            this.filterPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.sidebarPanel = new System.Windows.Forms.Panel();
            this.buttonTagDxf = new System.Windows.Forms.Button();
            this.buttonOuvrirFichiers = new System.Windows.Forms.Button();
            this.buttonExportDim = new System.Windows.Forms.Button();
            this.buttonFlatPatterns = new System.Windows.Forms.Button();
            this.buttonSaveDxfStep = new System.Windows.Forms.Button();
            this.buttonGenererDFT = new System.Windows.Forms.Button();
            this.labelUnlock = new System.Windows.Forms.Label();
            this.picBoxArrow = new System.Windows.Forms.PictureBox();
            this.btnBrowseSe = new System.Windows.Forms.Button();
            this.btnSettings = new System.Windows.Forms.Button();
            this.labelSelectedFiles = new System.Windows.Forms.Label();
            this.labelSelectedFilesCount = new System.Windows.Forms.Label();
            this.themeSwitchButton = new System.Windows.Forms.Button();
            this.sidebarPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxArrow)).BeginInit();
            this.SuspendLayout();
            // 
            // listBoxDxfFiles
            // 
            this.listBoxDxfFiles.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(38)))), ((int)(((byte)(70)))));
            this.listBoxDxfFiles.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.listBoxDxfFiles.Font = new System.Drawing.Font("Tahoma", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBoxDxfFiles.ForeColor = System.Drawing.Color.White;
            this.listBoxDxfFiles.FormattingEnabled = true;
            this.listBoxDxfFiles.ItemHeight = 34;
            this.listBoxDxfFiles.Location = new System.Drawing.Point(250, 184);
            this.listBoxDxfFiles.Name = "listBoxDxfFiles";
            this.listBoxDxfFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBoxDxfFiles.Size = new System.Drawing.Size(838, 374);
            this.listBoxDxfFiles.TabIndex = 4;
            // 
            // textBoxFolderPath
            // 
            this.textBoxFolderPath.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(38)))), ((int)(((byte)(70)))));
            this.textBoxFolderPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxFolderPath.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxFolderPath.ForeColor = System.Drawing.Color.White;
            this.textBoxFolderPath.Location = new System.Drawing.Point(250, 134);
            this.textBoxFolderPath.Name = "textBoxFolderPath";
            this.textBoxFolderPath.Size = new System.Drawing.Size(782, 39);
            this.textBoxFolderPath.TabIndex = 5;
            // 
            // buttonSelectAll
            // 
            this.buttonSelectAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(73)))), ((int)(((byte)(94)))));
            this.buttonSelectAll.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonSelectAll.FlatAppearance.BorderSize = 0;
            this.buttonSelectAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSelectAll.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSelectAll.ForeColor = System.Drawing.Color.White;
            this.buttonSelectAll.Location = new System.Drawing.Point(697, 700);
            this.buttonSelectAll.Name = "buttonSelectAll";
            this.buttonSelectAll.Size = new System.Drawing.Size(140, 45);
            this.buttonSelectAll.TabIndex = 9;
            this.buttonSelectAll.Text = "Select All";
            this.buttonSelectAll.UseVisualStyleBackColor = false;
            this.buttonSelectAll.Click += new System.EventHandler(this.button10_Click);
            // 
            // buttonKillSe
            // 
            this.buttonKillSe.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(231)))), ((int)(((byte)(76)))), ((int)(((byte)(60)))));
            this.buttonKillSe.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonKillSe.FlatAppearance.BorderSize = 0;
            this.buttonKillSe.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonKillSe.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.buttonKillSe.ForeColor = System.Drawing.Color.White;
            this.buttonKillSe.Location = new System.Drawing.Point(20, 674);
            this.buttonKillSe.Name = "buttonKillSe";
            this.buttonKillSe.Size = new System.Drawing.Size(180, 45);
            this.buttonKillSe.TabIndex = 7;
            this.buttonKillSe.Text = "Fermer Solid Edge";
            this.buttonKillSe.UseVisualStyleBackColor = false;
            this.buttonKillSe.Click += new System.EventHandler(this.buttonKillSe_Click);
            // 
            // filterPanel
            // 
            this.filterPanel.BackColor = System.Drawing.Color.Transparent;
            this.filterPanel.Location = new System.Drawing.Point(250, 706);
            this.filterPanel.Name = "filterPanel";
            this.filterPanel.Size = new System.Drawing.Size(420, 30);
            this.filterPanel.TabIndex = 0;
            // 
            // sidebarPanel
            // 
            this.sidebarPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(25)))), ((int)(((byte)(33)))), ((int)(((byte)(60)))));
            this.sidebarPanel.Controls.Add(this.buttonTagDxf);
            this.sidebarPanel.Controls.Add(this.buttonOuvrirFichiers);
            this.sidebarPanel.Controls.Add(this.buttonExportDim);
            this.sidebarPanel.Controls.Add(this.buttonFlatPatterns);
            this.sidebarPanel.Controls.Add(this.buttonSaveDxfStep);
            this.sidebarPanel.Controls.Add(this.buttonGenererDFT);
            this.sidebarPanel.Controls.Add(this.buttonKillSe);
            this.sidebarPanel.Location = new System.Drawing.Point(0, 0);
            this.sidebarPanel.Name = "sidebarPanel";
            this.sidebarPanel.Size = new System.Drawing.Size(220, 800);
            this.sidebarPanel.TabIndex = 27;
            // 
            // buttonTagDxf
            // 
            this.buttonTagDxf.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.buttonTagDxf.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonTagDxf.FlatAppearance.BorderSize = 0;
            this.buttonTagDxf.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonTagDxf.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.buttonTagDxf.ForeColor = System.Drawing.Color.White;
            this.buttonTagDxf.Location = new System.Drawing.Point(20, 247);
            this.buttonTagDxf.Name = "buttonTagDxf";
            this.buttonTagDxf.Size = new System.Drawing.Size(180, 45);
            this.buttonTagDxf.TabIndex = 2;
            this.buttonTagDxf.Text = "Taguer DXF";
            this.buttonTagDxf.UseVisualStyleBackColor = false;
            this.buttonTagDxf.Click += new System.EventHandler(this.buttonTagDxf_Click);
            this.buttonTagDxf.MouseEnter += new System.EventHandler(this.btnTaguerDxf_MouseEnter);
            this.buttonTagDxf.MouseLeave += new System.EventHandler(this.btnTaguerDxf_MouseLeave);
            // 
            // buttonOuvrirFichiers
            // 
            this.buttonOuvrirFichiers.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.buttonOuvrirFichiers.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonOuvrirFichiers.FlatAppearance.BorderSize = 0;
            this.buttonOuvrirFichiers.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOuvrirFichiers.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.buttonOuvrirFichiers.ForeColor = System.Drawing.Color.White;
            this.buttonOuvrirFichiers.Location = new System.Drawing.Point(20, 134);
            this.buttonOuvrirFichiers.Name = "buttonOuvrirFichiers";
            this.buttonOuvrirFichiers.Size = new System.Drawing.Size(180, 45);
            this.buttonOuvrirFichiers.TabIndex = 1;
            this.buttonOuvrirFichiers.Text = "Ouvrir Fichiers";
            this.buttonOuvrirFichiers.UseVisualStyleBackColor = false;
            this.buttonOuvrirFichiers.Click += new System.EventHandler(this.buttonOuvrirFichiers_Click);
            // 
            // buttonExportDim
            // 
            this.buttonExportDim.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.buttonExportDim.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonExportDim.FlatAppearance.BorderSize = 0;
            this.buttonExportDim.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonExportDim.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.buttonExportDim.ForeColor = System.Drawing.Color.White;
            this.buttonExportDim.Location = new System.Drawing.Point(20, 308);
            this.buttonExportDim.Name = "buttonExportDim";
            this.buttonExportDim.Size = new System.Drawing.Size(180, 45);
            this.buttonExportDim.TabIndex = 3;
            this.buttonExportDim.Text = "Exporter Dimensions";
            this.buttonExportDim.UseVisualStyleBackColor = false;
            this.buttonExportDim.Click += new System.EventHandler(this.buttonExportDim_Click);
            this.buttonExportDim.MouseEnter += new System.EventHandler(this.btnExporterDim_MouseEnter);
            this.buttonExportDim.MouseLeave += new System.EventHandler(this.btnExporterDim_MouseLeave);
            // 
            // buttonFlatPatterns
            // 
            this.buttonFlatPatterns.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.buttonFlatPatterns.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonFlatPatterns.FlatAppearance.BorderSize = 0;
            this.buttonFlatPatterns.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonFlatPatterns.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.buttonFlatPatterns.ForeColor = System.Drawing.Color.White;
            this.buttonFlatPatterns.Location = new System.Drawing.Point(20, 368);
            this.buttonFlatPatterns.Name = "buttonFlatPatterns";
            this.buttonFlatPatterns.Size = new System.Drawing.Size(180, 45);
            this.buttonFlatPatterns.TabIndex = 4;
            this.buttonFlatPatterns.Text = "Créer Flat Pattern";
            this.buttonFlatPatterns.UseVisualStyleBackColor = false;
            this.buttonFlatPatterns.Click += new System.EventHandler(this.buttonFlatPattern_Click);
            // 
            // buttonSaveDxfStep
            // 
            this.buttonSaveDxfStep.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.buttonSaveDxfStep.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonSaveDxfStep.FlatAppearance.BorderSize = 0;
            this.buttonSaveDxfStep.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSaveDxfStep.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.buttonSaveDxfStep.ForeColor = System.Drawing.Color.White;
            this.buttonSaveDxfStep.Location = new System.Drawing.Point(20, 430);
            this.buttonSaveDxfStep.Name = "buttonSaveDxfStep";
            this.buttonSaveDxfStep.Size = new System.Drawing.Size(180, 45);
            this.buttonSaveDxfStep.TabIndex = 5;
            this.buttonSaveDxfStep.Text = "Savegarder DXF && Step";
            this.buttonSaveDxfStep.UseVisualStyleBackColor = false;
            this.buttonSaveDxfStep.Click += new System.EventHandler(this.buttonSaveDxfStep_Click);
            // 
            // buttonGenererDFT
            // 
            this.buttonGenererDFT.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
            this.buttonGenererDFT.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonGenererDFT.FlatAppearance.BorderSize = 0;
            this.buttonGenererDFT.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonGenererDFT.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.buttonGenererDFT.ForeColor = System.Drawing.Color.White;
            this.buttonGenererDFT.Location = new System.Drawing.Point(20, 491);
            this.buttonGenererDFT.Name = "buttonGenererDFT";
            this.buttonGenererDFT.Size = new System.Drawing.Size(180, 45);
            this.buttonGenererDFT.TabIndex = 6;
            this.buttonGenererDFT.Text = "Générer DFT";
            this.buttonGenererDFT.UseVisualStyleBackColor = false;
            this.buttonGenererDFT.Click += new System.EventHandler(this.buttonGenererDFT_Click);
            this.buttonGenererDFT.MouseEnter += new System.EventHandler(this.btnGenererDft_MouseEnter);
            this.buttonGenererDFT.MouseLeave += new System.EventHandler(this.btnGenererDft_MouseLeave);
            // 
            // labelUnlock
            // 
            this.labelUnlock.AutoSize = true;
            this.labelUnlock.Font = new System.Drawing.Font("Segoe UI", 24F, System.Drawing.FontStyle.Bold);
            this.labelUnlock.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(236)))), ((int)(((byte)(240)))), ((int)(((byte)(241)))));
            this.labelUnlock.Location = new System.Drawing.Point(448, 18);
            this.labelUnlock.Name = "labelUnlock";
            this.labelUnlock.Size = new System.Drawing.Size(677, 65);
            this.labelUnlock.TabIndex = 21;
            this.labelUnlock.Text = "Veuillez choisir un répertoire";
            // 
            // picBoxArrow
            // 
            this.picBoxArrow.BackColor = System.Drawing.Color.Transparent;
            this.picBoxArrow.BackgroundImage = global::Application_Cyrell.Properties.Resources.logoArrow;
            this.picBoxArrow.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.picBoxArrow.Location = new System.Drawing.Point(1040, 40);
            this.picBoxArrow.Name = "picBoxArrow";
            this.picBoxArrow.Size = new System.Drawing.Size(55, 55);
            this.picBoxArrow.TabIndex = 20;
            this.picBoxArrow.TabStop = false;
            // 
            // btnBrowseSe
            // 
            this.btnBrowseSe.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnBrowseSe.BackColor = System.Drawing.Color.Transparent;
            this.btnBrowseSe.BackgroundImage = global::Application_Cyrell.Properties.Resources.search_in_folder;
            this.btnBrowseSe.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBrowseSe.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnBrowseSe.FlatAppearance.BorderSize = 0;
            this.btnBrowseSe.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowseSe.Location = new System.Drawing.Point(1047, 120);
            this.btnBrowseSe.Name = "btnBrowseSe";
            this.btnBrowseSe.Size = new System.Drawing.Size(55, 55);
            this.btnBrowseSe.TabIndex = 8;
            this.btnBrowseSe.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnBrowseSe.UseVisualStyleBackColor = false;
            this.btnBrowseSe.Click += new System.EventHandler(this.btnBrowseSe_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.BackColor = System.Drawing.Color.Transparent;
            this.btnSettings.BackgroundImage = global::Application_Cyrell.Properties.Resources.logoParam;
            this.btnSettings.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnSettings.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSettings.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.btnSettings.FlatAppearance.BorderSize = 0;
            this.btnSettings.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.btnSettings.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSettings.Location = new System.Drawing.Point(979, 674);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnSettings.Size = new System.Drawing.Size(93, 85);
            this.btnSettings.TabIndex = 10;
            this.btnSettings.UseVisualStyleBackColor = false;
            this.btnSettings.Click += new System.EventHandler(this.btnSettings_Click);
            // 
            // labelSelectedFiles
            // 
            this.labelSelectedFiles.AutoSize = true;
            this.labelSelectedFiles.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold);
            this.labelSelectedFiles.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(236)))), ((int)(((byte)(240)))), ((int)(((byte)(241)))));
            this.labelSelectedFiles.Location = new System.Drawing.Point(250, 618);
            this.labelSelectedFiles.Name = "labelSelectedFiles";
            this.labelSelectedFiles.Size = new System.Drawing.Size(223, 38);
            this.labelSelectedFiles.TabIndex = 22;
            this.labelSelectedFiles.Text = "Fichiers Choisis:";
            this.labelSelectedFiles.Visible = false;
            // 
            // labelSelectedFilesCount
            // 
            this.labelSelectedFilesCount.AutoSize = true;
            this.labelSelectedFilesCount.BackColor = System.Drawing.Color.Transparent;
            this.labelSelectedFilesCount.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold);
            this.labelSelectedFilesCount.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(204)))), ((int)(((byte)(113)))));
            this.labelSelectedFilesCount.Location = new System.Drawing.Point(400, 618);
            this.labelSelectedFilesCount.Name = "labelSelectedFilesCount";
            this.labelSelectedFilesCount.Size = new System.Drawing.Size(0, 38);
            this.labelSelectedFilesCount.TabIndex = 24;
            this.labelSelectedFilesCount.Visible = false;
            // 
            // themeSwitchButton
            // 
            this.themeSwitchButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.themeSwitchButton.FlatAppearance.BorderSize = 0;
            this.themeSwitchButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.themeSwitchButton.Location = new System.Drawing.Point(1091, 766);
            this.themeSwitchButton.Name = "themeSwitchButton";
            this.themeSwitchButton.Size = new System.Drawing.Size(20, 20);
            this.themeSwitchButton.TabIndex = 11;
            this.themeSwitchButton.UseVisualStyleBackColor = false;
            this.themeSwitchButton.Click += new System.EventHandler(this.themeSwitchButton_Click);
            // 
            // PanelSE
            // 
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.ClientSize = new System.Drawing.Size(1123, 798);
            this.Controls.Add(this.themeSwitchButton);
            this.Controls.Add(this.labelSelectedFilesCount);
            this.Controls.Add(this.labelSelectedFiles);
            this.Controls.Add(this.picBoxArrow);
            this.Controls.Add(this.btnBrowseSe);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.filterPanel);
            this.Controls.Add(this.buttonSelectAll);
            this.Controls.Add(this.textBoxFolderPath);
            this.Controls.Add(this.listBoxDxfFiles);
            this.Controls.Add(this.labelUnlock);
            this.Controls.Add(this.sidebarPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "PanelSE";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.sidebarPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picBoxArrow)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        // Add this method to handle theme switching
        private void themeSwitchButton_Click(object sender, EventArgs e)
        {
            // Toggle between dark and light theme
            if (this.BackColor.R == 37) // Dark mode is active
            {
                // Switch to light mode
                this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
                this.sidebarPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(220)))), ((int)(((byte)(220)))));
                this.listBoxDxfFiles.BackColor = System.Drawing.Color.White;
                this.listBoxDxfFiles.ForeColor = System.Drawing.Color.Black;
                this.textBoxFolderPath.BackColor = System.Drawing.Color.White;
                this.textBoxFolderPath.ForeColor = System.Drawing.Color.Black;
                this.labelUnlock.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(62)))), ((int)(((byte)(80)))));
                this.labelSelectedFiles.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(62)))), ((int)(((byte)(80)))));
                this.themeSwitchButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
                foreach (Control control in filterPanel.Controls)
                {
                    if (control is CheckBox cb)
                    {
                        cb.ForeColor = System.Drawing.Color.Black;
                    }
                }

                // Update button colors for light theme
                foreach (Control control in sidebarPanel.Controls)
                {
                    if (control is Button button && button != buttonKillSe)
                    {
                        button.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(152)))), ((int)(((byte)(219)))));
                    }
                }
            }
            else // Light mode is active
            {
                // Switch to dark mode
                this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
                this.sidebarPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(25)))), ((int)(((byte)(33)))), ((int)(((byte)(60)))));
                this.listBoxDxfFiles.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(38)))), ((int)(((byte)(70)))));
                this.listBoxDxfFiles.ForeColor = System.Drawing.Color.White;
                this.textBoxFolderPath.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(38)))), ((int)(((byte)(70)))));
                this.textBoxFolderPath.ForeColor = System.Drawing.Color.White;
                this.labelUnlock.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(236)))), ((int)(((byte)(240)))), ((int)(((byte)(241)))));
                this.labelSelectedFiles.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(236)))), ((int)(((byte)(240)))), ((int)(((byte)(241)))));
                this.themeSwitchButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
                foreach (Control control in filterPanel.Controls)
                {
                    if (control is CheckBox cb)
                    {
                        cb.ForeColor = System.Drawing.Color.White;
                    }
                }

                // Update button colors for dark theme
                foreach (Control control in sidebarPanel.Controls)
                {
                    if (control is Button button && button != buttonKillSe)
                    {
                        button.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(128)))), ((int)(((byte)(185)))));
                    }
                }
            }
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
            int selectedCount = listBoxDxfFiles.SelectedItems.Count;
            labelSelectedFilesCount.Text = selectedCount.ToString();
        }

        private void ListBoxDxfFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedCount = listBoxDxfFiles.SelectedItems.Count;
            labelSelectedFilesCount.Text = selectedCount.ToString();
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
            if (labelUnlock.Visible)
            {
                labelUnlock.Visible = false;
                picBoxArrow.Visible = false;
                buttonExportDim.Visible = true;
                buttonGenererDFT.Visible = true;
                buttonOuvrirFichiers.Visible = true;
                buttonSaveDxfStep.Visible = true;
                buttonTagDxf.Visible = true;
                buttonFlatPatterns.Visible = true;
                labelSelectedFiles.Visible = true;
                labelSelectedFilesCount.Visible = true;
                labelSelectedFilesCount.Text = "0";
                textBoxFolderPath.Location = (Point)new Size(textBoxFolderPath.Location.X, textBoxFolderPath.Location.Y - 104);
                listBoxDxfFiles.Location = (Point)new Size(listBoxDxfFiles.Location.X, listBoxDxfFiles.Location.Y - 104);
                btnBrowseSe.Location = (Point)new Size(btnBrowseSe.Location.X, btnBrowseSe.Location.Y - 104);
                listBoxDxfFiles.Size = (Size)new Size(listBoxDxfFiles.Size.Width, listBoxDxfFiles.Size.Height + 104);
            }
        }

        private void buttonTagDxf_Click(object sender, EventArgs e)
        {
            var processDxfCommand = new ProcessDxfCommand(textBoxFolderPath, listBoxDxfFiles, _panelSettings);
            processDxfCommand.Execute();
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
            if (listBoxDxfFiles.SelectedItems.Count == 0)
            {
                MessageBox.Show("Veuillez Choisir au moins un fichier PAR, PSM ou ASM à traiter");
                return;
            }

            // Example usage
            FormulaireDFT form = new FormulaireDFT();
            List<bool> parametres = new List<bool>();
            List<double> valNum = new List<double>();

            // Register handlers for both continue buttons
            form.OnContinue1 += (parList, dftInd, isoView, flatView, bendTable, refVars, countParts, bendTableAdv, autoScale, scale, spacingX, spacingY) => {
                parametres.Add(parList);
                parametres.Add(dftInd);
                parametres.Add(isoView);
                parametres.Add(flatView);
                parametres.Add(bendTable);
                parametres.Add(refVars);
                parametres.Add(countParts);
                parametres.Add(autoScale);
                valNum.Add(scale);

                var createDft = new CreateDftCommand(textBoxFolderPath, listBoxDxfFiles, parametres, valNum);
                createDft.Execute();
                parametres.Clear();
                valNum.Clear();
            };

            form.OnContinue2 += (parList, dftInd, isoView, flatView, bendTable, refVars, countParts, bendTableAdv, autoScale, scale, spacingX, spacingY) => {
                parametres.Add(refVars);
                parametres.Add(bendTableAdv);
                parametres.Add(autoScale);
                valNum.Add(scale);
                valNum.Add(spacingX);
                valNum.Add(spacingY);

                var createDft = new CreateFlatDftCommand(textBoxFolderPath, listBoxDxfFiles, parametres, valNum);
                createDft.Execute();
                parametres.Clear();
                valNum.Clear();
            };

            form.ShowDialog();
            
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

        private void buttonFlatPattern_Click(object sender, EventArgs e)
        {
            var createFlatPattern = new ProcessPsmCommand(textBoxFolderPath, listBoxDxfFiles);
            createFlatPattern.Execute();
        }
    }   
}