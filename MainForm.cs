using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application_Cyrell.LogiqueBouttonsSolidEdge;
using firstCSMacro;

namespace Application_Cyrell
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

    public partial class MainForm : Form
    {
        private Panel panelMenu;
        private Panel panelLOGO;
        private Panel panelSubXlQc;
        private Panel PnlNav;
        private Panel panelSubSolideEdge;
        private Panel panelContainer;
        private Panel panelBarreMenu;
        private PictureBox pictureBox1;
        private Button buttonAcceuil;
        private Button buttonSolidEdge;
        private Button buttonExcelQc;
        private Button btnClose;
        private Button btnReduce;
        private Button button8;
        private Button button5;
        private Button button4;
        public PanelSE pnlSe;
        public PanelXlQc pnlXlQc;
        public PanelSettings pnlSettings;
        private Label labelTimeDate;
        private System.Windows.Forms.Timer timer;
        private CancellationTokenSource cancelTokenTag;
        private CustomTooltipForm customTooltipTag;
        private CancellationTokenSource cancelTokenDim;
        private CustomTooltipForm customTooltipDimensions;
        private CancellationTokenSource cancelTokenDft;
        private Label labelAcceuil2;
        private Label labelAcceuil1;
        private CustomTooltipForm customTooltipDft;

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
            );

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);


        public MainForm()
        {
            InitializeComponent();
            Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 25, 25));
            PnlNav.Height = buttonAcceuil.Height;
            PnlNav.Top = buttonAcceuil.Top;
            PnlNav.Left = buttonAcceuil.Left;
            buttonAcceuil.BackColor = Color.FromArgb(46, 51, 73);

            InitializationTimerAcceuil();

            MouseDown += MainForm_MouseDown;
            customizeDesign();
            pnlSettings = new PanelSettings();
            pnlSe = new PanelSE();
            pnlXlQc = new PanelXlQc();
            pnlSettings.InitializeParent(pnlSe);
            pnlSe.InitializeSettings(pnlSettings);
        }

        private void InitializationTimerAcceuil()
        {
            labelTimeDate = new Label
            {
                Font = new Font("Arial", 36, FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(230)))), ((int)(((byte)(240))))),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };

            // Ajout du Label au Panel
            panelContainer.Controls.Add(labelTimeDate);

            // Ajout du Panel au Form
            this.Controls.Add(panelContainer);

            // Initialisation du Timer
            timer = new System.Windows.Forms.Timer
            {
                Interval = 1000 // Mise à jour chaque seconde
            };
            timer.Tick += Timer_Tick;
            timer.Start();

            // Mise à jour initiale
            UpdateTimeDate();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            UpdateTimeDate();
        }

        private void UpdateTimeDate()
        {
            // Définition de la culture française
            CultureInfo culture = new CultureInfo("fr-FR");

            // Formatage de l'heure et de la date en français
            string formattedDate = DateTime.Now.ToString("dddd dd MMMM yyyy", culture).ToUpper();
            string dateTimeString = DateTime.Now.ToString("HH:mm:ss\n", culture) + formattedDate;


            // Mise à jour du label
            labelTimeDate.Text = dateTimeString;
        }

        private void customizeDesign()
        {
            panelSubSolideEdge.Visible = false;
            panelSubXlQc.Visible = false;
            //panelSubFut2.Visible = false;
            //panelSubFut3.Visible = false;
        }

        private void hideSubMenu()
        {
            if (panelSubSolideEdge.Visible == true)
                panelSubSolideEdge.Visible = false;
            if (panelSubXlQc.Visible == true)
                panelSubXlQc.Visible = false;
            //if (panelSubFut2.Visible == true)
            //    panelSubFut2.Visible = false;
            //if (panelSubFut3.Visible == true)
            //    panelSubFut3.Visible = false;
        }

        private void showSubMenu(Panel subMenu)
        {
            if (subMenu.Visible == false)
            {
                hideSubMenu();
                subMenu.Visible = true;
            }
            else
                subMenu.Visible = false;
        }

        private void MainForm_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.panelMenu = new System.Windows.Forms.Panel();
            this.button8 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.panelSubXlQc = new System.Windows.Forms.Panel();
            this.buttonExcelQc = new System.Windows.Forms.Button();
            this.panelSubSolideEdge = new System.Windows.Forms.Panel();
            this.PnlNav = new System.Windows.Forms.Panel();
            this.buttonSolidEdge = new System.Windows.Forms.Button();
            this.buttonAcceuil = new System.Windows.Forms.Button();
            this.panelLOGO = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panelContainer = new System.Windows.Forms.Panel();
            this.labelAcceuil2 = new System.Windows.Forms.Label();
            this.labelAcceuil1 = new System.Windows.Forms.Label();
            this.panelBarreMenu = new System.Windows.Forms.Panel();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnReduce = new System.Windows.Forms.Button();
            this.panelMenu.SuspendLayout();
            this.panelLOGO.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panelContainer.SuspendLayout();
            this.panelBarreMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelMenu
            // 
            this.panelMenu.AutoScroll = true;
            this.panelMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.panelMenu.Controls.Add(this.button8);
            this.panelMenu.Controls.Add(this.button5);
            this.panelMenu.Controls.Add(this.button4);
            this.panelMenu.Controls.Add(this.panelSubXlQc);
            this.panelMenu.Controls.Add(this.buttonExcelQc);
            this.panelMenu.Controls.Add(this.panelSubSolideEdge);
            this.panelMenu.Controls.Add(this.PnlNav);
            this.panelMenu.Controls.Add(this.buttonSolidEdge);
            this.panelMenu.Controls.Add(this.buttonAcceuil);
            this.panelMenu.Controls.Add(this.panelLOGO);
            this.panelMenu.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelMenu.Location = new System.Drawing.Point(0, 0);
            this.panelMenu.Name = "panelMenu";
            this.panelMenu.Size = new System.Drawing.Size(266, 856);
            this.panelMenu.TabIndex = 0;
            // 
            // button8
            // 
            this.button8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.button8.Dock = System.Windows.Forms.DockStyle.Top;
            this.button8.FlatAppearance.BorderSize = 0;
            this.button8.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button8.Font = new System.Drawing.Font("Nirmala UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button8.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.button8.Location = new System.Drawing.Point(0, 672);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(266, 52);
            this.button8.TabIndex = 16;
            this.button8.Text = "PAGE FUTURE";
            this.button8.UseVisualStyleBackColor = false;
            this.button8.Visible = false;
            // 
            // button5
            // 
            this.button5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.button5.Dock = System.Windows.Forms.DockStyle.Top;
            this.button5.FlatAppearance.BorderSize = 0;
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button5.Font = new System.Drawing.Font("Nirmala UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.button5.Location = new System.Drawing.Point(0, 620);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(266, 52);
            this.button5.TabIndex = 15;
            this.button5.Text = "PAGE FUTURE";
            this.button5.UseVisualStyleBackColor = false;
            this.button5.Visible = false;
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.button4.Dock = System.Windows.Forms.DockStyle.Top;
            this.button4.FlatAppearance.BorderSize = 0;
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button4.Font = new System.Drawing.Font("Nirmala UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.button4.Location = new System.Drawing.Point(0, 568);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(266, 52);
            this.button4.TabIndex = 14;
            this.button4.Text = "PAGE FUTURE";
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Visible = false;
            // 
            // panelSubXlQc
            // 
            this.panelSubXlQc.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(32)))), ((int)(((byte)(39)))));
            this.panelSubXlQc.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelSubXlQc.Location = new System.Drawing.Point(0, 490);
            this.panelSubXlQc.Name = "panelSubXlQc";
            this.panelSubXlQc.Size = new System.Drawing.Size(266, 78);
            this.panelSubXlQc.TabIndex = 12;
            this.panelSubXlQc.Visible = false;
            // 
            // buttonExcelQc
            // 
            this.buttonExcelQc.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonExcelQc.Dock = System.Windows.Forms.DockStyle.Top;
            this.buttonExcelQc.FlatAppearance.BorderSize = 0;
            this.buttonExcelQc.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonExcelQc.Font = new System.Drawing.Font("Nirmala UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonExcelQc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.buttonExcelQc.Location = new System.Drawing.Point(0, 438);
            this.buttonExcelQc.Name = "buttonExcelQc";
            this.buttonExcelQc.Size = new System.Drawing.Size(266, 52);
            this.buttonExcelQc.TabIndex = 8;
            this.buttonExcelQc.Text = "Excel QC";
            this.buttonExcelQc.UseVisualStyleBackColor = false;
            this.buttonExcelQc.Click += new System.EventHandler(this.buttonExcelQc_Click);
            this.buttonExcelQc.Leave += new System.EventHandler(this.buttonExcelQc_Leave);
            // 
            // panelSubSolideEdge
            // 
            this.panelSubSolideEdge.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(32)))), ((int)(((byte)(39)))));
            this.panelSubSolideEdge.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelSubSolideEdge.Enabled = false;
            this.panelSubSolideEdge.Location = new System.Drawing.Point(0, 248);
            this.panelSubSolideEdge.Name = "panelSubSolideEdge";
            this.panelSubSolideEdge.Size = new System.Drawing.Size(266, 190);
            this.panelSubSolideEdge.TabIndex = 7;
            this.panelSubSolideEdge.Visible = false;
            // 
            // PnlNav
            // 
            this.PnlNav.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.PnlNav.Location = new System.Drawing.Point(0, 0);
            this.PnlNav.Name = "PnlNav";
            this.PnlNav.Size = new System.Drawing.Size(3, 100);
            this.PnlNav.TabIndex = 1;
            // 
            // buttonSolidEdge
            // 
            this.buttonSolidEdge.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonSolidEdge.Dock = System.Windows.Forms.DockStyle.Top;
            this.buttonSolidEdge.FlatAppearance.BorderSize = 0;
            this.buttonSolidEdge.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSolidEdge.Font = new System.Drawing.Font("Nirmala UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSolidEdge.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.buttonSolidEdge.Location = new System.Drawing.Point(0, 196);
            this.buttonSolidEdge.Name = "buttonSolidEdge";
            this.buttonSolidEdge.Size = new System.Drawing.Size(266, 52);
            this.buttonSolidEdge.TabIndex = 2;
            this.buttonSolidEdge.Text = "Solid Edge";
            this.buttonSolidEdge.UseVisualStyleBackColor = false;
            this.buttonSolidEdge.Click += new System.EventHandler(this.btnSE_Click);
            this.buttonSolidEdge.Leave += new System.EventHandler(this.btnSE_Leave);
            // 
            // buttonAcceuil
            // 
            this.buttonAcceuil.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonAcceuil.Dock = System.Windows.Forms.DockStyle.Top;
            this.buttonAcceuil.FlatAppearance.BorderSize = 0;
            this.buttonAcceuil.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonAcceuil.Font = new System.Drawing.Font("Nirmala UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonAcceuil.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.buttonAcceuil.Location = new System.Drawing.Point(0, 144);
            this.buttonAcceuil.Name = "buttonAcceuil";
            this.buttonAcceuil.Size = new System.Drawing.Size(266, 52);
            this.buttonAcceuil.TabIndex = 1;
            this.buttonAcceuil.Text = "Acceuil";
            this.buttonAcceuil.UseVisualStyleBackColor = false;
            this.buttonAcceuil.Click += new System.EventHandler(this.btnAcceuil_Click);
            this.buttonAcceuil.Leave += new System.EventHandler(this.btnAcceuil_Leave);
            // 
            // panelLOGO
            // 
            this.panelLOGO.Controls.Add(this.pictureBox1);
            this.panelLOGO.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelLOGO.Location = new System.Drawing.Point(0, 0);
            this.panelLOGO.Name = "panelLOGO";
            this.panelLOGO.Size = new System.Drawing.Size(266, 144);
            this.panelLOGO.TabIndex = 0;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Application_Cyrell.Properties.Resources.logoCyrellBlanc;
            this.pictureBox1.Location = new System.Drawing.Point(3, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(263, 141);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // panelContainer
            // 
            this.panelContainer.BackColor = System.Drawing.Color.Transparent;
            this.panelContainer.Controls.Add(this.labelAcceuil2);
            this.panelContainer.Controls.Add(this.labelAcceuil1);
            this.panelContainer.Location = new System.Drawing.Point(266, 61);
            this.panelContainer.Name = "panelContainer";
            this.panelContainer.Size = new System.Drawing.Size(1123, 801);
            this.panelContainer.TabIndex = 1;
            // 
            // labelAcceuil2
            // 
            this.labelAcceuil2.AutoSize = true;
            this.labelAcceuil2.Font = new System.Drawing.Font("Arial Rounded MT Bold", 36F);
            this.labelAcceuil2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.labelAcceuil2.Location = new System.Drawing.Point(162, 156);
            this.labelAcceuil2.Name = "labelAcceuil2";
            this.labelAcceuil2.Size = new System.Drawing.Size(1250, 83);
            this.labelAcceuil2.TabIndex = 3;
            this.labelAcceuil2.Text = "Veuillez choisir un onlget à gauche";
            // 
            // labelAcceuil1
            // 
            this.labelAcceuil1.AutoSize = true;
            this.labelAcceuil1.Font = new System.Drawing.Font("Arial Rounded MT Bold", 36F);
            this.labelAcceuil1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.labelAcceuil1.Location = new System.Drawing.Point(94, 101);
            this.labelAcceuil1.Name = "labelAcceuil1";
            this.labelAcceuil1.Size = new System.Drawing.Size(1440, 83);
            this.labelAcceuil1.TabIndex = 2;
            this.labelAcceuil1.Text = "Bienvenue dans l\'application Cyrell AMP";
            // 
            // panelBarreMenu
            // 
            this.panelBarreMenu.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.panelBarreMenu.AutoSize = true;
            this.panelBarreMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.panelBarreMenu.Controls.Add(this.btnClose);
            this.panelBarreMenu.Controls.Add(this.btnReduce);
            this.panelBarreMenu.Location = new System.Drawing.Point(257, 0);
            this.panelBarreMenu.Name = "panelBarreMenu";
            this.panelBarreMenu.Size = new System.Drawing.Size(1132, 64);
            this.panelBarreMenu.TabIndex = 6;
            this.panelBarreMenu.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MainForm_MouseDown);
            // 
            // btnClose
            // 
            this.btnClose.FlatAppearance.BorderSize = 0;
            this.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClose.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.btnClose.Location = new System.Drawing.Point(1068, 0);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(61, 61);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "X";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // btnReduce
            // 
            this.btnReduce.FlatAppearance.BorderSize = 0;
            this.btnReduce.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnReduce.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnReduce.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.btnReduce.Location = new System.Drawing.Point(1007, 0);
            this.btnReduce.Name = "btnReduce";
            this.btnReduce.Size = new System.Drawing.Size(61, 61);
            this.btnReduce.TabIndex = 5;
            this.btnReduce.Text = "-";
            this.btnReduce.UseVisualStyleBackColor = true;
            this.btnReduce.Click += new System.EventHandler(this.button9_Click);
            // 
            // MainForm
            // 
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.ClientSize = new System.Drawing.Size(1389, 856);
            this.Controls.Add(this.panelBarreMenu);
            this.Controls.Add(this.panelContainer);
            this.Controls.Add(this.panelMenu);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.panelMenu.ResumeLayout(false);
            this.panelLOGO.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panelContainer.ResumeLayout(false);
            this.panelContainer.PerformLayout();
            this.panelBarreMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void btnAcceuil_Click(object sender, EventArgs e)
        {
            PnlNav.Height = buttonAcceuil.Height;
            PnlNav.Top = buttonAcceuil.Top;
            PnlNav.Left = buttonAcceuil.Left;
            buttonAcceuil.BackColor = Color.FromArgb(46, 51, 73);

            panelContainer.BringToFront();
            pnlSe.Hide();
            pnlXlQc.Hide();
            pnlSettings.Hide();
            hideSubMenu();
        }

        private void btnAcceuil_Leave(object sender, EventArgs e)
        {
            buttonAcceuil.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnSE_Click(object sender, EventArgs e)
        {
            PnlNav.Height = buttonSolidEdge.Height;
            PnlNav.Top = buttonSolidEdge.Top;
            buttonSolidEdge.BackColor = Color.FromArgb(46, 51, 73);
            PnlNav.BringToFront();
            pnlSettings.Hide();

            openChildForm(() => pnlSe);
            hideSubMenu();
        }

        private void btnSE_Leave(object sender, EventArgs e)
        {
            buttonSolidEdge.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            var cancelCommand = new CancelCommand(this);
            cancelCommand.Execute();
        }


        private void buttonExcelQc_Click(object sender, EventArgs e)
        {
            //pnlSe.Hide();
            //pnlSettings.Hide();
            //panelContainer.BringToFront();
            PnlNav.Height = buttonExcelQc.Height;
            PnlNav.Top = buttonExcelQc.Top;
            buttonExcelQc.BackColor = Color.FromArgb(46, 51, 73);
            PnlNav.BringToFront();

            openChildForm(() => pnlXlQc);
            hideSubMenu();
        }

        private void buttonExcelQc_Leave(object sender, EventArgs e)
        {
            buttonExcelQc.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private Form activeForm = null;
        private Dictionary<Type, Form> formCache = new Dictionary<Type, Form>();

        private void openChildForm<T>(Func<T> formFactory) where T : Form
        {
            // Hide the currently active form
            if (activeForm != null)
            {
                activeForm.Hide();
            }

            // Check if the form is already in the cache
            if (!formCache.TryGetValue(typeof(T), out var childForm))
            {
                // Create a new instance of the form and add it to the cache
                childForm = formFactory();
                formCache[typeof(T)] = childForm;

                childForm.TopLevel = false;
                childForm.FormBorderStyle = FormBorderStyle.None;
                childForm.Dock = DockStyle.Fill;
                panelContainer.Controls.Add(childForm);
                panelContainer.Tag = childForm;
            }

            // Show the selected form
            activeForm = childForm;
            childForm.BringToFront();
            childForm.Show();
        }

        public void OpenChildForm<T>(Func<T> formFactory) where T : Form
        {
            openChildForm(formFactory);
        }
    }
}