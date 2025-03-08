using System;
using System.Windows.Forms;
using Application_Cyrell.LogiqueBouttonsExcel;

namespace firstCSMacro
{
    public class PanelXlQc : Form
    {
        private TextBox xlJobPathTxtBox;
        private TextBox dxfPathTxtBox;
        private TextBox stepPathTxtBox;
        private Button btnBrowseXl;
        private Button btnBrowseDxf;
        private Button btnBrowseStep;
        private TextBox textBox1;
        private string xlFilePath;
        private string dxfFilePath;
        private Button buttonVerifDim;
        private Button buttonVerifQte;
        private Label label1;
        private Label label2;
        private Label label3;
        private Button btnRefresh;
        private string stepFilePath;

        public PanelXlQc()
        {
            InitializeComponent();
        }
        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.xlJobPathTxtBox = new System.Windows.Forms.TextBox();
            this.dxfPathTxtBox = new System.Windows.Forms.TextBox();
            this.stepPathTxtBox = new System.Windows.Forms.TextBox();
            this.btnBrowseXl = new System.Windows.Forms.Button();
            this.btnBrowseDxf = new System.Windows.Forms.Button();
            this.btnBrowseStep = new System.Windows.Forms.Button();
            this.buttonVerifDim = new System.Windows.Forms.Button();
            this.buttonVerifQte = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Cursor = System.Windows.Forms.Cursors.Default;
            this.textBox1.Enabled = false;
            this.textBox1.Font = new System.Drawing.Font("Arial", 72F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(66)))), ((int)(((byte)(71)))), ((int)(((byte)(93)))));
            this.textBox1.Location = new System.Drawing.Point(335, 566);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(776, 111);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "EXCEL QC";
            // 
            // xlJobPathTxtBox
            // 
            this.xlJobPathTxtBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.xlJobPathTxtBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.xlJobPathTxtBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.xlJobPathTxtBox.ForeColor = System.Drawing.Color.White;
            this.xlJobPathTxtBox.Location = new System.Drawing.Point(101, 114);
            this.xlJobPathTxtBox.Name = "xlJobPathTxtBox";
            this.xlJobPathTxtBox.Size = new System.Drawing.Size(813, 19);
            this.xlJobPathTxtBox.TabIndex = 6;
            // 
            // dxfPathTxtBox
            // 
            this.dxfPathTxtBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.dxfPathTxtBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dxfPathTxtBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.dxfPathTxtBox.ForeColor = System.Drawing.Color.White;
            this.dxfPathTxtBox.Location = new System.Drawing.Point(101, 230);
            this.dxfPathTxtBox.Name = "dxfPathTxtBox";
            this.dxfPathTxtBox.Size = new System.Drawing.Size(813, 19);
            this.dxfPathTxtBox.TabIndex = 7;
            // 
            // stepPathTxtBox
            // 
            this.stepPathTxtBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.stepPathTxtBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.stepPathTxtBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.stepPathTxtBox.ForeColor = System.Drawing.Color.White;
            this.stepPathTxtBox.Location = new System.Drawing.Point(101, 352);
            this.stepPathTxtBox.Name = "stepPathTxtBox";
            this.stepPathTxtBox.Size = new System.Drawing.Size(813, 19);
            this.stepPathTxtBox.TabIndex = 8;
            // 
            // btnBrowseXl
            // 
            this.btnBrowseXl.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnBrowseXl.BackgroundImage = global::Application_Cyrell.Properties.Resources.search_in_folder;
            this.btnBrowseXl.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBrowseXl.FlatAppearance.BorderSize = 0;
            this.btnBrowseXl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowseXl.Location = new System.Drawing.Point(938, 90);
            this.btnBrowseXl.Name = "btnBrowseXl";
            this.btnBrowseXl.Size = new System.Drawing.Size(65, 65);
            this.btnBrowseXl.TabIndex = 12;
            this.btnBrowseXl.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnBrowseXl.UseVisualStyleBackColor = true;
            this.btnBrowseXl.Click += new System.EventHandler(this.btnBrowseXl_Click);
            // 
            // btnBrowseDxf
            // 
            this.btnBrowseDxf.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnBrowseDxf.BackgroundImage = global::Application_Cyrell.Properties.Resources.search_in_folder;
            this.btnBrowseDxf.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBrowseDxf.FlatAppearance.BorderSize = 0;
            this.btnBrowseDxf.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowseDxf.Location = new System.Drawing.Point(938, 204);
            this.btnBrowseDxf.Name = "btnBrowseDxf";
            this.btnBrowseDxf.Size = new System.Drawing.Size(65, 65);
            this.btnBrowseDxf.TabIndex = 13;
            this.btnBrowseDxf.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnBrowseDxf.UseVisualStyleBackColor = true;
            this.btnBrowseDxf.Click += new System.EventHandler(this.btnBrowseDxf_Click);
            // 
            // btnBrowseStep
            // 
            this.btnBrowseStep.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnBrowseStep.BackgroundImage = global::Application_Cyrell.Properties.Resources.search_in_folder;
            this.btnBrowseStep.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBrowseStep.FlatAppearance.BorderSize = 0;
            this.btnBrowseStep.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowseStep.Location = new System.Drawing.Point(938, 327);
            this.btnBrowseStep.Name = "btnBrowseStep";
            this.btnBrowseStep.Size = new System.Drawing.Size(65, 65);
            this.btnBrowseStep.TabIndex = 14;
            this.btnBrowseStep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnBrowseStep.UseVisualStyleBackColor = true;
            this.btnBrowseStep.Click += new System.EventHandler(this.btnBrowseStep_Click);
            // 
            // buttonVerifDim
            // 
            this.buttonVerifDim.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonVerifDim.FlatAppearance.BorderSize = 0;
            this.buttonVerifDim.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonVerifDim.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonVerifDim.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.buttonVerifDim.Location = new System.Drawing.Point(101, 444);
            this.buttonVerifDim.Name = "buttonVerifDim";
            this.buttonVerifDim.Size = new System.Drawing.Size(234, 40);
            this.buttonVerifDim.TabIndex = 18;
            this.buttonVerifDim.Text = "Verifier Dimension Coupe";
            this.buttonVerifDim.UseVisualStyleBackColor = false;
            this.buttonVerifDim.Click += new System.EventHandler(this.buttonVerifDim_Click);
            // 
            // buttonVerifQte
            // 
            this.buttonVerifQte.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.buttonVerifQte.FlatAppearance.BorderSize = 0;
            this.buttonVerifQte.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonVerifQte.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonVerifQte.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.buttonVerifQte.Location = new System.Drawing.Point(382, 444);
            this.buttonVerifQte.Name = "buttonVerifQte";
            this.buttonVerifQte.Size = new System.Drawing.Size(234, 40);
            this.buttonVerifQte.TabIndex = 19;
            this.buttonVerifQte.Text = "Verifier Quantité Pièces";
            this.buttonVerifQte.UseVisualStyleBackColor = false;
            this.buttonVerifQte.Click += new System.EventHandler(this.buttonVerifQte_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 25F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.label1.Location = new System.Drawing.Point(95, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(497, 40);
            this.label1.TabIndex = 20;
            this.label1.Text = "Emplacement du fichier Excel";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 25F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.label2.Location = new System.Drawing.Point(93, 180);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(605, 40);
            this.label2.TabIndex = 21;
            this.label2.Text = "Emplacement du répértoire des DXF";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 25F, System.Drawing.FontStyle.Bold);
            this.label3.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.label3.Location = new System.Drawing.Point(95, 307);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(610, 40);
            this.label3.TabIndex = 22;
            this.label3.Text = "Emplacement du répértoire des Step";
            // 
            // btnRefresh
            // 
            this.btnRefresh.BackColor = System.Drawing.Color.DarkOrange;
            this.btnRefresh.FlatAppearance.BorderSize = 0;
            this.btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRefresh.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRefresh.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnRefresh.Location = new System.Drawing.Point(916, 430);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(99, 38);
            this.btnRefresh.TabIndex = 23;
            this.btnRefresh.Text = "Rafraîchir";
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // PanelXlQc
            // 
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.ClientSize = new System.Drawing.Size(1123, 798);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonVerifQte);
            this.Controls.Add(this.buttonVerifDim);
            this.Controls.Add(this.btnBrowseStep);
            this.Controls.Add(this.btnBrowseDxf);
            this.Controls.Add(this.btnBrowseXl);
            this.Controls.Add(this.stepPathTxtBox);
            this.Controls.Add(this.dxfPathTxtBox);
            this.Controls.Add(this.xlJobPathTxtBox);
            this.Controls.Add(this.textBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "PanelXlQc";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void btnBrowseXl_Click(object sender, System.EventArgs e)
        {
            var clickParXl = new parcourirXlCommand(xlJobPathTxtBox);
            clickParXl.Execute();
            xlFilePath = xlJobPathTxtBox.Text;
            clickParXl = null;
        }

        private void btnBrowseDxf_Click(object sender, System.EventArgs e)
        {
            var clickParDxf = new parcourirDxfStepCommand(dxfPathTxtBox);
            clickParDxf.Execute();
            dxfFilePath = dxfPathTxtBox.Text;
            clickParDxf = null;
        }

        private void btnBrowseStep_Click(object sender, System.EventArgs e)
        {
            var clickParStep = new parcourirDxfStepCommand(stepPathTxtBox);
            clickParStep.Execute();
            stepFilePath = stepPathTxtBox.Text;
            clickParStep = null;
        }

        private void buttonVerifDim_Click(object sender, EventArgs e)
        {
            var clickVerifDimCoupe = new VerifDimCommand(xlJobPathTxtBox, dxfPathTxtBox);
            clickVerifDimCoupe.Execute();
            clickVerifDimCoupe = null;
        }

        private void buttonVerifQte_Click(object sender, EventArgs e)
        {
            var clickVerifQte = new VerifNbPiecesCommand(xlJobPathTxtBox, dxfPathTxtBox, stepPathTxtBox);
            clickVerifQte.Execute();
            clickVerifQte = null;
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            xlFilePath = "";
            dxfFilePath = "";
            stepFilePath = "";
            xlJobPathTxtBox.Text = "";
            dxfPathTxtBox.Text = "";
            stepPathTxtBox.Text = "";
        }
    }
}