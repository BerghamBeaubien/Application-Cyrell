using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Application_Cyrell;
using Application_Cyrell.Utils;

namespace firstCSMacro
{

    public partial class PanelSettings : Form
    {
        private FlowLayoutPanel flpDxf;
        private TextBox dxfSetting1Txt;
        private Button btnReturn;
        private TextBox titrePanel;
        private TextBox textBox5;
        private BouttonToggle chkBoxDxfTag1;
        private TextBox textBox1;
        private FlowLayoutPanel flpDim;
        private TextBox txtDim1;
        private BouttonToggle chkBoxDim1;
        private TextBox txtDim2;
        private BouttonToggle chkBoxDim2;
        private PanelSE _panelSe;

        public PanelSettings()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.titrePanel = new System.Windows.Forms.TextBox();
            this.flpDxf = new System.Windows.Forms.FlowLayoutPanel();
            this.dxfSetting1Txt = new System.Windows.Forms.TextBox();
            this.chkBoxDxfTag1 = new Application_Cyrell.Utils.BouttonToggle();
            this.btnReturn = new System.Windows.Forms.Button();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.flpDim = new System.Windows.Forms.FlowLayoutPanel();
            this.txtDim1 = new System.Windows.Forms.TextBox();
            this.chkBoxDim1 = new Application_Cyrell.Utils.BouttonToggle();
            this.txtDim2 = new System.Windows.Forms.TextBox();
            this.chkBoxDim2 = new Application_Cyrell.Utils.BouttonToggle();
            this.flpDxf.SuspendLayout();
            this.flpDim.SuspendLayout();
            this.SuspendLayout();
            // 
            // titrePanel
            // 
            this.titrePanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.titrePanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.titrePanel.Font = new System.Drawing.Font("Microsoft Sans Serif", 19F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titrePanel.ForeColor = System.Drawing.Color.Orange;
            this.titrePanel.Location = new System.Drawing.Point(431, 25);
            this.titrePanel.Name = "titrePanel";
            this.titrePanel.Size = new System.Drawing.Size(211, 44);
            this.titrePanel.TabIndex = 0;
            this.titrePanel.Text = "Paramètres";
            // 
            // flpDxf
            // 
            this.flpDxf.Controls.Add(this.dxfSetting1Txt);
            this.flpDxf.Controls.Add(this.chkBoxDxfTag1);
            this.flpDxf.Location = new System.Drawing.Point(89, 100);
            this.flpDxf.Name = "flpDxf";
            this.flpDxf.Size = new System.Drawing.Size(976, 45);
            this.flpDxf.TabIndex = 1;
            // 
            // dxfSetting1Txt
            // 
            this.dxfSetting1Txt.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.dxfSetting1Txt.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dxfSetting1Txt.Cursor = System.Windows.Forms.Cursors.Default;
            this.dxfSetting1Txt.Dock = System.Windows.Forms.DockStyle.Top;
            this.dxfSetting1Txt.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dxfSetting1Txt.ForeColor = System.Drawing.Color.NavajoWhite;
            this.dxfSetting1Txt.Location = new System.Drawing.Point(3, 3);
            this.dxfSetting1Txt.Name = "dxfSetting1Txt";
            this.dxfSetting1Txt.Size = new System.Drawing.Size(807, 32);
            this.dxfSetting1Txt.TabIndex = 0;
            this.dxfSetting1Txt.Text = "Garder DXF ouverts après TAG:";
            // 
            // chkBoxDxfTag1
            // 
            this.chkBoxDxfTag1.Checked = true;
            this.chkBoxDxfTag1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBoxDxfTag1.Location = new System.Drawing.Point(816, 3);
            this.chkBoxDxfTag1.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDxfTag1.Name = "chkBoxDxfTag1";
            this.chkBoxDxfTag1.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDxfTag1.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDxfTag1.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDxfTag1.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDxfTag1.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDxfTag1.TabIndex = 11;
            this.chkBoxDxfTag1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDxfTag1.UseVisualStyleBackColor = true;
            // 
            // btnReturn
            // 
            this.btnReturn.BackColor = System.Drawing.Color.Transparent;
            this.btnReturn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnReturn.FlatAppearance.BorderSize = 0;
            this.btnReturn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnReturn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnReturn.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.btnReturn.Location = new System.Drawing.Point(936, 710);
            this.btnReturn.Name = "btnReturn";
            this.btnReturn.Size = new System.Drawing.Size(129, 40);
            this.btnReturn.TabIndex = 8;
            this.btnReturn.Text = "Retour";
            this.btnReturn.UseVisualStyleBackColor = false;
            this.btnReturn.Click += new System.EventHandler(this.btnReturn_Click);
            // 
            // textBox5
            // 
            this.textBox5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.textBox5.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox5.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox5.ForeColor = System.Drawing.Color.SpringGreen;
            this.textBox5.Location = new System.Drawing.Point(89, 63);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(211, 37);
            this.textBox5.TabIndex = 9;
            this.textBox5.Text = "Taguer Dxf";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.Color.SpringGreen;
            this.textBox1.Location = new System.Drawing.Point(89, 162);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(299, 37);
            this.textBox1.TabIndex = 12;
            this.textBox1.Text = "Extraire Dimensions";
            // 
            // flpDim
            // 
            this.flpDim.Controls.Add(this.txtDim1);
            this.flpDim.Controls.Add(this.chkBoxDim1);
            this.flpDim.Controls.Add(this.txtDim2);
            this.flpDim.Controls.Add(this.chkBoxDim2);
            this.flpDim.Location = new System.Drawing.Point(89, 199);
            this.flpDim.Name = "flpDim";
            this.flpDim.Size = new System.Drawing.Size(976, 75);
            this.flpDim.TabIndex = 11;
            // 
            // txtDim1
            // 
            this.txtDim1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.txtDim1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtDim1.Cursor = System.Windows.Forms.Cursors.Default;
            this.txtDim1.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtDim1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDim1.ForeColor = System.Drawing.Color.NavajoWhite;
            this.txtDim1.Location = new System.Drawing.Point(3, 3);
            this.txtDim1.Name = "txtDim1";
            this.txtDim1.Size = new System.Drawing.Size(807, 32);
            this.txtDim1.TabIndex = 0;
            this.txtDim1.Text = "Afficher Message pour chaque pièce non-dépliée:";
            // 
            // chkBoxDim1
            // 
            this.chkBoxDim1.Checked = true;
            this.chkBoxDim1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBoxDim1.Location = new System.Drawing.Point(816, 3);
            this.chkBoxDim1.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDim1.Name = "chkBoxDim1";
            this.chkBoxDim1.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDim1.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDim1.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDim1.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDim1.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDim1.TabIndex = 12;
            this.chkBoxDim1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDim1.UseVisualStyleBackColor = true;
            // 
            // txtDim2
            // 
            this.txtDim2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.txtDim2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtDim2.Cursor = System.Windows.Forms.Cursors.Default;
            this.txtDim2.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtDim2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDim2.ForeColor = System.Drawing.Color.NavajoWhite;
            this.txtDim2.Location = new System.Drawing.Point(3, 41);
            this.txtDim2.Name = "txtDim2";
            this.txtDim2.Size = new System.Drawing.Size(807, 32);
            this.txtDim2.TabIndex = 13;
            this.txtDim2.Text = "Garder pièces choisies ouvertes dans Solid Edge:";
            // 
            // chkBoxDim2
            // 
            this.chkBoxDim2.Location = new System.Drawing.Point(816, 41);
            this.chkBoxDim2.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDim2.Name = "chkBoxDim2";
            this.chkBoxDim2.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDim2.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDim2.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDim2.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDim2.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDim2.TabIndex = 14;
            this.chkBoxDim2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDim2.UseVisualStyleBackColor = true;
            // 
            // PanelSettings
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(47)))), ((int)(((byte)(86)))));
            this.ClientSize = new System.Drawing.Size(1123, 798);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.flpDim);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.btnReturn);
            this.Controls.Add(this.flpDxf);
            this.Controls.Add(this.titrePanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "PanelSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.flpDxf.ResumeLayout(false);
            this.flpDxf.PerformLayout();
            this.flpDim.ResumeLayout(false);
            this.flpDim.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        public bool paramTag()
        {
            if (chkBoxDxfTag1.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool paramDim1()
        {
            if (chkBoxDim1.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool paramDim2()
        {
            if (chkBoxDim2.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void InitializeParent(PanelSE panelSe)
        {
            _panelSe = panelSe;
        }

        private void btnReturn_Click(object sender, EventArgs e)
        {
            var mainForm = this.ParentForm as MainForm;
            if (mainForm != null && _panelSe != null)
            {
                mainForm.OpenChildForm(() => _panelSe);
            }
            this.Hide();
        }
    }
}