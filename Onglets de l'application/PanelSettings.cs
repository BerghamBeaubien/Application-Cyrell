﻿using System;
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
        private FlowLayoutPanel flpDft;
        private TextBox dftTxt1;
        private TextBox textBox5;
        private TextBox textBox6;
        private BouttonToggle chkBoxDxfTag1;
        private BouttonToggle chkBoxDft1;
        private TextBox dftTxt2;
        private BouttonToggle chkBoxDft2;
        private TextBox textBox1;
        private FlowLayoutPanel flpDim;
        private TextBox txtDim1;
        private BouttonToggle chkBoxDim1;
        private TextBox txtDim2;
        private BouttonToggle chkBoxDim2;
        private TextBox dftTxt3;
        private BouttonToggle chkBoxDft3;
        private TextBox dftTxt4;
        private BouttonToggle chkBoxDft4;
        private TextBox dftTxt5;
        private BouttonToggle chkBoxDft5;
        private TextBox dftTxt6;
        private BouttonToggle chkBoxDft6;
        private TextBox dftTxt7;
        private BouttonToggle chkBoxDft7;
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
            this.flpDft = new System.Windows.Forms.FlowLayoutPanel();
            this.dftTxt1 = new System.Windows.Forms.TextBox();
            this.chkBoxDft1 = new Application_Cyrell.Utils.BouttonToggle();
            this.dftTxt2 = new System.Windows.Forms.TextBox();
            this.chkBoxDft2 = new Application_Cyrell.Utils.BouttonToggle();
            this.dftTxt3 = new System.Windows.Forms.TextBox();
            this.chkBoxDft3 = new Application_Cyrell.Utils.BouttonToggle();
            this.dftTxt4 = new System.Windows.Forms.TextBox();
            this.chkBoxDft4 = new Application_Cyrell.Utils.BouttonToggle();
            this.dftTxt5 = new System.Windows.Forms.TextBox();
            this.chkBoxDft5 = new Application_Cyrell.Utils.BouttonToggle();
            this.dftTxt6 = new System.Windows.Forms.TextBox();
            this.chkBoxDft6 = new Application_Cyrell.Utils.BouttonToggle();
            this.dftTxt7 = new System.Windows.Forms.TextBox();
            this.chkBoxDft7 = new Application_Cyrell.Utils.BouttonToggle();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.flpDim = new System.Windows.Forms.FlowLayoutPanel();
            this.txtDim1 = new System.Windows.Forms.TextBox();
            this.chkBoxDim1 = new Application_Cyrell.Utils.BouttonToggle();
            this.txtDim2 = new System.Windows.Forms.TextBox();
            this.chkBoxDim2 = new Application_Cyrell.Utils.BouttonToggle();
            this.flpDxf.SuspendLayout();
            this.flpDft.SuspendLayout();
            this.flpDim.SuspendLayout();
            this.SuspendLayout();
            // 
            // titrePanel
            // 
            this.titrePanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
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
            this.flpDxf.Location = new System.Drawing.Point(89, 127);
            this.flpDxf.Name = "flpDxf";
            this.flpDxf.Size = new System.Drawing.Size(976, 45);
            this.flpDxf.TabIndex = 1;
            // 
            // dxfSetting1Txt
            // 
            this.dxfSetting1Txt.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
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
            // flpDft
            // 
            this.flpDft.Controls.Add(this.dftTxt1);
            this.flpDft.Controls.Add(this.chkBoxDft1);
            this.flpDft.Controls.Add(this.dftTxt2);
            this.flpDft.Controls.Add(this.chkBoxDft2);
            this.flpDft.Controls.Add(this.dftTxt3);
            this.flpDft.Controls.Add(this.chkBoxDft3);
            this.flpDft.Controls.Add(this.dftTxt4);
            this.flpDft.Controls.Add(this.chkBoxDft4);
            this.flpDft.Controls.Add(this.dftTxt5);
            this.flpDft.Controls.Add(this.chkBoxDft5);
            this.flpDft.Controls.Add(this.dftTxt6);
            this.flpDft.Controls.Add(this.chkBoxDft6);
            this.flpDft.Controls.Add(this.dftTxt7);
            this.flpDft.Controls.Add(this.chkBoxDft7);
            this.flpDft.Location = new System.Drawing.Point(89, 407);
            this.flpDft.Name = "flpDft";
            this.flpDft.Size = new System.Drawing.Size(976, 288);
            this.flpDft.TabIndex = 4;
            // 
            // dftTxt1
            // 
            this.dftTxt1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.dftTxt1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dftTxt1.Cursor = System.Windows.Forms.Cursors.Default;
            this.dftTxt1.Dock = System.Windows.Forms.DockStyle.Top;
            this.dftTxt1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dftTxt1.ForeColor = System.Drawing.Color.NavajoWhite;
            this.dftTxt1.Location = new System.Drawing.Point(3, 3);
            this.dftTxt1.Name = "dftTxt1";
            this.dftTxt1.Size = new System.Drawing.Size(807, 32);
            this.dftTxt1.TabIndex = 0;
            this.dftTxt1.Text = "Générer Nomenclature (ASM & PAR):";
            // 
            // chkBoxDft1
            // 
            this.chkBoxDft1.Checked = true;
            this.chkBoxDft1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBoxDft1.Location = new System.Drawing.Point(816, 3);
            this.chkBoxDft1.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDft1.Name = "chkBoxDft1";
            this.chkBoxDft1.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDft1.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDft1.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDft1.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDft1.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDft1.TabIndex = 12;
            this.chkBoxDft1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDft1.UseVisualStyleBackColor = true;
            // 
            // dftTxt2
            // 
            this.dftTxt2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.dftTxt2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dftTxt2.Cursor = System.Windows.Forms.Cursors.Default;
            this.dftTxt2.Dock = System.Windows.Forms.DockStyle.Top;
            this.dftTxt2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dftTxt2.ForeColor = System.Drawing.Color.NavajoWhite;
            this.dftTxt2.Location = new System.Drawing.Point(3, 41);
            this.dftTxt2.MinimumSize = new System.Drawing.Size(807, 36);
            this.dftTxt2.Name = "dftTxt2";
            this.dftTxt2.Size = new System.Drawing.Size(807, 36);
            this.dftTxt2.TabIndex = 13;
            this.dftTxt2.Text = "Créer un dessin pour chaque pièce individuel d\'un assemblage:";
            // 
            // chkBoxDft2
            // 
            this.chkBoxDft2.Checked = true;
            this.chkBoxDft2.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBoxDft2.Location = new System.Drawing.Point(816, 41);
            this.chkBoxDft2.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDft2.Name = "chkBoxDft2";
            this.chkBoxDft2.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDft2.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDft2.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDft2.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDft2.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDft2.TabIndex = 14;
            this.chkBoxDft2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDft2.UseVisualStyleBackColor = true;
            // 
            // dftTxt3
            // 
            this.dftTxt3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.dftTxt3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dftTxt3.Cursor = System.Windows.Forms.Cursors.Default;
            this.dftTxt3.Dock = System.Windows.Forms.DockStyle.Top;
            this.dftTxt3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dftTxt3.ForeColor = System.Drawing.Color.NavajoWhite;
            this.dftTxt3.Location = new System.Drawing.Point(3, 83);
            this.dftTxt3.Name = "dftTxt3";
            this.dftTxt3.Size = new System.Drawing.Size(807, 32);
            this.dftTxt3.TabIndex = 15;
            this.dftTxt3.Text = "Générer Vue Isométrique:";
            // 
            // chkBoxDft3
            // 
            this.chkBoxDft3.Checked = true;
            this.chkBoxDft3.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBoxDft3.Location = new System.Drawing.Point(816, 83);
            this.chkBoxDft3.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDft3.Name = "chkBoxDft3";
            this.chkBoxDft3.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDft3.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDft3.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDft3.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDft3.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDft3.TabIndex = 16;
            this.chkBoxDft3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDft3.UseVisualStyleBackColor = true;
            // 
            // dftTxt4
            // 
            this.dftTxt4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.dftTxt4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dftTxt4.Cursor = System.Windows.Forms.Cursors.Default;
            this.dftTxt4.Dock = System.Windows.Forms.DockStyle.Top;
            this.dftTxt4.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dftTxt4.ForeColor = System.Drawing.Color.NavajoWhite;
            this.dftTxt4.Location = new System.Drawing.Point(3, 121);
            this.dftTxt4.Name = "dftTxt4";
            this.dftTxt4.Size = new System.Drawing.Size(807, 32);
            this.dftTxt4.TabIndex = 17;
            this.dftTxt4.Text = "Générer Vue Flat:";
            // 
            // chkBoxDft4
            // 
            this.chkBoxDft4.Checked = true;
            this.chkBoxDft4.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBoxDft4.Location = new System.Drawing.Point(816, 121);
            this.chkBoxDft4.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDft4.Name = "chkBoxDft4";
            this.chkBoxDft4.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDft4.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDft4.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDft4.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDft4.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDft4.TabIndex = 18;
            this.chkBoxDft4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDft4.UseVisualStyleBackColor = true;
            // 
            // dftTxt5
            // 
            this.dftTxt5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.dftTxt5.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dftTxt5.Cursor = System.Windows.Forms.Cursors.Default;
            this.dftTxt5.Dock = System.Windows.Forms.DockStyle.Top;
            this.dftTxt5.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dftTxt5.ForeColor = System.Drawing.Color.NavajoWhite;
            this.dftTxt5.Location = new System.Drawing.Point(3, 159);
            this.dftTxt5.MinimumSize = new System.Drawing.Size(807, 36);
            this.dftTxt5.Name = "dftTxt5";
            this.dftTxt5.Size = new System.Drawing.Size(807, 36);
            this.dftTxt5.TabIndex = 19;
            this.dftTxt5.Text = "Générer Bend Tables:";
            // 
            // chkBoxDft5
            // 
            this.chkBoxDft5.Checked = true;
            this.chkBoxDft5.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBoxDft5.Location = new System.Drawing.Point(816, 159);
            this.chkBoxDft5.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDft5.Name = "chkBoxDft5";
            this.chkBoxDft5.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDft5.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDft5.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDft5.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDft5.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDft5.TabIndex = 20;
            this.chkBoxDft5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDft5.UseVisualStyleBackColor = true;
            // 
            // dftTxt6
            // 
            this.dftTxt6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.dftTxt6.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dftTxt6.Cursor = System.Windows.Forms.Cursors.Default;
            this.dftTxt6.Dock = System.Windows.Forms.DockStyle.Top;
            this.dftTxt6.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dftTxt6.ForeColor = System.Drawing.Color.NavajoWhite;
            this.dftTxt6.Location = new System.Drawing.Point(3, 201);
            this.dftTxt6.MinimumSize = new System.Drawing.Size(807, 36);
            this.dftTxt6.Name = "dftTxt6";
            this.dftTxt6.Size = new System.Drawing.Size(807, 36);
            this.dftTxt6.TabIndex = 21;
            this.dftTxt6.Text = "Executer la Macro DenMarForr7:";
            // 
            // chkBoxDft6
            // 
            this.chkBoxDft6.Checked = true;
            this.chkBoxDft6.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBoxDft6.Location = new System.Drawing.Point(816, 201);
            this.chkBoxDft6.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDft6.Name = "chkBoxDft6";
            this.chkBoxDft6.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDft6.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDft6.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDft6.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDft6.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDft6.TabIndex = 22;
            this.chkBoxDft6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDft6.UseVisualStyleBackColor = true;
            // 
            // dftTxt7
            // 
            this.dftTxt7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.dftTxt7.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dftTxt7.Cursor = System.Windows.Forms.Cursors.Default;
            this.dftTxt7.Dock = System.Windows.Forms.DockStyle.Top;
            this.dftTxt7.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dftTxt7.ForeColor = System.Drawing.Color.NavajoWhite;
            this.dftTxt7.Location = new System.Drawing.Point(3, 243);
            this.dftTxt7.MinimumSize = new System.Drawing.Size(807, 36);
            this.dftTxt7.Name = "dftTxt7";
            this.dftTxt7.Size = new System.Drawing.Size(807, 36);
            this.dftTxt7.TabIndex = 23;
            this.dftTxt7.Text = "Générer Résumé des Quantitées de Pièces (Excel) :";
            // 
            // chkBoxDft7
            // 
            this.chkBoxDft7.Location = new System.Drawing.Point(816, 243);
            this.chkBoxDft7.MinimumSize = new System.Drawing.Size(45, 22);
            this.chkBoxDft7.Name = "chkBoxDft7";
            this.chkBoxDft7.OffBackColor = System.Drawing.Color.Gray;
            this.chkBoxDft7.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.chkBoxDft7.OnBackColor = System.Drawing.Color.LimeGreen;
            this.chkBoxDft7.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.chkBoxDft7.Size = new System.Drawing.Size(77, 32);
            this.chkBoxDft7.TabIndex = 24;
            this.chkBoxDft7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkBoxDft7.UseVisualStyleBackColor = true;
            // 
            // textBox5
            // 
            this.textBox5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.textBox5.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox5.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox5.ForeColor = System.Drawing.Color.SpringGreen;
            this.textBox5.Location = new System.Drawing.Point(89, 84);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(211, 37);
            this.textBox5.TabIndex = 9;
            this.textBox5.Text = "Tag Dxf";
            // 
            // textBox6
            // 
            this.textBox6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.textBox6.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox6.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox6.ForeColor = System.Drawing.Color.SpringGreen;
            this.textBox6.Location = new System.Drawing.Point(89, 364);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(211, 37);
            this.textBox6.TabIndex = 10;
            this.textBox6.Text = "Generer DFT";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.Color.SpringGreen;
            this.textBox1.Location = new System.Drawing.Point(89, 195);
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
            this.flpDim.Location = new System.Drawing.Point(89, 248);
            this.flpDim.Name = "flpDim";
            this.flpDim.Size = new System.Drawing.Size(976, 96);
            this.flpDim.TabIndex = 11;
            // 
            // txtDim1
            // 
            this.txtDim1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
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
            this.txtDim2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
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
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.ClientSize = new System.Drawing.Size(1123, 798);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.flpDim);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.btnReturn);
            this.Controls.Add(this.flpDxf);
            this.Controls.Add(this.titrePanel);
            this.Controls.Add(this.flpDft);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "PanelSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.flpDxf.ResumeLayout(false);
            this.flpDxf.PerformLayout();
            this.flpDft.ResumeLayout(false);
            this.flpDft.PerformLayout();
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

        public bool paramDft1()
        {
            if (chkBoxDft1.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool paramDft2()
        {
            if (chkBoxDft2.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool paramDft3()
        {
            if (chkBoxDft3.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool paramDft4()
        {
            if (chkBoxDft4.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool paramDft5()
        {
            if (chkBoxDft5.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool paramDft6()
        {
            if (chkBoxDft6.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool paramDft7()
        {
            if (chkBoxDft7.Checked)
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