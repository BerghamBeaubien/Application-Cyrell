using System;
using System.Drawing;
using System.Windows.Forms;
using Application_Cyrell.Utils;

public class FormulaireDFT : Form
{
    private TabControl tabControl;
    private TabPage tabStandard;
    private TabPage tabAvance;
    private Label lblExplication;

    // Onglet Standard
    private BouttonToggle tglParListPieceSolo;
    private BouttonToggle tglDftIndividuelAssemblage;
    private BouttonToggle tglIsoView;
    private BouttonToggle tglFlatView;
    private BouttonToggle tglBendTableToggle;
    private BouttonToggle tglRefVars;
    private BouttonToggle tglCountParts;
    private BouttonToggle tglScaleModeStandard;
    private Label lblScaleModeStandard;
    private Label lblCustomScaleStandard;
    private TextBox txtCustomScaleStandard;

    // Labels for standard tab
    private Label lblBendTableToggle;

    // Onglet Avancé
    private BouttonToggle tglBendTableToggleAvance;
    private Label lblScaleMode;
    private BouttonToggle tglScaleMode;
    private Label lblCustomScale;
    private TextBox txtCustomScale;
    private Label lblSpacingX;
    private Label lblSpacingY;
    private TextBox txtSpacingX;
    private TextBox txtSpacingY;
    private BouttonToggle tglRefVarsAvance;
    private BouttonToggle tglParListPieceSoloAvance;
    private Label lblRefVarsAvance;

    // Boutons de validation
    private Button btnContinue;
    private Button btnContinue2;
    private Button btnCancel;

    // Propriétés publiques pour accéder aux valeurs
    public bool ParamParListPieceSolo => tglParListPieceSolo.Checked;
    public bool ParamDftIndividuelAssemblage => tglDftIndividuelAssemblage.Checked;
    public bool ParamIsoView => tglIsoView.Checked;
    public bool ParamFlatView => tglFlatView.Checked;
    public bool ParamBendTableToggle => tglBendTableToggle.Checked;
    public bool ParamRefVars => tglRefVars.Checked || tglRefVarsAvance.Checked;
    public bool ParamCountParts => tglCountParts.Checked;

    public bool ParamBendTableToggleAvance => tglBendTableToggleAvance.Checked;
    public bool ParamPartsListAvance => tglParListPieceSoloAvance.Checked;
    public bool IsAutoScaleStandard => tglScaleModeStandard.Checked;
    public double CustomScaleStandard => double.TryParse(txtCustomScaleStandard.Text, out double scaleStd) ? scaleStd : 0.5;
    public double CustomScale => double.TryParse(txtCustomScale.Text, out double scale) ? scale : 0.5;
    public bool IsAutoScale => tglScaleMode.Checked;
    public double SpacingX => double.TryParse(txtSpacingX.Text, out double x) ? x : 2;
    public double SpacingY => double.TryParse(txtSpacingY.Text, out double y) ? y : 1.5;

    // Add delegates for the two continue buttons
    public delegate void ContinueHandler(bool parListPieceSolo, bool dftIndividuel, bool isoView, bool flatView, bool bendTable, bool refVars, bool countParts,
        bool bendTableAdv, bool parListAvance, bool autoScale, double scale, double spacingX, double spacingY);
    public event ContinueHandler OnContinue1;
    public event ContinueHandler OnContinue2;

    public FormulaireDFT()
    {
        InitializeComponents();
        ConfigureEvents();

        // Set default values
        tglParListPieceSolo.Checked = true;
        tglDftIndividuelAssemblage.Checked = true;
        tglIsoView.Checked = true;
        tglFlatView.Checked = false;
        tglBendTableToggle.Checked = false;
        tglRefVars.Checked = true;
        tglRefVarsAvance.Checked = true;
        tglCountParts.Checked = false;
        tglBendTableToggleAvance.Checked = true;
        tglScaleModeStandard.Checked = true;
        tglScaleMode.Checked = true;

        // Update visibility based on initial settings
        UpdateScaleControls();

        // Set initial button visibility
        btnContinue.Visible = true;
        btnContinue2.Visible = false;

    }

    private void TabControlDFT_SelectedIndexChanged(object sender, EventArgs e)
    {
        AfficherExplication();
    }
    private void InitializeComponents()
    {
        this.Text = "Paramètres DFT";
        this.Size = new Size(600, 480); // Increased height to accommodate new controls
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.Font = new Font("Segoe UI", 10F);

        // Configuration du TabControl
        tabControl = new TabControl
        {
            Location = new Point(10, 10),
            Size = new Size(570, 300), // Increased height
            Appearance = TabAppearance.FlatButtons,
            ItemSize = new Size(120, 30),
            SizeMode = TabSizeMode.Fixed
        };

        // Création des onglets
        tabStandard = new TabPage("Standard");
        tabAvance = new TabPage("Avancé");

        // Style des onglets pour ressembler à un navigateur moderne
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.DrawItem += TabControl_DrawItem;

        // Initialisation des contrôles de l'onglet Standard
        InitializeStandardTab();

        // Initialisation des contrôles de l'onglet Avancé
        InitializeAdvancedTab();

        // Boutons (modified to have two continue buttons)
        btnContinue = new Button
        {
            Text = "Continuer",
            Location = new Point(400, 380), // Adjusted position
            Size = new Size(90, 30),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.Green,
            ForeColor = Color.White
        };
        btnContinue.FlatAppearance.BorderSize = 0;
        btnContinue.Click += BtnContinue1_Click;

        btnContinue2 = new Button
        {
            Text = "Continuer",
            Location = new Point(400, 380), // Adjusted position
            Size = new Size(90, 30),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.SteelBlue,
            ForeColor = Color.White
        };
        btnContinue2.FlatAppearance.BorderSize = 0;
        btnContinue2.Click += BtnContinue2_Click;

        btnCancel = new Button
        {
            Text = "Annuler",
            DialogResult = DialogResult.Cancel,
            Location = new Point(500, 380), // Adjusted position
            Size = new Size(80, 30),
            FlatStyle = FlatStyle.Flat
        };
        btnCancel.FlatAppearance.BorderSize = 0;

        // Ajout des contrôles au formulaire
        tabControl.TabPages.Add(tabStandard);
        tabControl.TabPages.Add(tabAvance);

        // Create the explanation box
        lblExplication = new Label
        {
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Segoe UI", 10),
            Size = new Size(540, 40),
            Location = new Point(20, 320)
        };

        tabControl.SelectedIndexChanged += TabControlDFT_SelectedIndexChanged;
        tabControl.SelectedIndex = 0;
        AfficherExplication();

        this.Controls.Add(tabControl);
        this.Controls.Add(lblExplication);
        this.Controls.Add(btnContinue);
        this.Controls.Add(btnContinue2);
        this.Controls.Add(btnCancel);

        this.CancelButton = btnCancel;
    }

    private void AfficherExplication()
    {
        string nomOnglet = tabControl.SelectedTab.Text;

        if (nomOnglet == "Standard")
        {
            lblExplication.Text = "Le mode Standard permet de générer un dessin (DFT) de chaque pièce séléctionnée.\n" +
                "Chaque pièce sera dans une page différente";
        }
        else if (nomOnglet == "Avancé")
        {
            lblExplication.Text = "Le mode Avancé permet de générer un dessin (DFT) contenant toutes les pièces séléctionnées. " +
                "Chaque pièce est éspacée dans la page avec les valeurs spécifiées.";
        }
        else
        {
            lblExplication.Text = "";
        }
    }

    private void BtnContinue1_Click(object sender, EventArgs e)
    {
        double scale = IsAutoScaleStandard ? 0 : CustomScaleStandard;

        OnContinue1?.Invoke(
            ParamParListPieceSolo,
            ParamDftIndividuelAssemblage,
            ParamIsoView,
            ParamFlatView,
            ParamBendTableToggle,
            ParamRefVars,
            ParamCountParts,
            ParamBendTableToggleAvance,
            ParamPartsListAvance,
            IsAutoScaleStandard,
            scale,
            SpacingX,
            SpacingY
        );
        this.DialogResult = DialogResult.OK;
    }

    private void BtnContinue2_Click(object sender, EventArgs e)
    {
        double scale = IsAutoScale ? 0 : CustomScale;

        OnContinue2?.Invoke(
            ParamParListPieceSolo,
            ParamDftIndividuelAssemblage,
            ParamIsoView,
            ParamFlatView,
            ParamBendTableToggle,
            ParamRefVars,
            ParamCountParts,
            ParamBendTableToggleAvance,
            ParamPartsListAvance,
            IsAutoScale,
            scale,
            SpacingX,
            SpacingY
        );
        this.DialogResult = DialogResult.OK;
    }

    private void InitializeStandardTab()
    {
        tabStandard.BackColor = Color.White;

        int leftMargin = 30;
        int verticalSpacing = 40;
        int toggleWidth = 55;
        int toggleHeight = 25;
        int startY = 30;

        // Les labels et toggles seront alignés sur deux colonnes
        int column1X = leftMargin;
        int column2X = 300;
        int toggleOffsetX = 200;

        // Première colonne
        CreateToggleWithLabel(tabStandard, "Poser 'Parts List':", column1X, startY,
            out Label lblParListPieceSolo, out tglParListPieceSolo, toggleOffsetX);

        CreateToggleWithLabel(tabStandard, "Dessin individuel pour pièces d'un assemblage:", column1X, startY + verticalSpacing,
            out Label lblDftIndividuelAssemblage, out tglDftIndividuelAssemblage, toggleOffsetX + 200);

        CreateToggleWithLabel(tabStandard, "Vue isométrique:", column1X, startY + verticalSpacing * 2,
            out Label lblIsoView, out tglIsoView, toggleOffsetX);

        CreateToggleWithLabel(tabStandard, "Vue à plat:", column1X, startY + verticalSpacing * 3,
            out Label lblFlatView, out tglFlatView, toggleOffsetX);

        // Deuxième colonne
        CreateToggleWithLabel(tabStandard, "Générer Bend Table:", column1X, startY + verticalSpacing * 4,
            out lblBendTableToggle, out tglBendTableToggle, toggleOffsetX);

        CreateToggleWithLabel(tabStandard, "Rouler Macro DenMarForr7:", column2X, startY + verticalSpacing * 2,
            out Label lblRefVars, out tglRefVars, toggleOffsetX);

        CreateToggleWithLabel(tabStandard, "Générer Rapport des Pièces:", column2X, startY + verticalSpacing * 3,
            out Label lblCountParts, out tglCountParts, toggleOffsetX);

        // Add scale controls to standard tab
        lblScaleModeStandard = new Label
        {
            Text = "Echelle automatique :",
            Location = new Point(column2X, startY + verticalSpacing * 4),
            Size = new Size(toggleOffsetX - 10, 25),
            TextAlign = ContentAlignment.MiddleLeft
        };

        tglScaleModeStandard = new BouttonToggle
        {
            Location = new Point(column2X + toggleOffsetX, startY + verticalSpacing * 4),
            Size = new Size(55, 25),
            Checked = true // Automatique par défaut
        };

        lblCustomScaleStandard = new Label
        {
            Text = "Échelle personnalisée:",
            Location = new Point(column2X, startY + verticalSpacing * 5),
            Size = new Size(toggleOffsetX - 10, 25),
            TextAlign = ContentAlignment.MiddleLeft,
            Visible = false
        };

        txtCustomScaleStandard = new TextBox
        {
            Location = new Point(column2X + toggleOffsetX, startY + verticalSpacing * 5),
            Size = new Size(50, 25),
            Text = "0.5", // Valeur par défaut
            Visible = false
        };

        tabStandard.Controls.Add(lblScaleModeStandard);
        tabStandard.Controls.Add(tglScaleModeStandard);
        tabStandard.Controls.Add(lblCustomScaleStandard);
        tabStandard.Controls.Add(txtCustomScaleStandard);

        // Initialiser les couleurs des toggles
        tglParListPieceSolo.OnBackColor = Color.Green;
        tglParListPieceSolo.OffBackColor = Color.LightGray;

        tglDftIndividuelAssemblage.OnBackColor = Color.Green;
        tglDftIndividuelAssemblage.OffBackColor = Color.LightGray;

        tglIsoView.OnBackColor = Color.Green;
        tglIsoView.OffBackColor = Color.LightGray;

        tglFlatView.OnBackColor = Color.Green;
        tglFlatView.OffBackColor = Color.LightGray;

        tglBendTableToggle.OnBackColor = Color.Green;
        tglBendTableToggle.OffBackColor = Color.LightGray;

        tglRefVars.OnBackColor = Color.Green;
        tglRefVars.OffBackColor = Color.LightGray;

        tglCountParts.OnBackColor = Color.Green;
        tglCountParts.OffBackColor = Color.LightGray;

        tglScaleModeStandard.OnBackColor = Color.Green;
        tglScaleModeStandard.OffBackColor = Color.LightGray;

        // La table de pliage ne devrait être visible que si FlatView est cochée
        lblBendTableToggle.Visible = false;
        tglBendTableToggle.Visible = false;
    }

    private void InitializeAdvancedTab()
    {
        tabAvance.BackColor = Color.White;

        int leftMargin = 30;
        int verticalSpacing = 40;
        int labelWidth = 150;
        int textboxWidth = 50;
        int startY = 30;

        // Toggle pour la table de pliage
        CreateToggleWithLabel(tabAvance, "Générer Bend Table:", leftMargin, startY,
            out Label lblBendTableToggleAvance, out tglBendTableToggleAvance, 200);

        // Add ParListPieceSolo toggle to advanced tab (replacing RefVars position)
        CreateToggleWithLabel(tabAvance, "Générer 'Nomenclature':", leftMargin, startY + verticalSpacing,
            out Label lblParListPieceSoloAvance, out tglParListPieceSoloAvance, 200);

        // Add RefVars toggle but 300 pixels to the right
        CreateToggleWithLabel(tabAvance, "Rouler macro DenMarForr7:", 300, startY + verticalSpacing,
            out lblRefVarsAvance, out tglRefVarsAvance, 200);

        // Set RefVars visibility to initially hidden - will be controlled by the ParList toggle
        lblRefVarsAvance.Visible = false;
        tglRefVarsAvance.Visible = false;

        // Add event to control RefVars visibility based on ParList toggle state
        tglParListPieceSoloAvance.CheckedChanged += (sender, e) => {
            lblRefVarsAvance.Visible = tglParListPieceSoloAvance.Checked;
            tglRefVarsAvance.Visible = tglParListPieceSoloAvance.Checked;
        };

        tglBendTableToggleAvance.OnBackColor = Color.SteelBlue;
        tglBendTableToggleAvance.OffBackColor = Color.LightGray;

        tglParListPieceSoloAvance.OnBackColor = Color.SteelBlue;
        tglParListPieceSoloAvance.OffBackColor = Color.LightGray;

        tglRefVarsAvance.OnBackColor = Color.SteelBlue;
        tglRefVarsAvance.OffBackColor = Color.LightGray;

        // Toggle pour le mode d'échelle
        lblScaleMode = new Label
        {
            Text = "Echelle automatique :",
            Location = new Point(leftMargin, startY + verticalSpacing * 2),
            Size = new Size(labelWidth, 25),
            TextAlign = ContentAlignment.MiddleLeft
        };

        tglScaleMode = new BouttonToggle
        {
            Location = new Point(leftMargin + labelWidth + 10, startY + verticalSpacing * 2),
            Size = new Size(55, 25),
            Checked = true // Automatique par défaut
        };

        tglScaleMode.OnBackColor = Color.SteelBlue;
        tglScaleMode.OffBackColor = Color.LightGray;

        tabAvance.Controls.Add(lblScaleMode);
        tabAvance.Controls.Add(tglScaleMode);

        // TextBox pour l'échelle personnalisée
        lblCustomScale = new Label
        {
            Text = "Échelle personnalisée:",
            Location = new Point(leftMargin, startY + verticalSpacing * 3),
            Size = new Size(labelWidth, 25),
            TextAlign = ContentAlignment.MiddleLeft,
            Visible = false
        };

        txtCustomScale = new TextBox
        {
            Location = new Point(leftMargin + labelWidth + 10, startY + verticalSpacing * 3),
            Size = new Size(textboxWidth, 25),
            Text = "0.5", // Valeur par défaut
            Visible = false
        };

        // TextBox pour l'espacement X
        lblSpacingX = new Label
        {
            Text = "Espacement X (pouces):",
            Location = new Point(leftMargin, startY + verticalSpacing * 4),
            Size = new Size(labelWidth + 20, 25),
            TextAlign = ContentAlignment.MiddleLeft
        };

        txtSpacingX = new TextBox
        {
            Location = new Point(leftMargin + labelWidth + 30, startY + verticalSpacing * 4),
            Size = new Size(textboxWidth, 25),
            Text = "10" // Valeur par défaut
        };

        // TextBox pour l'espacement Y
        lblSpacingY = new Label
        {
            Text = "Espacement Y (pouces):",
            Location = new Point(leftMargin, startY + verticalSpacing * 5),
            Size = new Size(labelWidth + 20, 25),
            TextAlign = ContentAlignment.MiddleLeft
        };

        txtSpacingY = new TextBox
        {
            Location = new Point(leftMargin + labelWidth + 30, startY + verticalSpacing * 5),
            Size = new Size(textboxWidth, 25),
            Text = "6" // Valeur par défaut
        };

        // Ajout des contrôles à l'onglet Avancé
        tabAvance.Controls.Add(lblBendTableToggleAvance);
        tabAvance.Controls.Add(tglBendTableToggleAvance);
        tabAvance.Controls.Add(lblParListPieceSoloAvance);
        tabAvance.Controls.Add(tglParListPieceSoloAvance);
        tabAvance.Controls.Add(lblRefVarsAvance);
        tabAvance.Controls.Add(tglRefVarsAvance);
        tabAvance.Controls.Add(lblCustomScale);
        tabAvance.Controls.Add(txtCustomScale);
        tabAvance.Controls.Add(lblSpacingX);
        tabAvance.Controls.Add(txtSpacingX);
        tabAvance.Controls.Add(lblSpacingY);
        tabAvance.Controls.Add(txtSpacingY);
    }
    private void CreateToggleWithLabel(Control parent, string labelText, int x, int y,
                                     out Label label, out BouttonToggle toggle, int toggleOffsetX)
    {
        label = new Label
        {
            Text = labelText,
            Location = new Point(x, y),
            Size = new Size(toggleOffsetX - 10, 25),
            TextAlign = ContentAlignment.MiddleLeft
        };

        toggle = new BouttonToggle
        {
            Location = new Point(x + toggleOffsetX, y),
            Size = new Size(55, 25),
            Checked = false
        };

        parent.Controls.Add(label);
        parent.Controls.Add(toggle);
    }

    private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
    {
        // Style des onglets pour ressembler à un navigateur moderne
        Graphics g = e.Graphics;
        TabPage tabPage = tabControl.TabPages[e.Index];
        Rectangle tabBounds = tabControl.GetTabRect(e.Index);

        // Vérifier si cet onglet est sélectionné
        bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

        // Remplir le fond
        using (SolidBrush brush = new SolidBrush(isSelected ? Color.White : Color.WhiteSmoke))
        {
            g.FillRectangle(brush, tabBounds);
        }

        // Dessiner une ligne de couleur en bas de l'onglet sélectionné
        if (isSelected)
        {
            using (SolidBrush brush = new SolidBrush(isSelected ?
                (tabPage == tabStandard ? Color.Green : Color.SteelBlue) : Color.Transparent))
            {
                g.FillRectangle(brush, tabBounds.X, tabBounds.Bottom - 3, tabBounds.Width, 3);
            }
        }

        // Centrer le texte
        StringFormat stringFormat = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center
        };

        // Dessiner le texte
        string tabText = tabControl.TabPages[e.Index].Text;
        using (Brush textBrush = new SolidBrush(Color.Black))
        {
            StringFormat sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(tabText, this.Font, textBrush, tabBounds, sf);
        }
    }

    private void ConfigureEvents()
    {
        // Mettre à jour la visibilité de tglBendTableToggle en fonction de l'état de tglFlatView
        tglFlatView.CheckedChanged += (s, e) =>
        {
            lblBendTableToggle.Visible = tglFlatView.Checked;
            tglBendTableToggle.Visible = tglFlatView.Checked;
        };

        // Update toggle visibility based on tab selection
        tabControl.SelectedIndexChanged += (s, e) =>
        {
            btnContinue.Visible = tabControl.SelectedTab == tabStandard;
            btnContinue2.Visible = tabControl.SelectedTab == tabAvance;
        };

        // Configure scale mode toggles
        tglScaleMode.CheckedChanged += (s, e) => UpdateScaleControls();
        tglScaleModeStandard.CheckedChanged += (s, e) => UpdateScaleControls();

        // Sync RefVars toggles between tabs
        tglRefVars.CheckedChanged += (s, e) => tglRefVarsAvance.Checked = tglRefVars.Checked;
        tglRefVarsAvance.CheckedChanged += (s, e) => tglRefVars.Checked = tglRefVarsAvance.Checked;

        // Valider les entrées pour les textboxes numériques
        txtCustomScale.KeyPress += ValidateNumericInputWithDecimal;
        txtCustomScaleStandard.KeyPress += ValidateNumericInputWithDecimal;
        txtSpacingX.KeyPress += ValidateNumericInputWithDecimal;
        txtSpacingY.KeyPress += ValidateNumericInputWithDecimal;
    }

    private void UpdateScaleControls()
    {
        // Update standard tab scale controls
        bool isAutoStd = tglScaleModeStandard.Checked;
        lblCustomScaleStandard.Visible = !isAutoStd;
        txtCustomScaleStandard.Visible = !isAutoStd;
        lblScaleModeStandard.Text = isAutoStd ? "Echelle automatique" : "Echelle manuelle";

        // Update advanced tab scale controls
        bool isAutoAdv = tglScaleMode.Checked;
        lblCustomScale.Visible = !isAutoAdv;
        txtCustomScale.Visible = !isAutoAdv;
        lblScaleMode.Text = isAutoAdv ? "Echelle automatique" : "Echelle manuelle";
    }

    private void ValidateNumericInputWithDecimal(object sender, KeyPressEventArgs e)
    {
        // Allow digits, decimal point, and backspace
        if (!char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != ',' && e.KeyChar != (char)Keys.Back)
        {
            e.Handled = true;
        }

        // Ensure only one decimal point exists
        if ((e.KeyChar == '.' || e.KeyChar == ',') && (sender as TextBox)?.Text.IndexOfAny(new char[] { '.', ',' }) > -1)
        {
            e.Handled = true;
        }
    }
}