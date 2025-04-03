using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Application_Cyrell.Utils;

public class FormulaireDxfStep : Form
{
    private TextBox txtOutputPath;
    private CheckBox chkTagDxf;
    private CheckBox chkMacroDen;
    private CheckBox chkChangeName;
    private CheckBox chkFabbrica;
    private CheckBox chkSingleFile;
    private Button btnBrowseOutput;
    private Button btnContinue;
    private Button btnCancel;
    private BouttonToggle fileTypeToggle;
    private Label lblDxfOption;
    private Label lblStepOption;

    public string OutputPath => txtOutputPath.Text;
    public bool TagDxf => chkTagDxf.Checked;
    public bool ChangeName => chkChangeName.Checked;
    public bool Fabbrica => chkFabbrica.Checked && chkChangeName.Checked;
    public bool MacroDen => chkMacroDen.Checked;
    public bool OnlyDxf => chkSingleFile.Checked && !fileTypeToggle.Checked;
    public bool OnlyStep => chkSingleFile.Checked && fileTypeToggle.Checked;

    public FormulaireDxfStep()
    {
        InitializeComponents();
        this.fileTypeToggle.Visible = false;
    }

    private void InitializeComponents()
    {
        this.Text = "Sélection du répertoire de sortie";
        this.Size = new Size(600, 280); // Increased height for new controls
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.Font = new Font("Segoe UI", 9F);

        // Consistent margins and spacing
        int leftMargin = 20;
        int verticalSpacing = 10;
        int controlHeight = 25;

        // Output Path Controls
        Label lblOutput = new Label
        {
            Text = "Répertoire de sortie (DXF et STEP):",
            Location = new Point(leftMargin, 15),
            AutoSize = true
        };

        txtOutputPath = new TextBox
        {
            Location = new Point(leftMargin, lblOutput.Bottom + verticalSpacing),
            Width = this.ClientSize.Width - (leftMargin * 2 + 40),
            Height = controlHeight,
            ReadOnly = true
        };

        btnBrowseOutput = new Button
        {
            Text = "...",
            Location = new Point(txtOutputPath.Right + 10, txtOutputPath.Top - 1),
            Width = 30,
            Height = controlHeight
        };
        btnBrowseOutput.Click += (s, e) => BrowseFolder(txtOutputPath);

        // First row of options
        int firstRowTop = txtOutputPath.Bottom + verticalSpacing * 2;
        int checkboxSpacing = 20;

        chkTagDxf = new CheckBox
        {
            Text = "Tag DXF",
            Location = new Point(leftMargin, firstRowTop),
            AutoSize = true
        };

        chkMacroDen = new CheckBox
        {
            Text = "Executer Macro(DenMar)",
            Location = new Point(chkTagDxf.Right + checkboxSpacing, firstRowTop),
            AutoSize = true
        };

        // Second row of options
        int secondRowTop = firstRowTop + 30;

        chkChangeName = new CheckBox
        {
            Text = "Changer le nom",
            Location = new Point(leftMargin, secondRowTop),
            AutoSize = true
        };

        chkFabbrica = new CheckBox
        {
            Text = "Fabbrica",
            Location = new Point(chkChangeName.Right + checkboxSpacing, secondRowTop),
            AutoSize = true,
            Visible = true
        };

        // Third row - Single file option
        int thirdRowTop = secondRowTop + 30;

        chkSingleFile = new CheckBox
        {
            Text = "Générer un seul type de fichier",
            Location = new Point(leftMargin, thirdRowTop),
            AutoSize = true
        };
        chkSingleFile.CheckedChanged += (s, e) => UpdateFileTypeVisibility();

        int fourthRowTop = thirdRowTop + 30;

        // File type toggle (hidden by default)
        fileTypeToggle = new BouttonToggle()
        {
            Location = new Point(leftMargin + 100, fourthRowTop),
            Size = new Size(55, 25),
            OnBackColor = Color.Coral, // Different color from medium slate blue
            OffBackColor = Color.Coral,
            Checked = false
        };

        lblDxfOption = new Label()
        {
            Text = "DXF",
            Location = new Point(fileTypeToggle.Left - 50, fourthRowTop + 3),
            AutoSize = true,
            Visible = false
        };

        lblStepOption = new Label()
        {
            Text = "STEP",
            Location = new Point(fileTypeToggle.Right + 10, fourthRowTop + 3),
            AutoSize = true,
            Visible = false
        };

        // Buttons
        int buttonWidth = 80;
        int buttonHeight = 30;
        int buttonBottomMargin = 20;

        btnContinue = new Button
        {
            Text = "Continuer",
            DialogResult = DialogResult.OK,
            Location = new Point(this.ClientSize.Width - (buttonWidth * 2 + 30), thirdRowTop + 50),
            Width = buttonWidth,
            Height = buttonHeight,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.Coral,
            ForeColor = Color.White
        };
        btnContinue.FlatAppearance.BorderSize = 0;

        btnCancel = new Button
        {
            Text = "Annuler",
            DialogResult = DialogResult.Cancel,
            Location = new Point(btnContinue.Right + 10, btnContinue.Top),
            Width = buttonWidth,
            Height = buttonHeight,
            FlatStyle = FlatStyle.Flat
        };
        btnCancel.FlatAppearance.BorderSize = 0;

        // Form validation
        FormClosing += (s, e) =>
        {
            if (this.DialogResult == DialogResult.OK && string.IsNullOrWhiteSpace(txtOutputPath.Text))
            {
                MessageBox.Show("Veuillez sélectionner un répertoire de sortie pour continuer.",
                               "Répertoire requis",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Warning);
                e.Cancel = true;
            }
        };

        // Add controls
        this.Controls.AddRange(new Control[] {
            lblOutput, txtOutputPath, btnBrowseOutput,
            chkTagDxf, chkMacroDen, chkChangeName, chkFabbrica,
            chkSingleFile, fileTypeToggle, lblDxfOption, lblStepOption,
            btnContinue, btnCancel
        });

        this.AcceptButton = btnContinue;
        this.CancelButton = btnCancel;
    }

    private void UpdateFileTypeVisibility()
    {
        bool visible = chkSingleFile.Checked;
        fileTypeToggle.Visible = visible;
        lblDxfOption.Visible = visible;
        lblStepOption.Visible = visible;
    }

    private void BrowseFolder(TextBox textBox)
    {
        using (OpenFileDialog dialog = new OpenFileDialog())
        {
            dialog.CheckFileExists = false;
            dialog.CheckPathExists = true;
            dialog.ValidateNames = false;
            dialog.FileName = "Folder Selection."; // Placeholder text

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string selectedPath = Path.GetDirectoryName(dialog.FileName);
                textBox.Text = selectedPath;
            }
        }
    }
}