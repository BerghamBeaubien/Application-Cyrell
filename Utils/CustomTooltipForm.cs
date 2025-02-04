using System;
using System.Drawing;
using System.Windows.Forms;

public class CustomTooltipForm : Form
{
    private PictureBox pictureBox;
    private Label titleLabel;
    private Label overlayLabel;
    private TableLayoutPanel layoutPanel;

    public CustomTooltipForm(string title, string overlayText, Image tooltipImage)
    {
        // Initialize the form
        this.FormBorderStyle = FormBorderStyle.None;
        this.StartPosition = FormStartPosition.Manual;
        this.BackColor = Color.White;
        this.TopMost = true;
        this.AutoSize = true;
        this.AutoSizeMode = AutoSizeMode.GrowAndShrink;

        // Set a minimum size for the form
        this.MinimumSize = new Size(300, 200); // Default minimum width and height

        // Create a TableLayoutPanel
        layoutPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1, // Single column
            RowCount = 3,    // Title, Image, Overlay
            AutoSize = true
        };
        this.Controls.Add(layoutPanel);

        // Adjust RowStyles for better layout
        layoutPanel.RowStyles.Clear();
        layoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // Title row height
        layoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // Image row
        layoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // Overlay row height

        // Title label
        titleLabel = new Label
        {
            Text = title,
            Font = new Font("Arial", 12, FontStyle.Bold),
            ForeColor = Color.Black,
            TextAlign = ContentAlignment.MiddleCenter,
            Dock = DockStyle.Fill, // Fill the space in its cell
            Padding = new Padding(10),
            Visible = true // Ensure it's visible
        };
        layoutPanel.Controls.Add(titleLabel, 0, 0); // Add to the first row

        // PictureBox for the GIF
        pictureBox = new PictureBox
        {
            Image = tooltipImage,
            SizeMode = PictureBoxSizeMode.AutoSize,
            Dock = DockStyle.Fill // Ensure it stretches if needed
        };
        layoutPanel.Controls.Add(pictureBox, 0, 1); // Add to the second row

        // Overlay label
        overlayLabel = new Label
        {
            Text = overlayText,
            Font = new Font("Arial", 10, FontStyle.Regular),
            ForeColor = Color.Black,
            TextAlign = ContentAlignment.MiddleCenter,
            Dock = DockStyle.Fill, // Fill the space in its cell
            Padding = new Padding(10)
        };
        layoutPanel.Controls.Add(overlayLabel, 0, 2); // Add to the third row

        // Dynamically calculate the minimum size to accommodate all elements
        CalculateMinimumSize();
    }

    private void CalculateMinimumSize()
    {
        // Measure sizes of all components
        int titleHeight = titleLabel.PreferredHeight;
        int overlayHeight = overlayLabel.PreferredHeight;
        int imageWidth = pictureBox.Image != null ? pictureBox.Image.Width : 0;
        int imageHeight = pictureBox.Image != null ? pictureBox.Image.Height : 0;

        // Calculate minimum width and height
        int minWidth = Math.Max(300, imageWidth + 20); // Add padding to the image width
        int minHeight = Math.Max(200, titleHeight + imageHeight + overlayHeight + 40); // Add padding for spacing

        // Set the form's minimum size
        this.MinimumSize = new Size(minWidth, minHeight);
    }
}
