using System;
using System.IO;
using System.Windows.Forms;
using Application_Cyrell.LogiqueBouttonsSolidEdge;

namespace Application_Cyrell.LogiqueBouttonsExcel
{
    public class parcourirDxfStepCommand : IButtonManager
    {
        private TextBox _textBox;

        public parcourirDxfStepCommand(TextBox txtBox)
        {
            _textBox = txtBox;
        }

        // Overloaded Execute method to accept a TextBox parameter
        public void Execute()
        {
            // Configure the OpenFileDialog to behave like a folder selector
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.ValidateNames = false; // Allow invalid file names
                openFileDialog.CheckFileExists = false; // Don't require the file to exist
                openFileDialog.CheckPathExists = true; // Ensure the path exists
                openFileDialog.FileName = "Folder Selection."; // Default text to indicate folder selection
                openFileDialog.Title = "Select a folder"; // Dialog title

                // Show the dialog and check if the user clicked OK
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the directory path from the selected file name
                    string selectedDirectory = Path.GetDirectoryName(openFileDialog.FileName);

                    // Update the TextBox with the selected directory path
                    _textBox.Text = selectedDirectory;
                }
            }
        }
    }
}