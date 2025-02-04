using System;
using System.Windows.Forms; // Required for OpenFileDialog and TextBox
using Application_Cyrell.LogiqueBouttonsSolidEdge;

namespace Application_Cyrell.LogiqueBouttonsExcel
{
    public class parcourirXlCommand : IButtonManager
    {
        private TextBox _textBox;

        public parcourirXlCommand(TextBox txtBox)
        {
            _textBox = txtBox;
        }

        // Overloaded Execute method to accept a TextBox parameter
        public void Execute()
        {
            // Create an instance of OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the filter to restrict the file types that can be selected
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|All Files|*.*";

            openFileDialog.Title = "Veuillez choisir un fichier Excel";

            // Show the dialog and check if the user clicked OK
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Get the selected file's name (including the path)
                string selectedFileName = openFileDialog.FileName;

                // Update the TextBox with the selected file's name
                _textBox.Text = selectedFileName;
            }
        }
    }
}