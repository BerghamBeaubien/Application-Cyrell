using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static Application_Cyrell.MainForm;

namespace Application_Cyrell.LogiqueBouttonsSolidEdge
{
    public class BrowseCommand : IButtonManager
    {
        private ListBox _listBoxDxfFiles;
        private TextBox _textBoxFolderPath;
        private Dictionary<string, bool> _extensionFilters;

        public BrowseCommand(ListBox listBoxDxfFiles, TextBox textBoxFolderPath, Dictionary<string, bool> extensionFilters)
        {
            _listBoxDxfFiles = listBoxDxfFiles;
            _textBoxFolderPath = textBoxFolderPath;
            _extensionFilters = extensionFilters;
        }

        public void Execute()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.ValidateNames = false;
                openFileDialog.CheckFileExists = false;
                openFileDialog.CheckPathExists = true;
                openFileDialog.FileName = "Folder Selection.";
                openFileDialog.Title = "Select a folder containing DXF files";

                _listBoxDxfFiles.Items.Clear();

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = Path.GetDirectoryName(openFileDialog.FileName);
                    _textBoxFolderPath.Text = selectedPath;

                    var activeExtensions = _extensionFilters
                    .Where(kv => kv.Value)
                    .Select(kv => kv.Key)
                    .ToList();

                    string[] allFiles = Directory.GetFiles(selectedPath, "*.*")
                    .Where(file => activeExtensions.Any(ext =>
                        file.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                    .ToArray();

                    Array.Sort(allFiles, FileSorter.CompareFileNames);

                    foreach (string file in allFiles)
                    {
                        _listBoxDxfFiles.Items.Add(Path.GetFileName(file));
                    }
                }
            }
        }
    }
}