using Application_Cyrell.LogiqueBouttonsSolidEdge;
using System.Windows.Forms;

public class SelectAllCommand : SolidEdgeCommandBase
{
    public SelectAllCommand(TextBox textBoxFolderPath, ListBox listBoxDxfFiles)
        : base(textBoxFolderPath, listBoxDxfFiles) { }

    public override void Execute()
    {
        if (_listBoxDxfFiles.Items.Count == 0)
        {
            MessageBox.Show("No files available to select.");
            return;
        }

        for (int i = 0; i < _listBoxDxfFiles.Items.Count; i++)
        {
            _listBoxDxfFiles.SetSelected(i, true);
        }
    }
}
