using Application_Cyrell.LogiqueBouttonsSolidEdge;
using SolidEdgeCommunity;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class KillSECommand : SolidEdgeCommandBase
{
    public KillSECommand(TextBox textBoxFolderPath, ListBox listBoxDxfFiles)
        : base(textBoxFolderPath, listBoxDxfFiles) { }

    public override void Execute()
    {
        SolidEdgeFramework.Application seApp = null;

        // Formulaire de confirmation
        DialogResult result = MessageBox.Show("Cette opération va fermer Solid Edge et \nles documents ne seront pas sauvgardés \n\n" +
            "Voulez-vous continuer?", "Message de Confirmation", MessageBoxButtons.YesNo);

        if (result == DialogResult.Yes)
        {
            try
            {
                // Check if Solid Edge is running
                Process[] solidEdgeProcesses = Process.GetProcessesByName("Edge"); // "Edge" is the Solid Edge process name

                if (solidEdgeProcesses.Length == 0)
                {
                    MessageBox.Show("Aucune instance de Solid Edge n'est en cours d'exécution.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Connect to an existing instance of Solid Edge
                seApp = SolidEdgeUtils.Connect(true);
                seApp.Visible = true;

                // Get the Process ID of the Solid Edge application
                int pid = seApp.ProcessID;

                // Find the process using the PID and kill it
                Process process = Process.GetProcessById(pid);
                process.Kill(); // Terminate the Solid Edge process

                MessageBox.Show("Solid Edge a été tué avec succès");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error closing Solid Edge: " + ex.Message);
            }
            finally
            {
                if (seApp != null)
                {
                    Marshal.ReleaseComObject(seApp);
                    seApp = null;
                }
            }
        }
    }
}
