﻿using Microsoft.Office.Interop.Excel;
using SolidEdgeDraft;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Channels;
using System.Windows.Forms;

namespace Application_Cyrell
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                System.Windows.Forms.Application.EnableVisualStyles();
                System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
                System.Windows.Forms.Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Startup error:\n" + ex.ToString(), "Error");
            }
        }
    }
}