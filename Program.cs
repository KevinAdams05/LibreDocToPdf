using System;
using System.Windows.Forms;

namespace LibreDocToPdf
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            // Icon credit goes to: https://icon-icons.com/icon/pdf-reader-pro-macos-bigsur/189857

            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}