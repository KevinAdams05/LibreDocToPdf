using System.Windows.Forms;

namespace LibreDocToPdf
{
    partial class Form1
    {
        private TextBox txtFolder;
        private Button btnConvert;
        private Button btnCancel;
        private ProgressBar progressBar;
        private TextBox txtLog;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem optionsMenu;
        private ToolStripMenuItem retryMenuItem;
        private ToolStripMenuItem outputFolderMenuItem;
        private ToolStripMenuItem exportLogMenuItem;

        private void InitializeComponent()
        {
            txtFolder = new TextBox();
            btnConvert = new Button();
            btnCancel = new Button();
            progressBar = new ProgressBar();
            txtLog = new TextBox();
            menuStrip1 = new MenuStrip();

            optionsMenu = new ToolStripMenuItem("Options");
            retryMenuItem = new ToolStripMenuItem("Retry Failed Files");
            outputFolderMenuItem = new ToolStripMenuItem("Set Output Folder");
            exportLogMenuItem = new ToolStripMenuItem("Export Log");

            retryMenuItem.Click += retryMenuItem_Click;
            outputFolderMenuItem.Click += outputFolderMenuItem_Click;
            exportLogMenuItem.Click += exportLogMenuItem_Click;

            optionsMenu.DropDownItems.Add(retryMenuItem);
            optionsMenu.DropDownItems.Add(outputFolderMenuItem);
            optionsMenu.DropDownItems.Add(exportLogMenuItem);

            menuStrip1.Items.Add(optionsMenu);

            txtFolder.SetBounds(10, 30, 500, 25);

            btnConvert.Text = "Convert";
            btnConvert.SetBounds(520, 30, 100, 25);
            btnConvert.Click += btnConvert_Click;

            btnCancel.Text = "Cancel";
            btnCancel.SetBounds(520, 60, 100, 25);
            btnCancel.Click += btnCancel_Click;

            progressBar.SetBounds(10, 60, 500, 20);

            txtLog.SetBounds(10, 90, 610, 300);
            txtLog.Multiline = true;
            txtLog.ScrollBars = ScrollBars.Vertical;

            Controls.Add(menuStrip1);
            Controls.Add(txtFolder);
            Controls.Add(btnConvert);
            Controls.Add(btnCancel);
            Controls.Add(progressBar);
            Controls.Add(txtLog);

            MainMenuStrip = menuStrip1;
            Text = "LibreOffice DOC → PDF Converter";
            ClientSize = new System.Drawing.Size(640, 420);
        }
    }
}