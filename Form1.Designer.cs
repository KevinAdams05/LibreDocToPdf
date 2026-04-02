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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            txtFolder = new TextBox();
            btnConvert = new Button();
            btnCancel = new Button();
            progressBar = new ProgressBar();
            txtLog = new TextBox();
            menuStrip1 = new MenuStrip();
            optionsMenu = new ToolStripMenuItem();
            retryMenuItem = new ToolStripMenuItem();
            outputFolderMenuItem = new ToolStripMenuItem();
            exportLogMenuItem = new ToolStripMenuItem();
            outputLogLabel = new Label();
            label1 = new Label();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // txtFolder
            // 
            txtFolder.Location = new System.Drawing.Point(94, 30);
            txtFolder.Name = "txtFolder";
            txtFolder.Size = new System.Drawing.Size(416, 23);
            txtFolder.TabIndex = 1;
            // 
            // btnConvert
            // 
            btnConvert.Location = new System.Drawing.Point(10, 86);
            btnConvert.Name = "btnConvert";
            btnConvert.Size = new System.Drawing.Size(100, 25);
            btnConvert.TabIndex = 2;
            btnConvert.Text = "Convert";
            btnConvert.Click += btnConvert_Click;
            // 
            // btnCancel
            // 
            btnCancel.Location = new System.Drawing.Point(116, 86);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new System.Drawing.Size(100, 25);
            btnCancel.TabIndex = 3;
            btnCancel.Text = "Cancel";
            btnCancel.Click += btnCancel_Click;
            // 
            // progressBar
            // 
            progressBar.Location = new System.Drawing.Point(12, 60);
            progressBar.Name = "progressBar";
            progressBar.Size = new System.Drawing.Size(498, 20);
            progressBar.TabIndex = 4;
            // 
            // txtLog
            // 
            txtLog.Location = new System.Drawing.Point(10, 152);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Size = new System.Drawing.Size(500, 256);
            txtLog.TabIndex = 5;
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { optionsMenu });
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new System.Drawing.Size(517, 24);
            menuStrip1.TabIndex = 0;
            // 
            // optionsMenu
            // 
            optionsMenu.DropDownItems.AddRange(new ToolStripItem[] { retryMenuItem, outputFolderMenuItem, exportLogMenuItem });
            optionsMenu.Name = "optionsMenu";
            optionsMenu.Size = new System.Drawing.Size(61, 20);
            optionsMenu.Text = "Options";
            // 
            // retryMenuItem
            // 
            retryMenuItem.Name = "retryMenuItem";
            retryMenuItem.Size = new System.Drawing.Size(180, 22);
            retryMenuItem.Text = "Retry Failed Files";
            retryMenuItem.Click += retryMenuItem_Click;
            // 
            // outputFolderMenuItem
            // 
            outputFolderMenuItem.Name = "outputFolderMenuItem";
            outputFolderMenuItem.Size = new System.Drawing.Size(180, 22);
            outputFolderMenuItem.Text = "Set Output Folder";
            outputFolderMenuItem.Click += outputFolderMenuItem_Click;
            // 
            // exportLogMenuItem
            // 
            exportLogMenuItem.Name = "exportLogMenuItem";
            exportLogMenuItem.Size = new System.Drawing.Size(180, 22);
            exportLogMenuItem.Text = "Export Log";
            exportLogMenuItem.Click += exportLogMenuItem_Click;
            // 
            // outputLogLabel
            // 
            outputLogLabel.AccessibleDescription = "outputLogLabel";
            outputLogLabel.AccessibleName = "outputLogLabel";
            outputLogLabel.AutoSize = true;
            outputLogLabel.Location = new System.Drawing.Point(10, 134);
            outputLogLabel.Name = "outputLogLabel";
            outputLogLabel.Size = new System.Drawing.Size(71, 15);
            outputLogLabel.TabIndex = 6;
            outputLogLabel.Text = "Output Log:";
            // 
            // label1
            // 
            label1.AccessibleDescription = "outputLogLabel";
            label1.AccessibleName = "outputLogLabel";
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(10, 33);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(73, 15);
            label1.TabIndex = 7;
            label1.Text = "Source Path:";
            // 
            // Form1
            // 
            ClientSize = new System.Drawing.Size(517, 420);
            Controls.Add(label1);
            Controls.Add(outputLogLabel);
            Controls.Add(menuStrip1);
            Controls.Add(txtFolder);
            Controls.Add(btnConvert);
            Controls.Add(btnCancel);
            Controls.Add(progressBar);
            Controls.Add(txtLog);
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            Name = "Form1";
            Text = "DOC to PDF Converter";
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }
        private Label outputLogLabel;
        private Label label1;
    }
}