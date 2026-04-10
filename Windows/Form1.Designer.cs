using System.Windows.Forms;

namespace LibreDocToPdf
{
    partial class Form1
    {
        private TextBox txtFolder;
        private Button btnBrowse;
        private Button btnConvert;
        private Button btnCancel;
        private ThemedProgressBar progressBar;
        private TextBox txtLog;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem optionsMenu;
        private ToolStripMenuItem retryMenuItem;
        private ToolStripMenuItem outputFolderMenuItem;
        private ToolStripMenuItem exportLogMenuItem;
        private Label outputLogLabel;
        private Label label1;
        private ToolStripMenuItem helpMenu;
        private ToolStripMenuItem aboutMenuItem;
        private ToolStripMenuItem darkModeMenuItem;
        private CheckBox chkRecursive;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            txtFolder = new TextBox();
            btnBrowse = new Button();
            btnConvert = new Button();
            btnCancel = new Button();
            progressBar = new ThemedProgressBar();
            txtLog = new TextBox();
            menuStrip1 = new MenuStrip();
            optionsMenu = new ToolStripMenuItem();
            retryMenuItem = new ToolStripMenuItem();
            outputFolderMenuItem = new ToolStripMenuItem();
            exportLogMenuItem = new ToolStripMenuItem();
            helpMenu = new ToolStripMenuItem();
            aboutMenuItem = new ToolStripMenuItem();
            darkModeMenuItem = new ToolStripMenuItem();
            chkRecursive = new CheckBox();
            outputLogLabel = new Label();
            label1 = new Label();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            //
            // label1
            //
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(12, 33);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(73, 15);
            label1.TabIndex = 0;
            label1.Text = "Source Path:";
            //
            // txtFolder
            //
            txtFolder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtFolder.Location = new System.Drawing.Point(94, 30);
            txtFolder.Name = "txtFolder";
            txtFolder.Size = new System.Drawing.Size(276, 23);
            txtFolder.TabIndex = 1;
            //
            // btnBrowse
            //
            btnBrowse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnBrowse.FlatStyle = FlatStyle.Flat;
            btnBrowse.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(180, 180, 180);
            btnBrowse.Location = new System.Drawing.Point(376, 29);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new System.Drawing.Size(80, 25);
            btnBrowse.TabIndex = 2;
            btnBrowse.Text = "Browse...";
            btnBrowse.Click += btnBrowse_Click;
            //
            // btnConvert
            //
            btnConvert.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnConvert.BackColor = System.Drawing.Color.FromArgb(0, 120, 212);
            btnConvert.FlatAppearance.BorderSize = 0;
            btnConvert.FlatStyle = FlatStyle.Flat;
            btnConvert.ForeColor = System.Drawing.Color.White;
            btnConvert.Location = new System.Drawing.Point(462, 29);
            btnConvert.Name = "btnConvert";
            btnConvert.Size = new System.Drawing.Size(80, 25);
            btnConvert.TabIndex = 3;
            btnConvert.Text = "Convert";
            btnConvert.UseVisualStyleBackColor = false;
            btnConvert.Click += btnConvert_Click;
            //
            // btnCancel
            //
            btnCancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(180, 180, 180);
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.Location = new System.Drawing.Point(548, 29);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new System.Drawing.Size(80, 25);
            btnCancel.TabIndex = 4;
            btnCancel.Text = "Cancel";
            btnCancel.Click += btnCancel_Click;
            //
            // progressBar
            //
            progressBar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            progressBar.Location = new System.Drawing.Point(12, 62);
            progressBar.Name = "progressBar";
            progressBar.Size = new System.Drawing.Size(616, 23);
            progressBar.TabIndex = 5;
            //
            // chkRecursive
            //
            chkRecursive.AutoSize = true;
            chkRecursive.Checked = true;
            chkRecursive.CheckState = CheckState.Checked;
            chkRecursive.Location = new System.Drawing.Point(12, 92);
            chkRecursive.Name = "chkRecursive";
            chkRecursive.Size = new System.Drawing.Size(120, 19);
            chkRecursive.TabIndex = 6;
            chkRecursive.Text = "Include subfolders";
            //
            // outputLogLabel
            //
            outputLogLabel.AutoSize = true;
            outputLogLabel.Location = new System.Drawing.Point(12, 118);
            outputLogLabel.Name = "outputLogLabel";
            outputLogLabel.Size = new System.Drawing.Size(71, 15);
            outputLogLabel.TabIndex = 7;
            outputLogLabel.Text = "Output Log:";
            //
            // txtLog
            //
            txtLog.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            txtLog.Location = new System.Drawing.Point(12, 138);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Size = new System.Drawing.Size(616, 290);
            txtLog.TabIndex = 8;
            //
            // menuStrip1
            //
            menuStrip1.Items.AddRange(new ToolStripItem[] { optionsMenu, helpMenu });
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new System.Drawing.Size(640, 24);
            menuStrip1.TabIndex = 8;
            //
            // optionsMenu
            //
            optionsMenu.DropDownItems.AddRange(new ToolStripItem[] { retryMenuItem, outputFolderMenuItem, exportLogMenuItem, new ToolStripSeparator(), darkModeMenuItem });
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
            // darkModeMenuItem
            //
            darkModeMenuItem.Name = "darkModeMenuItem";
            darkModeMenuItem.Size = new System.Drawing.Size(180, 22);
            darkModeMenuItem.Text = "Dark Mode";
            darkModeMenuItem.CheckOnClick = true;
            darkModeMenuItem.Click += darkModeMenuItem_Click;
            //
            // helpMenu
            //
            helpMenu.DropDownItems.AddRange(new ToolStripItem[] { aboutMenuItem });
            helpMenu.Name = "helpMenu";
            helpMenu.Size = new System.Drawing.Size(44, 20);
            helpMenu.Text = "Help";
            //
            // aboutMenuItem
            //
            aboutMenuItem.Name = "aboutMenuItem";
            aboutMenuItem.Size = new System.Drawing.Size(180, 22);
            aboutMenuItem.Text = "About";
            aboutMenuItem.Click += aboutMenuItem_Click;
            //
            // Form1
            //
            ClientSize = new System.Drawing.Size(640, 440);
            Controls.Add(label1);
            Controls.Add(txtFolder);
            Controls.Add(btnBrowse);
            Controls.Add(btnConvert);
            Controls.Add(btnCancel);
            Controls.Add(progressBar);
            Controls.Add(chkRecursive);
            Controls.Add(outputLogLabel);
            Controls.Add(txtLog);
            Controls.Add(menuStrip1);
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            MinimumSize = new System.Drawing.Size(520, 350);
            Name = "Form1";
            Text = "DOC to PDF Converter";
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }
    }
}
