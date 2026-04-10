using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace LibreDocToPdf
{
    public class AboutForm : Form
    {
        public AboutForm(bool darkMode = false)
        {
            Text = "About DOC to PDF Converter";
            ClientSize = new Size(420, 310);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;

            AppTheme theme = darkMode ? AppTheme.Dark : AppTheme.Light;
            BackColor = theme.FormBack;
            ForeColor = theme.FormFore;

            string version = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0";

            int textLeft = 88;

            var picIcon = new PictureBox
            {
                Size = new Size(56, 56),
                Location = new Point(20, 20),
                SizeMode = PictureBoxSizeMode.Zoom
            };

            string iconPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "icon2.ico");
            if (File.Exists(iconPath))
            {
                picIcon.Image = new Icon(iconPath, 48, 48).ToBitmap();
            }
            else if (Owner?.Icon != null)
            {
                picIcon.Image = Owner.Icon.ToBitmap();
            }

            var lblTitle = new Label
            {
                Text = "DOC to PDF Converter",
                Font = new Font("Segoe UI", 14F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(textLeft, 20)
            };

            var lblVersion = new Label
            {
                Text = $"Version {version}",
                Font = new Font("Segoe UI", 10F),
                AutoSize = true,
                Location = new Point(textLeft, 50)
            };

            var lblAuthor = new Label
            {
                Text = "Author: Kevin Adams",
                AutoSize = true,
                Location = new Point(20, 90)
            };

            var lnkGitHub = new LinkLabel
            {
                Text = "github.com/KevinAdams05/LibreDocToPdf",
                AutoSize = true,
                Location = new Point(20, 110)
            };
            lnkGitHub.LinkClicked += (s, e) =>
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://github.com/KevinAdams05/LibreDocToPdf",
                    UseShellExecute = true
                });
            };

            var lblCredits = new Label
            {
                Text = "Credits:\n\n" +
                       "  Icon: icon-icons.com (pdf-reader-pro-macos-bigsur)\n\n" +
                       "  Powered by LibreOffice (libreoffice.org)\n\n" +
                       "Licensed under the MIT License.",
                AutoSize = true,
                Location = new Point(20, 145),
                MaximumSize = new Size(380, 0)
            };

            lnkGitHub.LinkColor = darkMode ? Color.FromArgb(100, 180, 255) : Color.Blue;

            var btnOk = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                FlatStyle = FlatStyle.Flat,
                BackColor = theme.ButtonBack,
                ForeColor = theme.ButtonFore,
                Size = new Size(80, 28),
                Location = new Point(320, 270)
            };
            btnOk.FlatAppearance.BorderColor = theme.ButtonBorder;

            AcceptButton = btnOk;
            Controls.AddRange(new Control[] { picIcon, lblTitle, lblVersion, lblAuthor, lnkGitHub, lblCredits, btnOk });
        }
    }
}
