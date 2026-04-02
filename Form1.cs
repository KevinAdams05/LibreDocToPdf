using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LibreDocToPdf
{
    public partial class Form1 : Form
    {
        private string sofficePath;
        private int totalFiles = 0;
        private int processedFiles = 0;
        private int retryCount = 2;
        private string? customOutputFolder = null;
        private string logFilePath;
        private CancellationTokenSource? cts;

        public Form1()
        {
            InitializeComponent();

            AllowDrop = true;
            DragEnter += Form1_DragEnter;
            DragDrop += Form1_DragDrop;

            sofficePath = DetectLibreOffice();

            Directory.CreateDirectory("logs");
            logFilePath = Path.Combine("logs", $"log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

            Log($"LibreOffice Path: {sofficePath}");
        }

        private string DetectLibreOffice()
        {
            string[] paths =
            {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            };

            foreach (var p in paths)
            {
                if (File.Exists(p))
                {
                    return p;
                }
            }

            var env = Environment.GetEnvironmentVariable("PATH")?.Split(';');
            if (env != null)
            {
                foreach (var p in env)
                {
                    try
                    {
                        var full = Path.Combine(p, "soffice.exe");
                        if (File.Exists(full))
                        {
                            return full;
                        }
                    }
                    catch { }
                }
            }

            MessageBox.Show("LibreOffice not found.");
            return "soffice.exe";
        }

        private void Log(string msg)
        {
            string line = $"{DateTime.Now:HH:mm:ss} - {msg}";

            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(() => txtLog.AppendText(line + Environment.NewLine));
            }
            else
            {
                txtLog.AppendText(line + Environment.NewLine);
            }

            File.AppendAllText(logFilePath, line + Environment.NewLine);
        }

        private void UpdateProgress()
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(() => progressBar.Value = processedFiles);
            }
            else
            {
                progressBar.Value = processedFiles;
            }
        }

        private void Form1_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void Form1_DragDrop(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is not string[] paths || paths.Length == 0)
            {
                return;
            }

            if (Directory.Exists(paths[0]))
            {
                txtFolder.Text = paths[0];
                Log($"Folder dropped: {paths[0]}");
            }
        }

        private async void btnConvert_Click(object sender, EventArgs e)
        {
            string folder = txtFolder.Text;

            if (!Directory.Exists(folder))
            {
                MessageBox.Show("Invalid folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            cts = new CancellationTokenSource();
            var token = cts.Token;

            var files = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                .ToList();

            totalFiles = files.Count;
            processedFiles = 0;

            progressBar.Maximum = totalFiles;
            progressBar.Value = 0;

            Log($"Found {totalFiles} files.");

            using SemaphoreSlim semaphore = new SemaphoreSlim(Environment.ProcessorCount);

            var tasks = files.Select(async file =>
            {
                await semaphore.WaitAsync(token);
                try
                {
                    token.ThrowIfCancellationRequested();
                    await ConvertWithRetry(file, token);
                }
                finally
                {
                    semaphore.Release();
                }
            });

            try
            {
                await Task.WhenAll(tasks);
                Log("All conversions completed.");
            }
            catch (OperationCanceledException)
            {
                Log("Operation cancelled.");
            }
        }

        private async Task ConvertWithRetry(string file, CancellationToken token)
        {
            for (int i = 1; i <= retryCount + 1; i++)
            {
                token.ThrowIfCancellationRequested();

                if (await ConvertToPdf(file, token))
                {
                    return;
                }

                Log($"Retry {i}/{retryCount} for {Path.GetFileName(file)}");
            }

            Log($"Failed: {file}");
        }

        private async Task<bool> ConvertToPdf(string file, CancellationToken token)
        {
            string outputDir = customOutputFolder ?? Path.GetDirectoryName(file)!;

            var psi = new ProcessStartInfo
            {
                FileName = sofficePath,
                Arguments = $"--headless --convert-to pdf --outdir \"{outputDir}\" \"{file}\"",
                CreateNoWindow = true,
                UseShellExecute = false
            };

            try
            {
                using var p = Process.Start(psi)!;
                await p.WaitForExitAsync(token);

                if (p.ExitCode == 0)
                {
                    Interlocked.Increment(ref processedFiles);
                    Log($"Complete: {Path.GetFileName(file)}");
                    UpdateProgress();
                    return true;
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                Log(ex.Message);
            }

            return false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            cts?.Cancel();
            Log("Cancel requested...");
        }

        private void retryMenuItem_Click(object sender, EventArgs e)
        {
            string input = Microsoft.VisualBasic.Interaction.InputBox("Retry count:", "Options", retryCount.ToString());

            if (int.TryParse(input, out int val))
            {
                retryCount = val;
                Log($"Retry set to {retryCount}");
            }
        }

        private void outputFolderMenuItem_Click(object sender, EventArgs e)
        {
            using FolderBrowserDialog dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                customOutputFolder = dlg.SelectedPath;
                Log($"Output folder: {customOutputFolder}");
            }
        }

        private void exportLogMenuItem_Click(object sender, EventArgs e)
        {
            using SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text Files (*.txt)|*.txt";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                File.Copy(logFilePath, sfd.FileName, true);
                MessageBox.Show("Log exported.");
            }
        }
    }
}