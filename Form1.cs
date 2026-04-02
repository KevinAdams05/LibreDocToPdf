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
            string[] paths = {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            };

            foreach (string p in paths)
            {
                if (File.Exists(p))
                {
                    return p;
                }
            }

            string[]? env = Environment.GetEnvironmentVariable("PATH")?.Split(';');
            if (env != null)
            {
                foreach (string p in env)
                {
                    try
                    {
                        string full = Path.Combine(p, "soffice.exe");
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

        private void KillStuckLibreOffice()
        {
            Process[] processes = Process.GetProcessesByName("soffice");
            foreach (Process p in processes)
            {
                try
                {
                    if ((DateTime.Now - p.StartTime).TotalMinutes > 5)
                    {
                        p.Kill();
                        Log($"Killed stuck LibreOffice process (PID {p.Id})");
                    }
                }
                catch { }
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

            // Kill stuck LibreOffice processes before starting
            KillStuckLibreOffice();

            cts = new CancellationTokenSource();
            CancellationToken token = cts.Token;

            string[] allFiles = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories);
            System.Collections.Generic.List<string> files = allFiles
                .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                .ToList();

            totalFiles = files.Count;
            processedFiles = 0;

            progressBar.Maximum = totalFiles;
            progressBar.Value = 0;

            Log($"Found {totalFiles} files.");

            SemaphoreSlim semaphore = new SemaphoreSlim(Environment.ProcessorCount);
            System.Collections.Generic.List<Task> tasks = new System.Collections.Generic.List<Task>();

            foreach (string file in files)
            {
                Task task = Task.Run(async () => {
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

                tasks.Add(task);
            }

            try
            {
                await Task.WhenAll(tasks);
                progressBar.Value = 100;
                Log("All conversions completed.");
            }
            catch (OperationCanceledException)
            {
                progressBar.Value = 100;
                Log("Operation cancelled.");
            }
        }

        private async Task ConvertWithRetry(string file, CancellationToken token)
        {
            string outputFile = Path.Combine(customOutputFolder ?? Path.GetDirectoryName(file)!, Path.GetFileNameWithoutExtension(file) + ".pdf");

            if (File.Exists(outputFile))
            {
                Log($"Skipping already converted: {Path.GetFileName(file)}");
                Interlocked.Increment(ref processedFiles);
                UpdateProgress();

                return;
            }

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

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = sofficePath,
                Arguments = $"--headless --convert-to pdf --outdir \"{outputDir}\" \"{file}\"",
                CreateNoWindow = true,
                UseShellExecute = false
            };

            try
            {
                Process p = Process.Start(psi)!;
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