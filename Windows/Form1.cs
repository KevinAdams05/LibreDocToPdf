using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
        private bool isDarkMode;
        private bool isDebugMode;
        private string paperSize = "Letter";
        private readonly ConcurrentDictionary<int, byte> launchedPids = new ConcurrentDictionary<int, byte>();
        private static readonly string settingsPath = Path.Combine(AppContext.BaseDirectory, "settings.txt");
        private static readonly string profileRoot = Path.Combine(Path.GetTempPath(), "LibreDocToPdf");
        private static readonly string templateProfileDir = Path.Combine(profileRoot, "_template");
        private bool templateReady = false;
        private readonly SemaphoreSlim templateLock = new SemaphoreSlim(1, 1);

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        public Form1()
        {
            InitializeComponent();

            AllowDrop = true;
            DragEnter += Form1_DragEnter;
            DragDrop += Form1_DragDrop;

            LoadSettings();
            darkModeMenuItem.Checked = isDarkMode;
            debugModeMenuItem.Checked = isDebugMode;
            UpdatePaperSizeMenuChecks();
            ApplyTheme();

            FormClosing += Form1_FormClosing;

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

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using FolderBrowserDialog dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtFolder.Text = dlg.SelectedPath;
                Log($"Folder selected: {dlg.SelectedPath}");
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

        private void DebugLog(string message)
        {
            if (!isDebugMode)
            {
                return;
            }
            Log($"[DEBUG {DateTime.Now:HH:mm:ss.fff}] {message}");
        }

        private void KillTrackedLibreOffice()
        {
            int killed = 0;
            foreach (int pid in launchedPids.Keys)
            {
                try
                {
                    Process p = Process.GetProcessById(pid);
                    if (!p.HasExited)
                    {
                        p.Kill(true);
                        killed++;
                    }
                }
                catch { }
            }
            launchedPids.Clear();
            if (killed > 0)
            {
                Log($"Cleaned up {killed} tracked LibreOffice process(es)");
            }
        }

        private void Form1_FormClosing(object? sender, FormClosingEventArgs e)
        {
            cts?.Cancel();
            KillTrackedLibreOffice();
        }

        private async Task EnsureTemplateProfileAsync(CancellationToken token)
        {
            if (templateReady)
            {
                return;
            }
            await templateLock.WaitAsync(token);
            try
            {
                if (templateReady)
                {
                    return;
                }
                if (Directory.Exists(templateProfileDir) && Directory.Exists(Path.Combine(templateProfileDir, "user")))
                {
                    DebugLog($"Template profile already exists at {templateProfileDir}");
                    templateReady = true;
                    return;
                }
                Log("Initializing LibreOffice profile template (one-time)...");
                Directory.CreateDirectory(templateProfileDir);
                string userInstallation = "file:///" + templateProfileDir.Replace('\\', '/');
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = sofficePath,
                    Arguments = $"-env:UserInstallation=\"{userInstallation}\" --headless --norestore --nologo --nofirststartwizard --nodefault --nolockcheck --terminate_after_init",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using Process p = Process.Start(psi)!;
                launchedPids.TryAdd(p.Id, 0);
                DebugLog($"Template init START PID={p.Id}");
                Task<string> stdoutTask = p.StandardOutput.ReadToEndAsync();
                Task<string> stderrTask = p.StandardError.ReadToEndAsync();
                Task exitTask = p.WaitForExitAsync(token);
                Task completed = await Task.WhenAny(exitTask, Task.Delay(60000, token));
                if (completed != exitTask)
                {
                    try { p.Kill(true); } catch { }
                    DebugLog("Template init timed out after 60s, killed");
                }
                else
                {
                    await exitTask;
                    DebugLog($"Template init EXIT PID={p.Id} exitCode={p.ExitCode}");
                }
                launchedPids.TryRemove(p.Id, out _);
                templateReady = true;
            }
            finally
            {
                templateLock.Release();
            }
        }

        private static void CopyDirectory(string source, string dest)
        {
            Directory.CreateDirectory(dest);
            foreach (string dir in Directory.GetDirectories(source, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dir.Replace(source, dest));
            }
            foreach (string file in Directory.GetFiles(source, "*", SearchOption.AllDirectories))
            {
                File.Copy(file, file.Replace(source, dest), true);
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

            KillTrackedLibreOffice();

            cts = new CancellationTokenSource();
            CancellationToken token = cts.Token;

            try
            {
                await EnsureTemplateProfileAsync(token);
            }
            catch (OperationCanceledException)
            {
                Log("Operation cancelled.");
                return;
            }
            catch (Exception ex)
            {
                Log($"Template profile init failed: {ex.Message}");
                return;
            }

            SearchOption searchOption = chkRecursive.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            string[] allFiles = Directory.GetFiles(folder, "*.*", searchOption);
            List<string> files = allFiles
                .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                .ToList();

            List<string> filesToProcess = files
                .Where(f => !File.Exists(Path.Combine(customOutputFolder ?? Path.GetDirectoryName(f)!, Path.GetFileNameWithoutExtension(f) + ".pdf")))
                .ToList();

            totalFiles = filesToProcess.Count;
            processedFiles = 0;

            progressBar.Minimum = 0;
            progressBar.Maximum = totalFiles;
            progressBar.Value = 0;

            int concurrency = Math.Min(Environment.ProcessorCount, 8);
            Log($"Found {files.Count} files, {totalFiles} to convert.");
            DebugLog($"Concurrency limit = {concurrency}");

            SemaphoreSlim semaphore = new SemaphoreSlim(concurrency);
            List<Task> tasks = new List<Task>();
            int inFlight = 0;

            foreach (string file in filesToProcess)
            {
                string fileCaptured = file;
                Task task = Task.Run(async () =>
                {
                    DebugLog($"QUEUE waiting semaphore: {Path.GetFileName(fileCaptured)} (slots free={semaphore.CurrentCount})");
                    await semaphore.WaitAsync(token);
                    int current = Interlocked.Increment(ref inFlight);
                    DebugLog($"ACQUIRE semaphore: {Path.GetFileName(fileCaptured)} (inFlight={current})");
                    try
                    {
                        token.ThrowIfCancellationRequested();
                        await ConvertWithRetry(fileCaptured, token);
                    }
                    finally
                    {
                        int remaining = Interlocked.Decrement(ref inFlight);
                        semaphore.Release();
                        DebugLog($"RELEASE semaphore: {Path.GetFileName(fileCaptured)} (inFlight={remaining})");
                    }
                });

                tasks.Add(task);
            }

            try
            {
                await Task.WhenAll(tasks);

                Interlocked.Exchange(ref processedFiles, totalFiles);
                UpdateProgress();

                Log("All conversions completed.");
            }
            catch (OperationCanceledException)
            {
                Interlocked.Exchange(ref processedFiles, totalFiles);
                UpdateProgress();

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
            string expectedOutput = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(file) + ".pdf");
            string fname = Path.GetFileName(file);

            string userProfileDir = Path.Combine(profileRoot, Guid.NewGuid().ToString("N"));
            CopyDirectory(templateProfileDir, userProfileDir);
            WritePaperSizeXcu(userProfileDir);
            string userInstallation = "file:///" + userProfileDir.Replace('\\', '/');

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = sofficePath,
                Arguments = $"-env:UserInstallation=\"{userInstallation}\" --headless --norestore --nologo --nofirststartwizard --nodefault --nolockcheck --convert-to pdf --outdir \"{outputDir}\" \"{file}\"",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            Process? p = null;
            int pid = -1;
            Stopwatch sw = Stopwatch.StartNew();
            try
            {
                p = Process.Start(psi)!;
                pid = p.Id;
                launchedPids.TryAdd(pid, 0);
                DebugLog($"START  PID={pid} FILE={fname}");

                Task<string> stdoutTask = p.StandardOutput.ReadToEndAsync();
                Task<string> stderrTask = p.StandardError.ReadToEndAsync();

                await p.WaitForExitAsync(token);
                sw.Stop();

                string stdout = await stdoutTask;
                string stderr = await stderrTask;
                bool outputExists = File.Exists(expectedOutput);

                DebugLog($"EXIT   PID={pid} FILE={fname} exitCode={p.ExitCode} elapsed={sw.ElapsedMilliseconds}ms outputExists={outputExists}");

                if (!string.IsNullOrWhiteSpace(stdout))
                {
                    DebugLog($"  stdout PID={pid}: {stdout.Trim()}");
                }
                if (!string.IsNullOrWhiteSpace(stderr))
                {
                    DebugLog($"  stderr PID={pid}: {stderr.Trim()}");
                }

                if (outputExists)
                {
                    Interlocked.Increment(ref processedFiles);
                    Log($"Complete: {fname}");
                    UpdateProgress();
                    return true;
                }

                if (!string.IsNullOrWhiteSpace(stderr))
                {
                    Log($"LibreOffice error: {stderr.Trim()}");
                }

                if (!string.IsNullOrWhiteSpace(stdout) && stdout.Contains("Error", StringComparison.OrdinalIgnoreCase))
                {
                    Log($"LibreOffice: {stdout.Trim()}");
                }
            }
            catch (OperationCanceledException)
            {
                sw.Stop();
                DebugLog($"CANCEL PID={pid} FILE={fname} elapsed={sw.ElapsedMilliseconds}ms");
                try
                {
                    if (p != null && !p.HasExited)
                    {
                        p.Kill(true);
                        DebugLog($"  killed PID={pid} after cancel");
                    }
                }
                catch (Exception killEx)
                {
                    DebugLog($"  kill after cancel failed PID={pid}: {killEx.Message}");
                }
            }
            catch (Exception ex)
            {
                sw.Stop();
                DebugLog($"EXCEPT PID={pid} FILE={fname} elapsed={sw.ElapsedMilliseconds}ms: {ex.GetType().Name}: {ex.Message}");
                Log(ex.Message);
            }
            finally
            {
                if (pid != -1)
                {
                    launchedPids.TryRemove(pid, out _);
                }
                p?.Dispose();
                try
                {
                    Directory.Delete(userProfileDir, true);
                }
                catch { }
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
        private void aboutMenuItem_Click(object sender, EventArgs e)
        {
            using AboutForm about = new AboutForm(isDarkMode);
            about.ShowDialog(this);
        }

        private void darkModeMenuItem_Click(object sender, EventArgs e)
        {
            isDarkMode = darkModeMenuItem.Checked;
            ApplyTheme();
            SaveSettings();
        }

        private void debugModeMenuItem_Click(object sender, EventArgs e)
        {
            isDebugMode = debugModeMenuItem.Checked;
            Log($"Debug logging {(isDebugMode ? "enabled" : "disabled")}");
            SaveSettings();
        }

        private void paperSizeItem_Click(object? sender, EventArgs e)
        {
            if (sender == paperLetterMenuItem)
            {
                paperSize = "Letter";
            }
            else if (sender == paperA4MenuItem)
            {
                paperSize = "A4";
            }
            else if (sender == paperLegalMenuItem)
            {
                paperSize = "Legal";
            }
            UpdatePaperSizeMenuChecks();
            Log($"Default paper size set to {paperSize}");
            SaveSettings();
        }

        private void UpdatePaperSizeMenuChecks()
        {
            paperLetterMenuItem.Checked = paperSize == "Letter";
            paperA4MenuItem.Checked = paperSize == "A4";
            paperLegalMenuItem.Checked = paperSize == "Legal";
        }

        private (int width, int height) GetPaperDimensionsHundredthsMm()
        {
            return paperSize switch
            {
                "A4" => (21000, 29700),
                "Legal" => (21590, 35560),
                _ => (21590, 27940),
            };
        }

        private void WritePaperSizeXcu(string profileDir)
        {
            (int width, int height) = GetPaperDimensionsHundredthsMm();
            string userDir = Path.Combine(profileDir, "user");
            Directory.CreateDirectory(userDir);
            string xcuPath = Path.Combine(userDir, "registrymodifications.xcu");
            string xml = $@"<?xml version=""1.0"" encoding=""UTF-8""?>
<oor:items xmlns:oor=""http://openoffice.org/2001/registry"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
 <item oor:path=""/org.openoffice.Office.Writer/DefaultPageSize""><prop oor:name=""Width"" oor:op=""fuse""><value>{width}</value></prop></item>
 <item oor:path=""/org.openoffice.Office.Writer/DefaultPageSize""><prop oor:name=""Height"" oor:op=""fuse""><value>{height}</value></prop></item>
 <item oor:path=""/org.openoffice.Office.Common/Save/Document""><prop oor:name=""PrinterIndependentLayout"" oor:op=""fuse""><value>2</value></prop></item>
</oor:items>
";
            File.WriteAllText(xcuPath, xml);
        }

        private void ApplyTheme()
        {
            AppTheme theme = isDarkMode ? AppTheme.Dark : AppTheme.Light;

            int useDarkMode = isDarkMode ? 1 : 0;
            DwmSetWindowAttribute(Handle, 20, ref useDarkMode, sizeof(int));

            BackColor = theme.FormBack;
            ForeColor = theme.FormFore;

            menuStrip1.BackColor = theme.MenuBack;
            menuStrip1.ForeColor = theme.MenuFore;
            menuStrip1.Renderer = theme.GetMenuRenderer();

            foreach (ToolStripMenuItem item in menuStrip1.Items)
            {
                item.ForeColor = theme.MenuFore;
                foreach (ToolStripItem sub in item.DropDownItems)
                {
                    sub.ForeColor = theme.MenuFore;
                    if (sub is ToolStripMenuItem subMenu)
                    {
                        foreach (ToolStripItem child in subMenu.DropDownItems)
                        {
                            child.ForeColor = theme.MenuFore;
                        }
                    }
                }
            }

            txtFolder.BackColor = theme.ControlBack;
            txtFolder.ForeColor = theme.ControlFore;

            txtLog.BackColor = theme.ControlBack;
            txtLog.ForeColor = theme.ControlFore;

            label1.ForeColor = theme.FormFore;
            outputLogLabel.ForeColor = theme.FormFore;

            btnBrowse.BackColor = theme.ButtonBack;
            btnBrowse.ForeColor = theme.ButtonFore;
            btnBrowse.FlatAppearance.BorderColor = theme.ButtonBorder;

            btnConvert.BackColor = theme.AccentBack;
            btnConvert.ForeColor = theme.AccentFore;

            btnCancel.BackColor = theme.ButtonBack;
            btnCancel.ForeColor = theme.ButtonFore;
            btnCancel.FlatAppearance.BorderColor = theme.ButtonBorder;

            chkRecursive.ForeColor = theme.FormFore;

            progressBar.BarColor = theme.AccentBack;
            progressBar.BarBackColor = theme.ControlBack;
            progressBar.BorderColor = theme.ButtonBorder;
            progressBar.Invalidate();
        }

        private void LoadSettings()
        {
            isDarkMode = false;
            isDebugMode = false;
            try
            {
                if (!File.Exists(settingsPath))
                {
                    return;
                }
                string[] lines = File.ReadAllLines(settingsPath);
                if (lines.Length == 1 && !lines[0].Contains('='))
                {
                    isDarkMode = lines[0].Trim() == "dark";
                    return;
                }
                foreach (string line in lines)
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0)
                    {
                        continue;
                    }
                    string key = line.Substring(0, eq).Trim();
                    string val = line.Substring(eq + 1).Trim();
                    if (key == "theme")
                    {
                        isDarkMode = val == "dark";
                    }
                    else if (key == "debug")
                    {
                        isDebugMode = val == "true";
                    }
                    else if (key == "paper")
                    {
                        if (val == "Letter" || val == "A4" || val == "Legal")
                        {
                            paperSize = val;
                        }
                    }
                }
            }
            catch { }
        }

        private void SaveSettings()
        {
            try
            {
                string content = $"theme={(isDarkMode ? "dark" : "light")}{Environment.NewLine}debug={(isDebugMode ? "true" : "false")}{Environment.NewLine}paper={paperSize}{Environment.NewLine}";
                File.WriteAllText(settingsPath, content);
            }
            catch { }
        }
    }
}