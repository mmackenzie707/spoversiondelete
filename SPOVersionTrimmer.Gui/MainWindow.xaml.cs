
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

namespace SPOVersionTrimmer.Gui
{
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // Default report path
            var safe = "VersionTrim_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv";
            ReportPathBox.Text = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                safe
            );
        }

        private void Log(string msg)
        {
            LogBox.AppendText($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}\r\n");
            LogBox.ScrollToEnd();
        }

        private string RunnerPath =>
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PowerShell", "Runner.ps1");

        private async Task<string> RunPwshAsync(string args)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "pwsh.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{RunnerPath}\" {args}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            var p = new Process { StartInfo = psi };
            var output = new StringBuilder();
            var error = new StringBuilder();

            p.OutputDataReceived += (_, e) => { if (e.Data != null) output.AppendLine(e.Data); };
            p.ErrorDataReceived += (_, e) => { if (e.Data != null) error.AppendLine(e.Data); };

            p.Start();
            p.BeginOutputReadLine();
            p.BeginErrorReadLine();
            await p.WaitForExitAsync();

            if (p.ExitCode != 0)
                throw new Exception(error.ToString().Trim());

            return output.ToString();
        }
        
        private void SetBusy(bool busy, string status)
        {
            BusyBar.IsIndeterminate = busy;
            BusyBar.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            StatusText.Text = status;
            ConnectBtn.IsEnabled = !busy;
            
            // Only enable Run if already enabled by successful connect
            RunBtn.IsEnabled = !busy && RunBtn.IsEnabled;
        }

        private async void ConnectBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SetBusy(true, "Connecting...");
                Log("Connecting using Device Login...");

                var args =
                    $"-Action TestConnect " +
                    $"-SiteUrl \"{SiteUrlBox.Text}\" " +
                    $"-ClientId \"{ClientIdBox.Text}\" " +
                    $"-TenantId \"{TenantIdBox.Text}\"";

                var json = await RunPwshAsync(args);
                using var doc = JsonDocument.Parse(json);

                if (!doc.RootElement.GetProperty("ok").GetBoolean())
                {
                    var err = doc.RootElement.GetProperty("error").GetString();
                    Log($"ERROR: {err}");
                    MessageBox.Show(err, "Connect Failed");
                    return;
                }

                var title = doc.RootElement.GetProperty("title").GetString();
                var url = doc.RootElement.GetProperty("url").GetString();
                Log($"Connected to: {title} ({url})");

                // Load libraries
                Log("Loading document libraries...");
                var libsJson = await RunPwshAsync(
                    $"-Action ListLibraries -SiteUrl \"{SiteUrlBox.Text}\" -ClientId \"{ClientIdBox.Text}\" -TenantId \"{TenantIdBox.Text}\""
                );

                using var libsDoc = JsonDocument.Parse(libsJson);
                if (!libsDoc.RootElement.GetProperty("ok").GetBoolean())
                {
                    var err = libsDoc.RootElement.GetProperty("error").GetString();
                    Log($"ERROR loading libraries: {err}");
                    MessageBox.Show(err, "Library Load Failed");
                    return;
                }

                LibraryCombo.Items.Clear();
                foreach (var lib in libsDoc.RootElement.GetProperty("libraries").EnumerateArray())
                {
                    LibraryCombo.Items.Add(lib.GetProperty("Title").GetString());
                }

                if (LibraryCombo.Items.Count > 0) LibraryCombo.SelectedIndex = 0;

                RunBtn.IsEnabled = true;
                SetBusy(false, "Connected");
            }
            catch (Exception ex)
            {
                SetBusy(false, "Error");
                Log("ERROR: " + ex.Message);
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private async void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (LibraryCombo.SelectedItem == null)
                {
                    MessageBox.Show("Select a library first.", "Missing Library");
                    return;
                }

                var keep = int.Parse(KeepVersionsBox.Text);
                var page = int.Parse(PageSizeBox.Text);
                var whatIf = WhatIfCheck.IsChecked == true;

                SetBusy(true, "Running...");
                Log($"Running trim on '{LibraryCombo.SelectedItem}' KeepVersions={keep} PageSize={page} WhatIf={whatIf}");

                var args =
                    $"-Action RunTrim " +
                    $"-SiteUrl \"{SiteUrlBox.Text}\" " +
                    $"-ClientId \"{ClientIdBox.Text}\" " +
                    $"-TenantId \"{TenantIdBox.Text}\" " +
                    $"-LibraryName \"{LibraryCombo.SelectedItem}\" " +
                    $"-KeepVersions {keep} " +
                    $"-PageSize {page} " +
                    $"-ReportPath \"{ReportPathBox.Text}\" " +
                    (whatIf ? "-WhatIf" : "");

                var json = await RunPwshAsync(args);
                using var doc = JsonDocument.Parse(json);

                if (!doc.RootElement.GetProperty("ok").GetBoolean())
                {
                    var err = doc.RootElement.GetProperty("error").GetString();
                    Log($"ERROR: {err}");
                    MessageBox.Show(err, "Trim Failed");
                    return;
                }

                var summary = doc.RootElement.GetProperty("summary");
                Log("Completed.");
                Log($"ProcessedFiles: {summary.GetProperty("ProcessedFiles")}");
                Log($"SkippedFiles:   {summary.GetProperty("SkippedFiles")}");
                Log($"FailedFiles:    {summary.GetProperty("FailedFiles")}");
                Log($"DeletedTotal:   {summary.GetProperty("DeletedVersionsTotal")}");
                Log($"ReportPath:     {summary.GetProperty("ReportPath")}");

                SetBusy(false, "Completed");
            }
            catch (Exception ex)
            {
                SetBusy(false, "Error");
                Log("ERROR: " + ex.Message);
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void BrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*",
                FileName = ReportPathBox.Text
            };
            if (dlg.ShowDialog() == true)
                ReportPathBox.Text = dlg.FileName;
        }
    }
}
