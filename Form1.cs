using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using ImportData.Core;
using ImportData.Services;

namespace ImportData
{
    /// <summary>
    /// Main User Interface for the Auto Import Data application.
    /// Manages real-time file monitoring, database synchronization, and system health status.
    /// </summary>
    public partial class Form1 : Form
    {
        private const int MaxLogLines = 1000; 
        
        private readonly AppConfig _config;
        private readonly DatabaseService _dbService;
        private readonly ExcelService _excelService;
        private FileSystemWatcher _watcher;
        private bool _isProcessing;
        private System.Windows.Forms.Timer _healthTimer;
        private string _lastHealthState = "";
        private bool _isSystemHealthy = false;
        private NotifyIcon _trayIcon;

        public Form1()
        {
            InitializeComponent(); 
            
            // UI Initialization
            lstLogs.DrawMode = DrawMode.OwnerDrawFixed;
            lstLogs.DrawItem += LstLogs_DrawItem; 
            
            // Services Initialization
            _config = new AppConfig();
            _config.Load();

            _dbService = new DatabaseService(_config, Log); 
            _excelService = new ExcelService(Log);          

            // Tray Icon Initialization
            _trayIcon = new NotifyIcon()
            {
                Icon = this.Icon,
                Text = "Auto Import Service (Running)",
                Visible = true
            };

            _trayIcon.DoubleClick += (s, e) => {
                this.Show(); 
                this.WindowState = FormWindowState.Normal; 
            };

            this.ShowInTaskbar = true; 
            this.WindowState = FormWindowState.Normal; 

            this.Shown += Form1_Shown; 
            this.FormClosing += Form1_FormClosing;  
        }

        private async void Form1_Shown(object sender, EventArgs e)
        {
            Log($"Monitoring: {_config.BaseFolder}");
            
            _healthTimer = new System.Windows.Forms.Timer { Interval = 10000 };
            _healthTimer.Tick += async (s, ev) => await PerformHealthCheckAsync(); 
            
            RegisterAutoStart();
            
            await PerformHealthCheckAsync(); 
            _healthTimer.Start();
        }

        /// <summary>
        /// Periodically checks system health including directory existence and database connectivity.
        /// </summary>
        private async Task PerformHealthCheckAsync()
        {
            string previousFolder = _config.BaseFolder;
            _config.Load(); 

            string currentState = "HEALTHY";
            bool isDirectoryOk = Directory.Exists(_config.BaseFolder);
            bool isDatabaseOk = isDirectoryOk && await _dbService.TestConnectionAsync();

            if (!isDirectoryOk) currentState = "PATH_ERROR";
            else if (!isDatabaseOk) currentState = "DB_ERROR";

            _isSystemHealthy = isDirectoryOk && isDatabaseOk;

            // Handle Folder Change
            if (_isSystemHealthy && previousFolder != _config.BaseFolder)
            {
                Log($"[INFO] Target folder changed to: {_config.BaseFolder}"); 
                RestartWatcher(); 
                await SynchronizeAsync(); 
            }

            // Handle Health State Changes
            if (currentState != _lastHealthState)
            {
                if (_isSystemHealthy) 
                {
                    if (_lastHealthState != "") Log("[INFO] System recovery: Connection restored.");
                    UpdateStatus("System Ready", Color.Green);
                    RestartWatcher(); 
                    await SynchronizeAsync(); 
                }
                else 
                {
                    StopWatcher(); 
                    if (!isDirectoryOk) 
                    {
                        Log($"[ERROR] Directory not found: {_config.BaseFolder}"); 
                        UpdateStatus("Path Error", Color.Red); 
                    }
                    else if (!isDatabaseOk) 
                    {
                        Log("[ERROR] Database connection failed.");
                        UpdateStatus("SQL Connection Error", Color.Red);
                    }
                }
                _lastHealthState = currentState;
            }
        }

        private void RestartWatcher()
        {
            StopWatcher();
            if (!Directory.Exists(_config.BaseFolder)) 
            {
                Directory.CreateDirectory(_config.BaseFolder); 
            }

            _watcher = new FileSystemWatcher(_config.BaseFolder) 
            {
                IncludeSubdirectories = true,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                Filter = "*.*",
                EnableRaisingEvents = true 
            };

            _watcher.Created += OnFileEvent;
            _watcher.Changed += OnFileEvent;
            _watcher.Renamed += OnFileEvent;
        }

        private void StopWatcher()
        {
            if (_watcher != null) 
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null; 
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();     
                _trayIcon.ShowBalloonTip(2000, "Import Data", "Service is still running in background.", ToolTipIcon.Info); 
            }
        }

        private async void BtnChangeFolder_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select Folder for Measurement Data";
                dialog.SelectedPath = _config.BaseFolder;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _config.BaseFolder = dialog.SelectedPath;
                    _config.Save(); 
                    Log($"[INFO] Folder updated: {_config.BaseFolder}");
                    await PerformHealthCheckAsync();
                }
            }
        }

        private void UpdateStatus(string message, Color color)
        {
            if (this.InvokeRequired) 
            {
                this.Invoke(new Action(() => UpdateStatus(message, color))); 
                return;
            }
            lblStatus.Text = message; 
            lblStatus.ForeColor = color; 
        }

        private void LstLogs_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            string text = lstLogs.Items[e.Index].ToString(); 
            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            // Background
            if (isSelected)
            {
                using (var backBrush = new SolidBrush(Color.FromArgb(0, 120, 215))) 
                    e.Graphics.FillRectangle(backBrush, e.Bounds);
            }
            else
            {
                e.Graphics.FillRectangle(Brushes.Black, e.Bounds);
            }

            // Foreground
            int timeEndIndex = text.IndexOf(']');
            string timePart = timeEndIndex > 0 ? text.Substring(0, timeEndIndex + 1) : "";
            string msgPart = timeEndIndex > 0 ? text.Substring(timeEndIndex + 1) : text;
            
            Brush textBrush = isSelected ? Brushes.White : Brushes.Lime;
            e.Graphics.DrawString(timePart, e.Font, textBrush, new PointF(e.Bounds.X + 5, e.Bounds.Y + 2));
            e.Graphics.DrawString(msgPart, e.Font, textBrush, new PointF(e.Bounds.X + 110, e.Bounds.Y + 2));
        }

        private void Log(string message)
        {
            if (lstLogs.InvokeRequired) 
            {
                lstLogs.Invoke(new Action(() => Log(message))); 
                return;
            }
            
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            lstLogs.Items.Add($"[{timestamp}] {message}"); 
            lstLogs.SelectedIndex = lstLogs.Items.Count - 1; 
            
            if (lstLogs.Items.Count > MaxLogLines) 
                lstLogs.Items.RemoveAt(0); 
        }

        private async void OnFileEvent(object sender, FileSystemEventArgs e)
        {
            if (!_isSystemHealthy) return; 

            string ext = Path.GetExtension(e.FullPath).ToLower();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm") return;

            // Only process files in today's folder
            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            if (!e.FullPath.Contains(todayFolder)) return; 

            if (_isProcessing) return; 
            
            try 
            {
                _isProcessing = true; 
                await ProcessSingleFileAsync(e.FullPath);
            }
            finally 
            {
                _isProcessing = false; 
            }
        }

        private async Task SynchronizeAsync()
        {
            if (!_isSystemHealthy) return; 

            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            string sourcePath = Path.Combine(_config.BaseFolder, todayFolder);
            
            if (!Directory.Exists(sourcePath)) 
            {
                UpdateStatus($"Waiting for: {todayFolder}", Color.White);
                return;
            }

            UpdateStatus("Synchronizing...", Color.Yellow); 

            try
            {
                string[] files = Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories)
                    .Where(f => f.EndsWith(".xlsx") || f.EndsWith(".xls") || f.EndsWith(".xlsm"))
                    .ToArray();

                foreach (string file in files) 
                {
                    await ProcessSingleFileAsync(file);
                }
                
                UpdateStatus("System Ready", Color.Green); 
            }
            catch (Exception ex)
            {
                Log($"[ERROR] Synchronization failed: {ex.Message}"); 
                UpdateStatus("Sync Error", Color.Red); 
            }
        }

        private async Task ProcessSingleFileAsync(string filePath)
        {
            string fileName = Path.GetFileName(filePath);

            try
            {
                if (await _dbService.IsFileImportedAsync(filePath)) return; 

                if (!await IsFileReadyAsync(filePath))
                {
                    Log($"[SKIP] File locked or busy: {fileName}");
                    return;
                }

                Log($"[INFO] Processing: {fileName}");

                var data = _excelService.ReadExcelFile(filePath);
                if (data == null || data.Rows.Count == 0) return; 

                await _dbService.ExecuteImportBatchAsync(data, fileName, filePath);
            }
            catch (Exception ex)
            {
                Log($"[ERROR] Failed to process {fileName}: {ex.Message}"); 
            }
        }

        private async Task<bool> IsFileReadyAsync(string filePath)
        {
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                        return true; 
                }
                catch (IOException) 
                {
                    await Task.Delay(1000); 
                }
            }
            return false;
        }

        private void RegisterAutoStart()
        {
            try 
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key != null) key.SetValue("AutoImportData", Application.ExecutablePath);
                }
            }
            catch (Exception ex) 
            {
                Log($"[WARN] Failed to set auto-start: {ex.Message}");
            }
        }
    }
}
