using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Threading;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ANVESHA_TCRX_HEALTH_STATUS_GUI_V2.ViewModels;

namespace ANVESHA_TCRX_HEALTH_STATUS_GUI_V2
{
    public partial class MainWindow : Window
    {
        public MainViewModel _vm;

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — OFFLINE COMPARE PRIVATE FIELDS
        // ══════════════════════════════════════════════════════════════════
        private string _offlineExpectedFile = string.Empty;
        private string _offlineSingleFile   = string.Empty;
        private string _offlineMainFile     = string.Empty;
        private string _offlineRedundantFile = string.Empty;
        private bool   _isOfflineMultipleMode = false;
        private string _logFolderPath = 
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        // ── Frame layout constants (match your 26-byte frame) ────────────
        private const int  FRAME_SIZE        = 26;
        private const int  CMD_OFFSET        = 16;   // byte index where 8-byte payload starts
        private const int  CMD_SIZE          = 8;
        private const int  SOURCE_BYTE_INDEX = 4;    // byte that carries MAIN/RED bit
        private const byte SOURCE_BIT_MASK   = 0x01; // 0=MAIN, 1=REDUNDANT
        private const byte HDR1              = 0x0D;
        private const byte HDR2              = 0x0A;
        private const byte FOOTER            = 0xAB;

        public MainWindow()
        {
            InitializeComponent();
            _vm = new MainViewModel(Dispatcher);
            DataContext = _vm;
        }

        private void Window_Closing(object sender,
            System.ComponentModel.CancelEventArgs e)
        {
            if (_vm != null)
            {
                _vm.Port1.Disconnect();
                _vm.Port2.Disconnect();
                _vm.Dispose();
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  COMMON BUTTONS (All Tabs)
        // ══════════════════════════════════════════════════════════════════
        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            _vm.Port1.RefreshPorts();
            _vm.Port2.RefreshPorts();
            _vm.AppStatus = string.Format(
                "COM ports refreshed: {0} available",
                _vm.Port1.AvailablePorts.Count);
        }

        private void BtnConnect_Click(object sender, RoutedEventArgs e)
        {
            // Validate port selection
            if (string.IsNullOrEmpty(_vm.Port1.SelectedPort) ||
                string.IsNullOrEmpty(_vm.Port2.SelectedPort))
            {
                MessageBox.Show(
                    "Select a COM port for BOTH PORT 1 and PORT 2.",
                    "Port Selection Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // Validate different ports
            if (_vm.Port1.SelectedPort == _vm.Port2.SelectedPort)
            {
                MessageBox.Show(
                    "PORT 1 and PORT 2 must use different COM ports.\n" +
                    "Both are currently set to: " + _vm.Port1.SelectedPort,
                    "Port Conflict",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // Disconnect first (cleanup any existing connections)
            _vm.Port1.Disconnect();
            _vm.Port2.Disconnect();

            // Attempt connection
            bool ok1 = _vm.Port1.Connect();
            bool ok2 = _vm.Port2.Connect();

            if (ok1 && ok2)
            {
                // Both connected successfully - NO IsAcquiring change
                _vm.AppStatus = string.Format(
                    "BOTH PORTS CONNECTED  |  " +
                    "PORT1=[{0}]@{1}  PORT2=[{2}]@{3}",
                    _vm.Port1.SelectedPort, _vm.Port1.SelectedBaudRate,
                    _vm.Port2.SelectedPort, _vm.Port2.SelectedBaudRate);
            }
            else
            {
                // Connection failed - cleanup
                if (ok1) _vm.Port1.Disconnect();
                if (ok2) _vm.Port2.Disconnect();

                string failMsg = "Connection FAILED:\n\n";
                if (!ok1) failMsg += "  PORT 1 [" + _vm.Port1.SelectedPort + "]\n";
                if (!ok2) failMsg += "  PORT 2 [" + _vm.Port2.SelectedPort + "]\n";
                failMsg +=
                    "\nCheck:\n" +
                    "  RS422 adapter connected\n" +
                    "  Correct COM port selected\n" +
                    "  Port not in use by another application";

                MessageBox.Show(failMsg, "Connection Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);

                _vm.AppStatus = "Connection failed.";
            }

            _vm.UpdateAcquisitionState();
        }

        private void BtnDisconnect_Click(object sender, RoutedEventArgs e)
        {
            long p1 = _vm.Port1.TotalFrames;
            long p2 = _vm.Port2.TotalFrames;

            // Disconnect both ports - NO IsAcquiring change
            _vm.Port1.Disconnect();
            _vm.Port2.Disconnect();
            _vm.UpdateAcquisitionState();

            _vm.AppStatus = string.Format(
                "BOTH DISCONNECTED  |  PORT1={0} frames  PORT2={1} frames",
                p1, p2);
        }

        private void BtnClearAll_Click(object sender, RoutedEventArgs e)
        {
            _vm.Port1.ClearLogs();
            _vm.Port2.ClearLogs();
            _vm.AppStatus = "All display logs cleared.";
        }

        // ══════════════════════════════════════════════════════════════════
        //  START ALL / STOP ALL BUTTONS
        // ══════════════════════════════════════════════════════════════════
        private void BtnStartAll_Click(object sender, RoutedEventArgs e)
        {
            // Validate port selection
            if (string.IsNullOrEmpty(_vm.Port1.SelectedPort) ||
                string.IsNullOrEmpty(_vm.Port2.SelectedPort))
            {
                MessageBox.Show(
                    "Select a COM port for BOTH PORT 1 and PORT 2.",
                    "Port Selection Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // Validate different ports
            if (_vm.Port1.SelectedPort == _vm.Port2.SelectedPort)
            {
                MessageBox.Show(
                    "PORT 1 and PORT 2 must use different COM ports.\n" +
                    "Both are currently set to: " + _vm.Port1.SelectedPort,
                    "Port Conflict",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // Disconnect first (cleanup)
            _vm.Port1.Disconnect();
            _vm.Port2.Disconnect();

            // Attempt connection
            bool ok1 = _vm.Port1.Connect();
            bool ok2 = _vm.Port2.Connect();

            if (ok1 && ok2)
            {
                // START ACQUISITION
                _vm.IsAcquiring = true;
                _vm.UpdateAcquisitionState();

                _vm.AppStatus = string.Format(
                    "ACQUISITION STARTED  |  " +
                    "PORT1=[{0}]@{1}  PORT2=[{2}]@{3}",
                    _vm.Port1.SelectedPort, _vm.Port1.SelectedBaudRate,
                    _vm.Port2.SelectedPort, _vm.Port2.SelectedBaudRate);
            }
            else
            {
                // Connection failed - cleanup
                if (ok1) _vm.Port1.Disconnect();
                if (ok2) _vm.Port2.Disconnect();

                string failMsg = "Connection FAILED:\n\n";
                if (!ok1) failMsg += "  PORT 1 [" + _vm.Port1.SelectedPort + "]\n";
                if (!ok2) failMsg += "  PORT 2 [" + _vm.Port2.SelectedPort + "]\n";
                failMsg +=
                    "\nCheck:\n" +
                    "  RS422 adapter connected\n" +
                    "  Correct COM port selected\n" +
                    "  Port not in use by another application";

                MessageBox.Show(failMsg, "Connection Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);

                _vm.IsAcquiring = false;
                _vm.UpdateAcquisitionState();
                _vm.AppStatus = "Start acquisition failed.";
            }
        }

        private void BtnStopAll_Click(object sender, RoutedEventArgs e)
        {
            long p1 = _vm.Port1.TotalFrames;
            long p2 = _vm.Port2.TotalFrames;

            // STOP ACQUISITION
            _vm.IsAcquiring = false;

            // Disconnect both ports
            _vm.Port1.Disconnect();
            _vm.Port2.Disconnect();

            _vm.UpdateAcquisitionState();

            _vm.AppStatus = string.Format(
                "ACQUISITION STOPPED  |  PORT1={0} frames  PORT2={1} frames",
                p1, p2);
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 1 — 26 Bytes Raw
        // ══════════════════════════════════════════════════════════════════
        private void TbPort1Raw_TextChanged(object sender, TextChangedEventArgs e)
        {
            ScrollToEnd(sender as TextBox);
        }

        private void TbPort2Raw_TextChanged(object sender, TextChangedEventArgs e)
        {
            ScrollToEnd(sender as TextBox);
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 2 — 8 Bytes Payload
        // ══════════════════════════════════════════════════════════════════
        private void TbPort1Payload_TextChanged(object sender, TextChangedEventArgs e)
        {
            ScrollToEnd(sender as TextBox);
        }

        private void TbPort2Payload_TextChanged(object sender, TextChangedEventArgs e)
        {
            ScrollToEnd(sender as TextBox);
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 3 — RX Live Compare
        // ══════════════════════════════════════════════════════════════════
        private void TbPort1RxLive_TextChanged(object sender, TextChangedEventArgs e)
        {
            ScrollToEnd(sender as TextBox);
        }

        private void TbPort2RxLive_TextChanged(object sender, TextChangedEventArgs e)
        {
            ScrollToEnd(sender as TextBox);
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 4 — RX Live Stats Compact
        // ══════════════════════════════════════════════════════════════════
        private void TbPort1Compact_TextChanged(object sender, TextChangedEventArgs e)
        {
            ScrollToEnd(sender as TextBox);
        }

        private void TbPort2Compact_TextChanged(object sender, TextChangedEventArgs e)
        {
            ScrollToEnd(sender as TextBox);
        }
        // ══════════════════════════════════════════════════════════════════
        //  HELPER — Auto-scroll TextBox
        // ══════════════════════════════════════════════════════════════════
        private static void ScrollToEnd(TextBox tb)
        {
            if (tb != null) tb.ScrollToEnd();
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — OFFLINE COMPARE HELPER CLASSES
        // ══════════════════════════════════════════════════════════════════
        #region Offline Compare Helper Classes

        private class ExpectedCommand
        {
            public byte[] CommandBytes;
            public string HexString;

            public ExpectedCommand(byte[] bytes)
            {
                CommandBytes = bytes;
                HexString    = BitConverter.ToString(bytes).Replace("-", " ");
            }
        }

        private class LogCommand
        {
            public DateTime Timestamp;
            public string   Port;
            public int      FrameNumber;
            public byte[]   Command8;
            public byte[]   FullFrame26;
            public string   Source;     // "MAIN" or "REDUNDANT"
            public string   HexString;
            // kept for single-file compatibility
            public string   OriginalLine;
        }

        // ── Per-command result (used by CompareLoopsWithExpected) ─────────
        private class ExpectedCommandResult
        {
            public ExpectedCommand Expected;
            public bool   FoundInMain;
            public bool   FoundInRedundant;
            public int    MainLoopCount;
            public int    RedundantLoopCount;
            public string Status;        // "FOUND IN MAIN" | "FOUND IN REDUNDANT" | "FAILED"
            public string MatchedSource; // "MAIN" | "REDUNDANT" | "NONE"
        }

        // ── Overall result ────────────────────────────────────────────────
        private class OfflineComparisonResult
        {
            public int TotalExpected;
            public int TotalMainLoops;
            public int TotalRedundantLoops;
            public int TotalFound;
            public int FoundInMain;
            public int FoundInRedundant;
            public int TotalMissing;
            public double SuccessRate;
            public List<ExpectedCommandResult> CommandResults 
                = new List<ExpectedCommandResult>();
        }

        #endregion

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — OFFLINE FILE MODE SWITCHING
        // ══════════════════════════════════════════════════════════════════
        private void BtnOfflineModeSingle_Click(object sender, RoutedEventArgs e)
        {
            _isOfflineMultipleMode = false;

            borderSingleFileMode.Visibility  = Visibility.Visible;
            stackMultipleFileMode.Visibility = Visibility.Collapsed;

            btnOfflineModeSingle.Background = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF00D4FF"));
            btnOfflineModeSingle.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF0A0A1A"));

            btnOfflineModeMultiple.Background = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF0A1628"));
            btnOfflineModeMultiple.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF95A5A6"));

            txtOfflineModeDesc.Text =
                "Single combined log file (Main + Redundant data)";

            _offlineSingleFile          = "";
            txtOfflineSingleFile.Text   =
                "No file selected - Click Browse to choose file";

            UpdateOfflineFilesStatus();
            _vm.AppStatus = "Offline Mode: Single File";
        }

        private void BtnOfflineModeMultiple_Click(object sender, RoutedEventArgs e)
        {
            _isOfflineMultipleMode = true;

            borderSingleFileMode.Visibility  = Visibility.Collapsed;
            stackMultipleFileMode.Visibility = Visibility.Visible;

            btnOfflineModeMultiple.Background = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF00D4FF"));
            btnOfflineModeMultiple.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF0A0A1A"));

            btnOfflineModeSingle.Background = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF0A1628"));
            btnOfflineModeSingle.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF95A5A6"));

            txtOfflineModeDesc.Text = "Separate MAIN and REDUNDANT log files";

            _offlineMainFile              = "";
            _offlineRedundantFile         = "";
            txtOfflineMainFile.Text       = "No file selected - Click Browse";
            txtOfflineRedundantFile.Text  = "No file selected - Click Browse";

            UpdateOfflineFilesStatus();
            _vm.AppStatus = "Offline Mode: Multiple Files";
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — FILE BROWSE BUTTONS
        // ══════════════════════════════════════════════════════════════════
        private void BtnBrowseExpectedFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title           = "Select Expected Commands File",
                Filter          = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                InitialDirectory = _logFolderPath
            };

            if (dlg.ShowDialog() == true)
            {
                _offlineExpectedFile        = dlg.FileName;
                txtOfflineExpectedFile.Text = _offlineExpectedFile;

                try
                {
                    var fi       = new FileInfo(_offlineExpectedFile);
                    int cmdCount = CountValidCommandsInFile(_offlineExpectedFile);
                    txtExpectedFileName.Text  = fi.Name;
                    txtExpectedFileSize.Text  = string.Format("{0:N0} bytes", fi.Length);
                    txtExpectedFileCount.Text = string.Format("{0} commands", cmdCount);
                    borderExpectedFileInfo.Visibility = Visibility.Visible;
                }
                catch { borderExpectedFileInfo.Visibility = Visibility.Collapsed; }

                UpdateOfflineFilesStatus();
                _vm.AppStatus = "Expected: " + Path.GetFileName(_offlineExpectedFile);
            }
        }

        private void BtnBrowseSingleFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title            = "Select Combined Log File",
                Filter           = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                InitialDirectory = _logFolderPath
            };

            if (dlg.ShowDialog() == true)
            {
                _offlineSingleFile        = dlg.FileName;
                txtOfflineSingleFile.Text = _offlineSingleFile;

                try
                {
                    var      fi    = new FileInfo(_offlineSingleFile);
                    string[] lines = File.ReadAllLines(_offlineSingleFile);
                    txtSingleFileName.Text  = fi.Name;
                    txtSingleFileSize.Text  = string.Format("{0:N0} bytes", fi.Length);
                    txtSingleFileLines.Text = string.Format("{0} lines", lines.Length);
                    borderSingleFileInfo.Visibility = Visibility.Visible;
                }
                catch { borderSingleFileInfo.Visibility = Visibility.Collapsed; }

                UpdateOfflineFilesStatus();
                _vm.AppStatus = "Log file: " + Path.GetFileName(_offlineSingleFile);
            }
        }

        private void BtnBrowseMainFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title            = "Select Main Data Log File",
                Filter           = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                InitialDirectory = Path.Combine(_logFolderPath, "Main")
            };

            if (dlg.ShowDialog() == true)
            {
                _offlineMainFile        = dlg.FileName;
                txtOfflineMainFile.Text = _offlineMainFile;

                try
                {
                    var      fi    = new FileInfo(_offlineMainFile);
                    string[] lines = File.ReadAllLines(_offlineMainFile);
                    txtMainFileName.Text  = fi.Name;
                    txtMainFileSize.Text  = string.Format("{0:N0} bytes", fi.Length);
                    txtMainFileLines.Text = string.Format("{0} lines", lines.Length);
                    borderMainFileInfo.Visibility = Visibility.Visible;
                }
                catch { borderMainFileInfo.Visibility = Visibility.Collapsed; }

                UpdateOfflineFilesStatus();
                _vm.AppStatus = "Main: " + Path.GetFileName(_offlineMainFile);
            }
        }

        private void BtnBrowseRedundantFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title            = "Select Redundant Data Log File",
                Filter           = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                InitialDirectory = Path.Combine(_logFolderPath, "Redundant")
            };

            if (dlg.ShowDialog() == true)
            {
                _offlineRedundantFile        = dlg.FileName;
                txtOfflineRedundantFile.Text = _offlineRedundantFile;

                try
                {
                    var      fi    = new FileInfo(_offlineRedundantFile);
                    string[] lines = File.ReadAllLines(_offlineRedundantFile);
                    txtRedundantFileName.Text  = fi.Name;
                    txtRedundantFileSize.Text  = string.Format("{0:N0} bytes", fi.Length);
                    txtRedundantFileLines.Text = string.Format("{0} lines", lines.Length);
                    borderRedundantFileInfo.Visibility = Visibility.Visible;
                }
                catch { borderRedundantFileInfo.Visibility = Visibility.Collapsed; }

                UpdateOfflineFilesStatus();
                _vm.AppStatus = "Redundant: " + Path.GetFileName(_offlineRedundantFile);
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — CLEAR FILE BUTTONS
        // ══════════════════════════════════════════════════════════════════
        private void BtnClearExpectedFile_Click(object sender, RoutedEventArgs e)
        {
            _offlineExpectedFile        = "";
            txtOfflineExpectedFile.Text =
                "No file selected - Click Browse to choose file";
            borderExpectedFileInfo.Visibility = Visibility.Collapsed;
            UpdateOfflineFilesStatus();
        }

        private void BtnClearSingleFile_Click(object sender, RoutedEventArgs e)
        {
            _offlineSingleFile        = "";
            txtOfflineSingleFile.Text =
                "No file selected - Click Browse to choose file";
            borderSingleFileInfo.Visibility = Visibility.Collapsed;
            UpdateOfflineFilesStatus();
        }

        private void BtnClearMainFile_Click(object sender, RoutedEventArgs e)
        {
            _offlineMainFile        = "";
            txtOfflineMainFile.Text = "No file selected - Click Browse";
            borderMainFileInfo.Visibility = Visibility.Collapsed;
            UpdateOfflineFilesStatus();
        }

        private void BtnClearRedundantFile_Click(object sender, RoutedEventArgs e)
        {
            _offlineRedundantFile        = "";
            txtOfflineRedundantFile.Text = "No file selected - Click Browse";
            borderRedundantFileInfo.Visibility = Visibility.Collapsed;
            UpdateOfflineFilesStatus();
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — UPDATE OFFLINE FILES STATUS
        // ══════════════════════════════════════════════════════════════════
        private void UpdateOfflineFilesStatus()
        {
            int fileCount     = 0;
            int requiredCount = _isOfflineMultipleMode ? 3 : 2;

            if (!string.IsNullOrEmpty(_offlineExpectedFile)) fileCount++;

            if (_isOfflineMultipleMode)
            {
                if (!string.IsNullOrEmpty(_offlineMainFile))      fileCount++;
                if (!string.IsNullOrEmpty(_offlineRedundantFile)) fileCount++;
            }
            else
            {
                if (!string.IsNullOrEmpty(_offlineSingleFile)) fileCount++;
            }

            bool allSelected = (fileCount == requiredCount);

            txtOfflineFilesStatus.Text = string.Format(
                "{0} / {1} selected", fileCount, requiredCount);

            if (allSelected)
            {
                txtOfflineReadyStatus.Text     = "All files selected - Ready!";
                txtOfflineReadyStatus.Foreground =
                    new SolidColorBrush(Color.FromRgb(0, 255, 136));
                btnCompareOfflineFiles.IsEnabled = true;
            }
            else
            {
                txtOfflineReadyStatus.Text = string.Format(
                    "Need {0} more file(s)", requiredCount - fileCount);
                txtOfflineReadyStatus.Foreground =
                    new SolidColorBrush(Color.FromRgb(255, 68, 68));
                btnCompareOfflineFiles.IsEnabled = false;
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — COUNT VALID COMMANDS IN FILE
        // ══════════════════════════════════════════════════════════════════
        private int CountValidCommandsInFile(string filePath)
        {
            int count = 0;
            try
            {
                foreach (string line in File.ReadAllLines(filePath))
                {
                    string t = line.Trim();
                    if (string.IsNullOrEmpty(t)) continue;
                    if (t.StartsWith("#") || t.StartsWith("//")) continue;
                    byte[] b = ParseExpectedCommandLine(t);
                    if (b != null && b.Length == 8) count++;
                }
            }
            catch { }
            return count;
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — CLEAR OFFLINE RESULTS
        // ══════════════════════════════════════════════════════════════════
        private void BtnClearOfflineResults_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                txtOfflineCompareResults.Clear();
                txtOfflineMatchStatus.Text = "Ready - Select files to compare";

                if (ellipseOfflineStatus != null)
                    ellipseOfflineStatus.Fill =
                        new SolidColorBrush(Color.FromRgb(255, 68, 68));

                if (panelOfflinePlaceholder != null)
                    panelOfflinePlaceholder.Visibility = Visibility.Visible;

                btnCompareOfflineFiles.Content = "START COMPARE";
                UpdateOfflineFilesStatus();
                _vm.AppStatus = "Offline results cleared";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("Error clearing:\n{0}", ex.Message),
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — EXPORT OFFLINE REPORT
        // ══════════════════════════════════════════════════════════════════
        private void BtnExportOfflineReport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtOfflineCompareResults.Text))
            {
                MessageBox.Show("No results to export!",
                    "Nothing to Export", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            try
            {
                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    Title            = "Export Offline Comparison Report",
                    Filter           = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                    FileName         = string.Format("Offline_Report_{0:yyyyMMdd_HHmmss}.txt", DateTime.Now),
                    InitialDirectory = _logFolderPath
                };

                if (dlg.ShowDialog() == true)
                {
                    File.WriteAllText(dlg.FileName,
                        txtOfflineCompareResults.Text, Encoding.UTF8);

                    MessageBox.Show(
                        string.Format("Report saved:\n{0}", dlg.FileName),
                        "Export Complete", MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    _vm.AppStatus = "Report exported: " +
                        Path.GetFileName(dlg.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("Export failed:\n{0}", ex.Message),
                    "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — OPEN LOG FOLDERS
        // ══════════════════════════════════════════════════════════════════
        private void BtnOpenLogFolders_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Directory.Exists(_logFolderPath))
                {
                    Process.Start("explorer.exe",
                        string.Format("\"{0}\"", _logFolderPath));
                    _vm.AppStatus = "Opened DataLogs folder";
                }
                else
                {
                    MessageBox.Show(
                        string.Format("Folder not found:\n{0}", _logFolderPath),
                        "Not Found", MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("Error:\n{0}", ex.Message),
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateStatus(string message)
        {
            // kept for compatibility – forward to AppStatus
            _vm.AppStatus = message;
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — BUTTON CLICK — OFFLINE COMPARE  (worker thread)
        // ══════════════════════════════════════════════════════════════════
        private void BtnCompareOfflineFiles_Click(object sender, RoutedEventArgs e)
        {
            // ── validation ───────────────────────────────────────────────
            if (string.IsNullOrEmpty(_offlineExpectedFile) ||
                !File.Exists(_offlineExpectedFile))
            {
                MessageBox.Show("Please select Expected Commands file first!",
                    "Missing File", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!_isOfflineMultipleMode)
            {
                if (string.IsNullOrEmpty(_offlineSingleFile) ||
                    !File.Exists(_offlineSingleFile))
                {
                    MessageBox.Show("Please select Log File first!",
                        "Missing File", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            else
            {
                if (string.IsNullOrEmpty(_offlineMainFile) ||
                    !File.Exists(_offlineMainFile))
                {
                    MessageBox.Show("Please select Main Data file first!",
                        "Missing File", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (string.IsNullOrEmpty(_offlineRedundantFile) ||
                    !File.Exists(_offlineRedundantFile))
                {
                    MessageBox.Show("Please select Redundant Data file first!",
                        "Missing File", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            // ── UI reset ─────────────────────────────────────────────────
            txtOfflineCompareResults.Clear();
            if (panelOfflinePlaceholder != null)
                panelOfflinePlaceholder.Visibility = Visibility.Collapsed;

            txtOfflineMatchStatus.Text = "Initializing...";
            if (ellipseOfflineStatus != null)
                ellipseOfflineStatus.Fill =
                    new SolidColorBrush(Color.FromRgb(255, 140, 0));

            btnCompareOfflineFiles.IsEnabled = false;
            btnCompareOfflineFiles.Content   = "Processing...";

            // ── capture fields for thread ────────────────────────────────
            DateTime startTime           = DateTime.Now;
            string capturedExpectedFile  = _offlineExpectedFile;
            string capturedSingleFile    = _offlineSingleFile;
            string capturedMainFile      = _offlineMainFile;
            string capturedRedFile       = _offlineRedundantFile;
            bool   isMultiMode           = _isOfflineMultipleMode;

            // ── worker thread ────────────────────────────────────────────
            var workerThread = new System.Threading.Thread(delegate()
            {
                try
                {
                    // 1. load expected commands
                    Dispatcher.Invoke(new Action(delegate
                    {
                        txtOfflineMatchStatus.Text = "Loading expected commands...";
                        UpdateStatus("Loading expected commands...");
                    }));

                    List<ExpectedCommand> expectedCommands =
                        LoadOfflineExpectedCommands(capturedExpectedFile);

                    if (expectedCommands.Count == 0)
                    {
                        Dispatcher.Invoke(new Action(delegate
                        {
                            txtOfflineMatchStatus.Text =
                                "No valid expected commands found";
                            if (ellipseOfflineStatus != null)
                                ellipseOfflineStatus.Fill =
                                    new SolidColorBrush(Color.FromRgb(255, 68, 68));
                            btnCompareOfflineFiles.IsEnabled = true;
                            btnCompareOfflineFiles.Content   = "START COMPARE";
                            MessageBox.Show("No valid expected commands found!",
                                "Error", MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                        }));
                        return;
                    }

                    byte[] firstCmd = expectedCommands[0].CommandBytes;
                    byte[] lastCmd  =
                        expectedCommands[expectedCommands.Count - 1].CommandBytes;
                    string firstHex =
                        BitConverter.ToString(firstCmd).Replace("-", " ");
                    string lastHex  =
                        BitConverter.ToString(lastCmd).Replace("-", " ");

                    // 2. extract loops
                    List<List<LogCommand>> mainLoops;
                    List<List<LogCommand>> redundantLoops;

                    if (!isMultiMode)
                    {
                        Dispatcher.Invoke(new Action(delegate
                        {
                            txtOfflineMatchStatus.Text =
                                "Extracting loops from single file...";
                        }));
                        ExtractLoopsFromSingleFile(
                            capturedSingleFile, firstCmd, lastCmd,
                            out mainLoops, out redundantLoops);
                    }
                    else
                    {
                        Dispatcher.Invoke(new Action(delegate
                        {
                            txtOfflineMatchStatus.Text = "Extracting MAIN loops...";
                        }));
                        mainLoops = ExtractLoopsFromFile(
                            capturedMainFile, "MAIN", firstCmd, lastCmd);

                        Dispatcher.Invoke(new Action(delegate
                        {
                            txtOfflineMatchStatus.Text =
                                "Extracting REDUNDANT loops...";
                        }));
                        redundantLoops = ExtractLoopsFromFile(
                            capturedRedFile, "REDUNDANT", firstCmd, lastCmd);
                    }

                    // 3. no loops guard
                    if (mainLoops.Count == 0 && redundantLoops.Count == 0)
                    {
                        Dispatcher.Invoke(new Action(delegate
                        {
                            txtOfflineMatchStatus.Text = "No loops detected!";
                            if (ellipseOfflineStatus != null)
                                ellipseOfflineStatus.Fill =
                                    new SolidColorBrush(Color.FromRgb(255, 68, 68));
                            btnCompareOfflineFiles.IsEnabled = true;
                            btnCompareOfflineFiles.Content   = "START COMPARE";
                            MessageBox.Show(
                                string.Format(
                                    "No loops detected!\n\n"
                                  + "Loop markers used:\n"
                                  + "  START : {0}\n"
                                  + "  END   : {1}\n\n"
                                  + "Possible causes:\n"
                                  + "  Commands not found in log file\n"
                                  + "  Log file format does not match\n"
                                  + "  Source byte filtering excluding frames",
                                    firstHex, lastHex),
                                "No Loops Found",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                        }));
                        return;
                    }

                    Dispatcher.Invoke(new Action(delegate
                    {
                        txtOfflineMatchStatus.Text = string.Format(
                            "Found: MAIN={0} loops, RED={1} loops. Comparing...",
                            mainLoops.Count, redundantLoops.Count);
                    }));

                    // 4. compare
                    OfflineComparisonResult compareResults =
                        CompareLoopsWithExpected(
                            expectedCommands, mainLoops, redundantLoops);

                    // 5. loop-level pass/fail tally
                    int totalLoops    = Math.Max(mainLoops.Count, redundantLoops.Count);
                    int passCount     = 0;
                    int failCount     = 0;
                    int totalFromRed  = 0;

                    for (int loopIdx = 0; loopIdx < totalLoops; loopIdx++)
                    {
                        List<LogCommand> mainLoop =
                            loopIdx < mainLoops.Count
                            ? mainLoops[loopIdx] : null;
                        List<LogCommand> redLoop  =
                            loopIdx < redundantLoops.Count
                            ? redundantLoops[loopIdx] : null;

                        int trulyMissing = 0;

                        foreach (var expected in expectedCommands)
                        {
                            bool inMain = false;
                            if (mainLoop != null)
                            {
                                foreach (var cmd in mainLoop)
                                {
                                    if (CompareByteArrays(
                                        expected.CommandBytes, cmd.Command8))
                                    { inMain = true; break; }
                                }
                            }

                            if (!inMain)
                            {
                                bool inRed = false;
                                if (redLoop != null)
                                {
                                    foreach (var cmd in redLoop)
                                    {
                                        if (CompareByteArrays(
                                            expected.CommandBytes, cmd.Command8))
                                        { inRed = true; break; }
                                    }
                                }

                                if (inRed) totalFromRed++;
                                else       trulyMissing++;
                            }
                        }

                        if (trulyMissing == 0) passCount++;
                        else                   failCount++;
                    }

                    TimeSpan elapsed         = DateTime.Now - startTime;
                    int  finalPassCount      = passCount;
                    int  finalFailCount      = failCount;
                    int  finalTotalLoops     = totalLoops;
                    int  finalExpectedCount  = expectedCommands.Count;
                    int  finalTookFromRed    = totalFromRed;
                    double finalElapsed      = elapsed.TotalSeconds;

                    // 6. update UI on dispatcher
                    Dispatcher.Invoke(new Action(delegate
                    {
                        GenerateOfflineComparisonReport(
                            compareResults,
                            expectedCommands,
                            mainLoops,
                            redundantLoops);

                        bool overallPerfect =
                            (finalPassCount == finalTotalLoops &&
                             finalTotalLoops > 0);

                        if (ellipseOfflineStatus != null)
                        {
                            Color c;
                            if (overallPerfect)
                                c = Color.FromRgb(0, 255, 136);
                            else if (finalPassCount > 0)
                                c = Color.FromRgb(255, 140, 0);
                            else
                                c = Color.FromRgb(255, 68, 68);
                            ellipseOfflineStatus.Fill = new SolidColorBrush(c);
                        }

                        txtOfflineMatchStatus.Text = string.Format(
                            "Complete: {0}/{1} loops PASS | {2:F1}s",
                            finalPassCount, finalTotalLoops, finalElapsed);

                        UpdateStatus(string.Format(
                            "Done: {0}/{1} PASS, {2}/{1} FAIL",
                            finalPassCount, finalTotalLoops, finalFailCount));

                        btnCompareOfflineFiles.IsEnabled = true;
                        btnCompareOfflineFiles.Content   = "START COMPARE";

                        ShowFinalSummaryMessageBox(
                            finalPassCount, finalFailCount,
                            finalTotalLoops, finalExpectedCount,
                            finalTookFromRed, finalElapsed);
                    }));
                }
                catch (Exception ex)
                {
                    string errMsg = ex.Message;
                    Dispatcher.Invoke(new Action(delegate
                    {
                        if (ellipseOfflineStatus != null)
                            ellipseOfflineStatus.Fill =
                                new SolidColorBrush(Color.FromRgb(255, 68, 68));
                        txtOfflineMatchStatus.Text       = "Error occurred";
                        UpdateStatus("ERROR: " + errMsg);
                        btnCompareOfflineFiles.IsEnabled = true;
                        btnCompareOfflineFiles.Content   = "START COMPARE";
                        MessageBox.Show("Error:\n\n" + errMsg,
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }));
                }
            });

            workerThread.IsBackground = true;
            workerThread.Start();
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — LOAD EXPECTED COMMANDS  (with duplicate suppression)
        // ══════════════════════════════════════════════════════════════════
        private List<ExpectedCommand> LoadOfflineExpectedCommands(string filePath)
        {
            var commands = new List<ExpectedCommand>();

            try
            {
                foreach (string line in File.ReadAllLines(filePath))
                {
                    string t = line.Trim();
                    if (string.IsNullOrEmpty(t)) continue;
                    if (t.StartsWith("#") || t.StartsWith("//")) continue;

                    byte[] bytes = ParseExpectedCommandLine(t);
                    if (bytes == null || bytes.Length != 8) continue;

                    // skip duplicates
                    bool duplicate = false;
                    foreach (var existing in commands)
                    {
                        if (CompareByteArrays(existing.CommandBytes, bytes))
                        { duplicate = true; break; }
                    }

                    if (!duplicate)
                        commands.Add(new ExpectedCommand(bytes));
                }
            }
            catch (Exception ex)
            {
                throw new Exception(
                    string.Format("Failed to load expected commands:\n{0}",
                        ex.Message));
            }

            return commands;
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — PARSE EXPECTED COMMAND LINE
        // ══════════════════════════════════════════════════════════════════
        private byte[] ParseExpectedCommandLine(string line)
        {
            try
            {
                string hexOnly = System.Text.RegularExpressions.Regex
                    .Replace(line, @"[^0-9A-Fa-f]", "");

                if (hexOnly.Length != 16) return null;

                byte[] bytes = new byte[8];
                for (int i = 0; i < 8; i++)
                    bytes[i] = Convert.ToByte(hexOnly.Substring(i * 2, 2), 16);

                return bytes;
            }
            catch { return null; }
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — PARSE A LOG LINE TO A 26-BYTE FRAME
        // ══════════════════════════════════════════════════════════════════
        private byte[] ParseLineToFrame(string line)
        {
            try
            {
                // strip timestamp prefix  "HH:MM:SS.mmm : xx xx xx ..."
                string hex = line.Trim();
                if (hex.Contains(":"))
                {
                    int lastColon = hex.LastIndexOf(':');
                    string after  = hex.Substring(lastColon + 1).Trim();
                    if (after.Length > 0) hex = after;
                }

                string[] tokens = hex.Split(
                    new char[] { ' ', '\t', '-', ',' },
                    StringSplitOptions.RemoveEmptyEntries);

                var result = new List<byte>();
                foreach (string tok in tokens)
                {
                    string t = tok.Trim();
                    if (string.IsNullOrEmpty(t)) continue;

                    if (t.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                        result.Add(Convert.ToByte(t.Substring(2), 16));
                    else if (t.Length == 2)
                        result.Add(Convert.ToByte(t, 16));
                }

                return result.Count == FRAME_SIZE ? result.ToArray() : null;
            }
            catch { return null; }
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — HEADER LINE FILTER
        // ══════════════════════════════════════════════════════════════════
        private bool IsHeaderLine(string line)
        {
            string t = line.Trim();
            return t.StartsWith("=")   ||
                   t.StartsWith("-")   ||
                   t.StartsWith("#")   ||
                   t.StartsWith("//")  ||
                   t.StartsWith("PORT") ||
                   t.StartsWith("Date") ||
                   t.StartsWith("Time") ||
                   t.StartsWith("Frame") ||
                   t.StartsWith("Timestamp");
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — CONTAINS COMMAND (duplicate guard within one loop)
        // ══════════════════════════════════════════════════════════════════
        private bool ContainsCommand(List<LogCommand> list, byte[] command8)
        {
            foreach (var c in list)
                if (CompareByteArrays(c.Command8, command8)) return true;
            return false;
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — EXTRACT LOOPS FROM SINGLE FILE
        // ══════════════════════════════════════════════════════════════════
        private void ExtractLoopsFromSingleFile(
            string filePath,
            byte[] firstExpectedCmd,
            byte[] lastExpectedCmd,
            out List<List<LogCommand>> mainLoops,
            out List<List<LogCommand>> redundantLoops)
        {
            mainLoops      = new List<List<LogCommand>>();
            redundantLoops = new List<List<LogCommand>>();

            List<LogCommand> currentMainLoop = null;
            List<LogCommand> currentRedLoop  = null;
            bool mainStarted = false;
            bool redStarted  = false;

            bool firstEqualsLast =
                CompareByteArrays(firstExpectedCmd, lastExpectedCmd);

            try
            {
                foreach (string line in File.ReadAllLines(filePath))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    if (IsHeaderLine(line)) continue;

                    byte[] frame = ParseLineToFrame(line);
                    if (frame == null) continue;

                    byte[] command8 = new byte[CMD_SIZE];
                    Array.Copy(frame, CMD_OFFSET, command8, 0, CMD_SIZE);
                    if (IsAllZeros(command8)) continue;

                    // determine MAIN vs REDUNDANT from source bit
                    bool   isMain    = (frame[SOURCE_BYTE_INDEX] & SOURCE_BIT_MASK) == 0;
                    int    frameCnt  = (frame[2] << 8) | frame[3];
                    bool   isFirst   = CompareByteArrays(command8, firstExpectedCmd);
                    bool   isLast    = CompareByteArrays(command8, lastExpectedCmd);

                    var cmd = new LogCommand
                    {
                        Timestamp   = DateTime.Now,
                        Port        = line.Contains("[PORT1]") ? "PORT1" : "PORT2",
                        FrameNumber = frameCnt,
                        Command8    = command8,
                        FullFrame26 = frame,
                        Source      = isMain ? "MAIN" : "REDUNDANT",
                        HexString   = BitConverter.ToString(command8).Replace("-", " "),
                        OriginalLine = line
                    };

                    if (isMain)
                    {
                        if (firstEqualsLast)
                        {
                            if (isFirst)
                            {
                                if (currentMainLoop != null &&
                                    currentMainLoop.Count > 0)
                                    mainLoops.Add(currentMainLoop);
                                currentMainLoop = new List<LogCommand>();
                                mainStarted     = true;
                            }
                            if (mainStarted && currentMainLoop != null)
                                if (!ContainsCommand(currentMainLoop, command8))
                                    currentMainLoop.Add(cmd);
                        }
                        else
                        {
                            if (isFirst)
                            {
                                currentMainLoop = new List<LogCommand>();
                                mainStarted     = true;
                            }
                            if (mainStarted && currentMainLoop != null)
                            {
                                if (!ContainsCommand(currentMainLoop, command8))
                                    currentMainLoop.Add(cmd);

                                if (isLast && !isFirst &&
                                    currentMainLoop.Count > 0)
                                {
                                    mainLoops.Add(currentMainLoop);
                                    currentMainLoop = null;
                                    mainStarted     = false;
                                }
                            }
                        }
                    }
                    else  // REDUNDANT
                    {
                        if (firstEqualsLast)
                        {
                            if (isFirst)
                            {
                                if (currentRedLoop != null &&
                                    currentRedLoop.Count > 0)
                                    redundantLoops.Add(currentRedLoop);
                                currentRedLoop = new List<LogCommand>();
                                redStarted     = true;
                            }
                            if (redStarted && currentRedLoop != null)
                                if (!ContainsCommand(currentRedLoop, command8))
                                    currentRedLoop.Add(cmd);
                        }
                        else
                        {
                            if (isFirst)
                            {
                                currentRedLoop = new List<LogCommand>();
                                redStarted     = true;
                            }
                            if (redStarted && currentRedLoop != null)
                            {
                                if (!ContainsCommand(currentRedLoop, command8))
                                    currentRedLoop.Add(cmd);

                                if (isLast && !isFirst &&
                                    currentRedLoop.Count > 0)
                                {
                                    redundantLoops.Add(currentRedLoop);
                                    currentRedLoop = null;
                                    redStarted     = false;
                                }
                            }
                        }
                    }
                }

                // flush incomplete loops
                if (currentMainLoop != null && currentMainLoop.Count > 0)
                    mainLoops.Add(currentMainLoop);
                if (currentRedLoop  != null && currentRedLoop.Count  > 0)
                    redundantLoops.Add(currentRedLoop);
            }
            catch (Exception ex)
            {
                throw new Exception(
                    string.Format("ExtractLoopsFromSingleFile failed:\n{0}",
                        ex.Message));
            }
        }

      
        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — EXTRACT LOOPS FROM FILE  (multiple-file mode)
        // ══════════════════════════════════════════════════════════════════
        private List<List<LogCommand>> ExtractLoopsFromFile(
            string filePath,
            string expectedSource,
            byte[] firstExpectedCmd,
            byte[] lastExpectedCmd)
        {
            var allLoops = new List<List<LogCommand>>();

            try
            {
                string[] lines = File.ReadAllLines(filePath);

                // ── pass 1: verify first/last commands exist ──────────────────
                bool firstExists = false, lastExists = false;

                // Check if commands exist with expected source byte
                foreach (string scanLine in lines)
                {
                    if (string.IsNullOrWhiteSpace(scanLine)) continue;
                    if (IsHeaderLine(scanLine)) continue;

                    byte[] frame = ParseLineToFrame(scanLine);
                    if (frame == null) continue;

                    byte[] command8 = new byte[CMD_SIZE];
                    Array.Copy(frame, CMD_OFFSET, command8, 0, CMD_SIZE);
                    if (IsAllZeros(command8)) continue;

                    bool isMain = (frame[SOURCE_BYTE_INDEX] & SOURCE_BIT_MASK) == 0;
                    string detSrc = isMain ? "MAIN" : "REDUNDANT";

                    if (detSrc == expectedSource)
                    {
                        if (CompareByteArrays(command8, firstExpectedCmd))
                            firstExists = true;
                        if (CompareByteArrays(command8, lastExpectedCmd))
                            lastExists = true;
                    }

                    if (firstExists && lastExists) break;
                }

                // ── Determine active source mode ──────────────────────────────
                string activeSource = expectedSource;
                bool useSourceByte = true;

                // If markers not found with expected source, check if they exist at all
                if (!firstExists || !lastExists)
                {
                    firstExists = false;
                    lastExists = false;

                    foreach (string scanLine in lines)
                    {
                        if (string.IsNullOrWhiteSpace(scanLine)) continue;
                        if (IsHeaderLine(scanLine)) continue;

                        byte[] frame = ParseLineToFrame(scanLine);
                        if (frame == null) continue;

                        byte[] command8 = new byte[CMD_SIZE];
                        Array.Copy(frame, CMD_OFFSET, command8, 0, CMD_SIZE);
                        if (IsAllZeros(command8)) continue;

                        if (CompareByteArrays(command8, firstExpectedCmd))
                            firstExists = true;
                        if (CompareByteArrays(command8, lastExpectedCmd))
                            lastExists = true;

                        if (firstExists && lastExists) break;
                    }

                    // If found without source filtering, trust the file itself
                    if (firstExists && lastExists)
                    {
                        activeSource = "ANY";
                        useSourceByte = false;  // Don't filter by source byte
                    }
                }

                // If still no markers found, return empty
                if (!firstExists || !lastExists)
                    return allLoops;

                bool firstEqualsLast =
                    CompareByteArrays(firstExpectedCmd, lastExpectedCmd);

                // ── pass 2: extract loops ─────────────────────────────────────
                List<LogCommand> currentLoop = null;
                bool loopStarted = false;

                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    if (IsHeaderLine(line)) continue;

                    byte[] frame = ParseLineToFrame(line);
                    if (frame == null) continue;

                    byte[] command8 = new byte[CMD_SIZE];
                    Array.Copy(frame, CMD_OFFSET, command8, 0, CMD_SIZE);
                    if (IsAllZeros(command8)) continue;

                    bool isMain = (frame[SOURCE_BYTE_INDEX] & SOURCE_BIT_MASK) == 0;
                    string detSrc = isMain ? "MAIN" : "REDUNDANT";

                    // Apply source filtering only if useSourceByte is true
                    if (useSourceByte && activeSource != "ANY" && detSrc != activeSource)
                        continue;

                    // If not using source byte, assign based on expected source
                    if (!useSourceByte)
                        detSrc = expectedSource;

                    int frameCnt = (frame[2] << 8) | frame[3];
                    bool isFirst = CompareByteArrays(command8, firstExpectedCmd);
                    bool isLast = CompareByteArrays(command8, lastExpectedCmd);

                    var cmd = new LogCommand
                    {
                        Timestamp = DateTime.Now,
                        Port = line.Contains("[PORT1]") ? "PORT1" :
                                       line.Contains("PORT1") ? "PORT1" : "PORT2",
                        FrameNumber = frameCnt,
                        Command8 = command8,
                        FullFrame26 = frame,
                        Source = detSrc,
                        HexString = BitConverter.ToString(command8).Replace("-", " "),
                        OriginalLine = line
                    };

                    if (firstEqualsLast)
                    {
                        if (isFirst)
                        {
                            if (currentLoop != null && currentLoop.Count > 0)
                                allLoops.Add(currentLoop);
                            currentLoop = new List<LogCommand>();
                            loopStarted = true;
                        }
                        if (loopStarted && currentLoop != null)
                            if (!ContainsCommand(currentLoop, command8))
                                currentLoop.Add(cmd);
                    }
                    else
                    {
                        if (isFirst)
                        {
                            currentLoop = new List<LogCommand>();
                            loopStarted = true;
                        }

                        if (loopStarted && currentLoop != null)
                        {
                            if (!ContainsCommand(currentLoop, command8))
                                currentLoop.Add(cmd);

                            if (isLast && !isFirst && currentLoop.Count > 0)
                            {
                                allLoops.Add(currentLoop);
                                currentLoop = null;
                                loopStarted = false;
                            }
                        }
                    }
                }

                if (currentLoop != null && currentLoop.Count > 0)
                    allLoops.Add(currentLoop);
            }
            catch (Exception ex)
            {
                throw new Exception(
                    string.Format("ExtractLoopsFromFile failed [{0}]:\n{1}",
                        expectedSource, ex.Message));
            }

            return allLoops;
        }
        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — COMPARE LOOPS WITH EXPECTED
        // ══════════════════════════════════════════════════════════════════
        private OfflineComparisonResult CompareLoopsWithExpected(
            List<ExpectedCommand>       expectedCommands,
            List<List<LogCommand>>      mainLoops,
            List<List<LogCommand>>      redundantLoops)
        {
            var result = new OfflineComparisonResult
            {
                TotalExpected      = expectedCommands.Count,
                TotalMainLoops     = mainLoops.Count,
                TotalRedundantLoops = redundantLoops.Count
            };

            foreach (var expected in expectedCommands)
            {
                var cmdResult = new ExpectedCommandResult
                {
                    Expected           = expected,
                    FoundInMain        = false,
                    FoundInRedundant   = false,
                    MainLoopCount      = 0,
                    RedundantLoopCount = 0,
                    Status             = "FAILED",
                    MatchedSource      = "NONE"
                };

                // check every MAIN loop
                for (int li = 0; li < mainLoops.Count; li++)
                {
                    foreach (var cmd in mainLoops[li])
                    {
                        if (CompareByteArrays(expected.CommandBytes, cmd.Command8))
                        {
                            cmdResult.FoundInMain = true;
                            cmdResult.MainLoopCount++;
                            break;
                        }
                    }
                }

                // only check REDUNDANT if not found in MAIN
                if (!cmdResult.FoundInMain)
                {
                    for (int li = 0; li < redundantLoops.Count; li++)
                    {
                        foreach (var cmd in redundantLoops[li])
                        {
                            if (CompareByteArrays(expected.CommandBytes, cmd.Command8))
                            {
                                cmdResult.FoundInRedundant = true;
                                cmdResult.RedundantLoopCount++;
                                break;
                            }
                        }
                    }
                }

                if (cmdResult.FoundInMain)
                {
                    result.TotalFound++;
                    result.FoundInMain++;
                    cmdResult.Status        = "FOUND IN MAIN";
                    cmdResult.MatchedSource = "MAIN";
                }
                else if (cmdResult.FoundInRedundant)
                {
                    result.TotalFound++;
                    result.FoundInRedundant++;
                    cmdResult.Status        = "FOUND IN REDUNDANT";
                    cmdResult.MatchedSource = "REDUNDANT";
                }
                else
                {
                    result.TotalMissing++;
                    cmdResult.Status        = "FAILED";
                    cmdResult.MatchedSource = "NONE";
                }

                result.CommandResults.Add(cmdResult);
            }

            result.SuccessRate = result.TotalExpected > 0
                ? (result.TotalFound * 100.0 / result.TotalExpected) : 0.0;

            return result;
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — GENERATE OFFLINE COMPARISON REPORT
        // ══════════════════════════════════════════════════════════════════
        private void GenerateOfflineComparisonReport(
            OfflineComparisonResult      compareResults,
            List<ExpectedCommand>        expectedCommands,
            List<List<LogCommand>>       mainLoops,
            List<List<LogCommand>>       redundantLoops)
        {
            var    sb      = new StringBuilder();
            string sepLine = new string('=', 80);
            int totalLoops = Math.Max(mainLoops.Count, redundantLoops.Count);

            sb.AppendLine(sepLine);
            sb.AppendLine("OFFLINE LOOP-BY-LOOP VERIFICATION REPORT");
            sb.AppendLine("Logic: MAIN checked first -> missing only checked in REDUNDANT");
            sb.AppendLine(sepLine);
            sb.AppendLine(string.Format(
                "Date          : {0:yyyy-MM-dd HH:mm:ss}", DateTime.Now));
            sb.AppendLine(string.Format(
                "Mode          : {0}",
                _isOfflineMultipleMode ? "Multiple Files" : "Single File"));
            sb.AppendLine();
            sb.AppendLine("FILES:");
            sb.AppendLine(string.Format(
                "  Expected    : {0}", Path.GetFileName(_offlineExpectedFile)));

            if (_isOfflineMultipleMode)
            {
                sb.AppendLine(string.Format(
                    "  Main        : {0}", Path.GetFileName(_offlineMainFile)));
                sb.AppendLine(string.Format(
                    "  Redundant   : {0}", Path.GetFileName(_offlineRedundantFile)));
            }
            else
            {
                sb.AppendLine(string.Format(
                    "  Log File    : {0}", Path.GetFileName(_offlineSingleFile)));
            }

            sb.AppendLine();

            string firstCmdHex = expectedCommands.Count > 0
                ? BitConverter.ToString(expectedCommands[0].CommandBytes)
                    .Replace("-", " ") : "N/A";
            string lastCmdHex  = expectedCommands.Count > 0
                ? BitConverter.ToString(
                    expectedCommands[expectedCommands.Count - 1].CommandBytes)
                    .Replace("-", " ") : "N/A";

            sb.AppendLine("LOOP DETECTION:");
            sb.AppendLine(string.Format("  Loop Start   : {0}", firstCmdHex));
            sb.AppendLine(string.Format("  Loop End     : {0}", lastCmdHex));
            sb.AppendLine(string.Format("  MAIN Loops   : {0}", mainLoops.Count));
            sb.AppendLine(string.Format("  RED Loops    : {0}", redundantLoops.Count));
            sb.AppendLine(string.Format("  Total Loops  : {0}", totalLoops));
            sb.AppendLine(string.Format("  Expected Cmds: {0}", expectedCommands.Count));
            sb.AppendLine();
            sb.AppendLine(sepLine);
            sb.AppendLine();
            sb.AppendLine("LOOP BY LOOP VERIFICATION");
            sb.AppendLine("MAIN checked first -> missing only checked in REDUNDANT");
            sb.AppendLine();

            int passCount    = 0;
            int failCount    = 0;
            int totalFromRed = 0;

            for (int loopIdx = 0; loopIdx < totalLoops; loopIdx++)
            {
                int loopNum = loopIdx + 1;

                List<LogCommand> mainLoop =
                    loopIdx < mainLoops.Count      ? mainLoops[loopIdx]      : null;
                List<LogCommand> redLoop  =
                    loopIdx < redundantLoops.Count ? redundantLoops[loopIdx] : null;

                var foundInMain   = new List<string>();
                var takenFromRed  = new List<string>();
                var trulyMissing  = new List<string>();

                foreach (var expected in expectedCommands)
                {
                    string cmdHex =
                        BitConverter.ToString(expected.CommandBytes)
                            .Replace("-", " ");

                    bool inMain = false;
                    if (mainLoop != null)
                    {
                        foreach (var cmd in mainLoop)
                        {
                            if (CompareByteArrays(expected.CommandBytes, cmd.Command8))
                            { inMain = true; break; }
                        }
                    }

                    if (inMain)
                    {
                        foundInMain.Add(cmdHex);
                    }
                    else
                    {
                        bool inRed = false;
                        if (redLoop != null)
                        {
                            foreach (var cmd in redLoop)
                            {
                                if (CompareByteArrays(
                                    expected.CommandBytes, cmd.Command8))
                                { inRed = true; break; }
                            }
                        }

                        if (inRed)
                        {
                            takenFromRed.Add(cmdHex);
                            totalFromRed++;
                        }
                        else
                        {
                            trulyMissing.Add(cmdHex);
                        }
                    }
                }

                bool loopPass  = (trulyMissing.Count == 0);
                if (loopPass) passCount++;
                else          failCount++;

                int totalFound = foundInMain.Count + takenFromRed.Count;

                if (loopPass && takenFromRed.Count == 0)
                {
                    sb.AppendLine(string.Format(
                        "LOOP {0:D3} : {1}/{2}  PASS  [All from MAIN]",
                        loopNum, totalFound, expectedCommands.Count));
                }
                else if (loopPass && takenFromRed.Count > 0)
                {
                    sb.AppendLine(string.Format(
                        "LOOP {0:D3} : {1}/{2}  PASS  [MAIN:{3}  RED:{4}]",
                        loopNum, totalFound, expectedCommands.Count,
                        foundInMain.Count, takenFromRed.Count));

                    sb.Append("  From REDUNDANT: ");
                    for (int i = 0; i < Math.Min(5, takenFromRed.Count); i++)
                    {
                        sb.Append(takenFromRed[i]);
                        if (i < Math.Min(5, takenFromRed.Count) - 1)
                            sb.Append(" | ");
                    }
                    if (takenFromRed.Count > 5)
                        sb.Append(string.Format(
                            " ... +{0} more", takenFromRed.Count - 5));
                    sb.AppendLine();
                }
                else
                {
                    sb.AppendLine(string.Format(
                        "LOOP {0:D3} : {1}/{2}  FAIL  [MAIN:{3}  RED:{4}  MISS:{5}]",
                        loopNum, totalFound, expectedCommands.Count,
                        foundInMain.Count, takenFromRed.Count,
                        trulyMissing.Count));

                    sb.Append("  MISSING: ");
                    for (int i = 0; i < Math.Min(5, trulyMissing.Count); i++)
                    {
                        sb.Append(trulyMissing[i]);
                        if (i < Math.Min(5, trulyMissing.Count) - 1)
                            sb.Append(" | ");
                    }
                    if (trulyMissing.Count > 5)
                        sb.Append(string.Format(
                            " ... +{0} more", trulyMissing.Count - 5));
                    sb.AppendLine();

                    if (takenFromRed.Count > 0)
                    {
                        sb.Append("  From REDUNDANT: ");
                        for (int i = 0; i < Math.Min(5, takenFromRed.Count); i++)
                        {
                            sb.Append(takenFromRed[i]);
                            if (i < Math.Min(5, takenFromRed.Count) - 1)
                                sb.Append(" | ");
                        }
                        if (takenFromRed.Count > 5)
                            sb.Append(string.Format(
                                " ... +{0} more", takenFromRed.Count - 5));
                        sb.AppendLine();
                    }
                }
            }

            sb.AppendLine();
            sb.AppendLine(string.Format(
                "RESULT: {0}/{1} PASS  |  {2}/{1} FAIL",
                passCount, totalLoops, failCount));
            sb.AppendLine();
            sb.AppendLine(sepLine);
            sb.AppendLine();

            bool overallPerfect = (passCount == totalLoops && totalLoops > 0);

            sb.AppendLine("FINAL SUMMARY");
            sb.AppendLine();
            sb.AppendLine(string.Format(
                "Expected Commands : {0}", expectedCommands.Count));
            sb.AppendLine(string.Format(
                "Total Loops       : {0}", totalLoops));
            sb.AppendLine(string.Format(
                "PASS Loops        : {0}", passCount));
            sb.AppendLine(string.Format(
                "FAIL Loops        : {0}", failCount));
            sb.AppendLine(string.Format(
                "Success Rate      : {0:F2}%",
                totalLoops > 0 ? (passCount * 100.0 / totalLoops) : 0));

            if (totalFromRed > 0)
            {
                sb.AppendLine();
                sb.AppendLine(string.Format(
                    "Commands recovered from REDUNDANT : {0}", totalFromRed));
            }

            sb.AppendLine();
            sb.AppendLine(sepLine);
            sb.AppendLine();

            if (overallPerfect)
            {
                sb.AppendLine("VERIFICATION SUCCESSFUL");
                sb.AppendLine();
                sb.AppendLine(string.Format(
                    "{0} commands x {1} loops = PASS",
                    expectedCommands.Count, totalLoops));
                if (totalFromRed > 0)
                    sb.AppendLine(string.Format(
                        "Note: {0} command(s) recovered from REDUNDANT",
                        totalFromRed));
            }
            else
            {
                sb.AppendLine("VERIFICATION FAILED");
                sb.AppendLine();
                sb.AppendLine(string.Format(
                    "{0}/{1} loops FAILED", failCount, totalLoops));
                if (totalFromRed > 0)
                    sb.AppendLine(string.Format(
                        "Note: {0} command(s) recovered from REDUNDANT",
                        totalFromRed));
            }

            sb.AppendLine();
            sb.AppendLine(sepLine);
            sb.AppendLine();
            sb.AppendLine("LEGEND:");
            sb.AppendLine(
                "  PASS [All from MAIN]  = All commands found in MAIN");
            sb.AppendLine(
                "  PASS [MAIN:X  RED:Y]  = X from MAIN, Y recovered from REDUNDANT");
            sb.AppendLine(
                "  FAIL [MISS:N]         = N commands not found in MAIN or REDUNDANT");
            sb.AppendLine();
            sb.AppendLine(sepLine);

            txtOfflineCompareResults.Text = sb.ToString();
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — SHOW FINAL SUMMARY MESSAGE BOX
        // ══════════════════════════════════════════════════════════════════
        private void ShowFinalSummaryMessageBox(
            int    finalPassCount,
            int    finalFailCount,
            int    finalTotalLoops,
            int    finalExpectedCount,
            int    finalTookFromRed,
            double finalElapsed)
        {
            bool overallPerfect =
                (finalPassCount == finalTotalLoops && finalTotalLoops > 0);

            var summary = new StringBuilder();

            summary.AppendLine("FINAL SUMMARY");
            summary.AppendLine(new string('=', 50));
            summary.AppendLine();
            summary.AppendLine(string.Format(
                "Expected Commands : {0}", finalExpectedCount));
            summary.AppendLine(string.Format(
                "Total Loops       : {0}", finalTotalLoops));
            summary.AppendLine(string.Format(
                "PASS Loops        : {0}", finalPassCount));
            summary.AppendLine(string.Format(
                "FAIL Loops        : {0}", finalFailCount));
            summary.AppendLine(string.Format(
                "Success Rate      : {0:F2}%",
                finalTotalLoops > 0
                    ? (finalPassCount * 100.0 / finalTotalLoops) : 0));

            if (finalTookFromRed > 0)
            {
                summary.AppendLine();
                summary.AppendLine(string.Format(
                    "Recovered from REDUNDANT : {0} command(s)",
                    finalTookFromRed));
            }

            summary.AppendLine();
            summary.AppendLine(new string('=', 50));
            summary.AppendLine();

            if (overallPerfect)
            {
                summary.AppendLine("VERIFICATION SUCCESSFUL");
                summary.AppendLine();
                summary.AppendLine(string.Format(
                    "{0} commands x {1} loops = PASS",
                    finalExpectedCount, finalTotalLoops));
            }
            else
            {
                summary.AppendLine("VERIFICATION FAILED");
                summary.AppendLine();
                summary.AppendLine(string.Format(
                    "{0}/{1} loops FAILED", finalFailCount, finalTotalLoops));
                summary.AppendLine(
                    "See detailed report for missing commands.");
            }

            summary.AppendLine();
            summary.AppendLine(new string('=', 50));
            summary.AppendLine(string.Format(
                "Processing Time : {0:F2} seconds", finalElapsed));

            MessageBox.Show(
                summary.ToString(),
                "Offline Comparison Complete",
                MessageBoxButton.OK,
                overallPerfect
                    ? MessageBoxImage.Information
                    : MessageBoxImage.Warning);
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 5 — BYTE COMPARISON HELPERS
        // ══════════════════════════════════════════════════════════════════
        private bool CompareByteArrays(byte[] a, byte[] b)
        {
            if (a == null || b == null || a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
                if (a[i] != b[i]) return false;
            return true;
        }

        private bool IsAllZeros(byte[] data)
        {
            if (data == null) return true;
            foreach (byte b in data)
                if (b != 0) return false;
            return true;
        }


        // ══════════════════════════════════════════════════════════════════
        //  TAB 6 — ANALYZE FILE (WITH GAP DETECTION)
        // ══════════════════════════════════════════════════════════════════
        // ══════════════════════════════════════════════════════════════════
        //  TAB 6 — FRAME COUNTER MODE SWITCHING
        // ══════════════════════════════════════════════════════════════════
        private bool _isFrameCounterOfflineMode = false;
        private string _frameCounterFilePath = string.Empty;

        private void BtnFrameCounterOnline_Click(object sender, RoutedEventArgs e)
        {
            _isFrameCounterOfflineMode = false;

            panelFrameCounterOnline.Visibility = Visibility.Visible;
            panelFrameCounterOffline.Visibility = Visibility.Collapsed;

            btnFrameCounterOnline.Background = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF00D4FF"));
            btnFrameCounterOnline.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF0A0A1A"));

            btnFrameCounterOffline.Background = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF0A1628"));
            btnFrameCounterOffline.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF95A5A6"));

            txtFrameCounterModeDesc.Text = "Real-time frame counting from connected ports";
            _vm.AppStatus = "Frame Counter: ONLINE mode";
        }

        private void BtnFrameCounterOffline_Click(object sender, RoutedEventArgs e)
        {
            _isFrameCounterOfflineMode = true;

            panelFrameCounterOnline.Visibility = Visibility.Collapsed;
            panelFrameCounterOffline.Visibility = Visibility.Visible;

            btnFrameCounterOffline.Background = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF00D4FF"));
            btnFrameCounterOffline.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF0A0A1A"));

            btnFrameCounterOnline.Background = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF0A1628"));
            btnFrameCounterOnline.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString("#FF95A5A6"));

            txtFrameCounterModeDesc.Text = "Analyze frame statistics from log files";
            _vm.AppStatus = "Frame Counter: OFFLINE mode";
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 6 — BROWSE FILE
        // ══════════════════════════════════════════════════════════════════
        private void BtnBrowseFrameCounterFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Title = "Select Log File for Frame Analysis";
            dlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            dlg.InitialDirectory = _logFolderPath;

            if (dlg.ShowDialog() == true)
            {
                _frameCounterFilePath = dlg.FileName;
                txtFrameCounterFilePath.Text = _frameCounterFilePath;
                btnAnalyzeFrameFile.IsEnabled = true;
                _vm.AppStatus = "File selected: " + Path.GetFileName(_frameCounterFilePath);
            }
        }


        // ══════════════════════════════════════════════════════════════════
        //  TAB 6 — ANALYZE FILE BUTTON CLICK
        // ══════════════════════════════════════════════════════════════════
        private string _currentFrameCounterLogPath = string.Empty;

        private void BtnAnalyzeFrameFile_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_frameCounterFilePath) ||
                !File.Exists(_frameCounterFilePath))
            {
                MessageBox.Show("Please select a valid log file first!",
                    "File Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ── Disable button immediately ────────────────────────────────
            btnAnalyzeFrameFile.IsEnabled = false;
            btnAnalyzeFrameFile.Content = "ANALYZING...";

            // ── Hide old results / show placeholder ───────────────────────
            panelFrameCounterPlaceholder.Visibility = Visibility.Visible;
            scrollFrameCounterResults.Visibility = Visibility.Collapsed;
           // txtFrameCounterSummary.Text = string.Empty;
         //   txtFrameCounterDetailedLog.Text = string.Empty;

            DateTime startTime = DateTime.Now;
            string capturedFilePath = _frameCounterFilePath;

            // ── ALL heavy work on background thread ───────────────────────
            System.Threading.Thread workerThread =
                new System.Threading.Thread(delegate()
                {
                    string summaryText = string.Empty;
                    string detailText = string.Empty;
                    string logFilePath = string.Empty;
                    string errorMessage = string.Empty;
                    bool success = false;

                    try
                    {
                        // ── STEP 1: Analysis (background) ─────────────────
                        FrameCounterResult result =
                            AnalyzeLogFileWithGaps(capturedFilePath);

                        TimeSpan elapsed = DateTime.Now - startTime;

                        summaryText = result.SummaryLog;
                        //detailText = result.DetailedLog;
                        logFilePath = result.LogFilePath;

                        success = true;

                        // ── STEP 2: Update UI (dispatcher) ────────────────
                        Dispatcher.Invoke(new Action(delegate
                        {
                            _currentFrameCounterLogPath = result.LogFilePath;

                            // ★ Display SINGLE combined log (no duplicate)
                            txtFrameCounterDisplay.Text = result.GuiLog;

                            // Show results panel
                            panelFrameCounterPlaceholder.Visibility = Visibility.Collapsed;
                            scrollFrameCounterResults.Visibility = Visibility.Visible;

                            btnAnalyzeFrameFile.IsEnabled = true;
                            btnAnalyzeFrameFile.Content = "ANALYZE";

                            _vm.AppStatus = string.Format(
                                "Analysis complete | Frames={0:N0} Gaps={1} Loss={2:F2}%",
                                result.MainValid + result.RedValid,
                                result.GapCount,
                                result.FrameLossPercentage);

                            MessageBox.Show(
                                string.Format(
                                    "ANALYSIS COMPLETE\n\n" +
                                    "Valid Frames : {0:N0}\n" +
                                    "  MAIN       : {1:N0}\n" +
                                    "  REDUNDANT  : {2:N0}\n\n" +
                                    "Frame Gaps   : {3}\n" +
                                    "Missed Frames: {4:N0}\n" +
                                    "Frame Loss   : {5:F2}%\n\n" +
                                    "Full log saved to:\n{6}",
                                    result.MainValid + result.RedValid,
                                    result.MainValid,
                                    result.RedValid,
                                    result.GapCount,
                                    result.TotalMissedFrames,
                                    result.FrameLossPercentage,
                                    Path.GetFileName(result.LogFilePath)),
                                "Frame Analysis Complete",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
                        }));
                    }
                    catch (Exception ex)
                    {
                        errorMessage = ex.Message;
                        Dispatcher.Invoke(new Action(delegate
                        {
                            btnAnalyzeFrameFile.IsEnabled = true;
                            btnAnalyzeFrameFile.Content = "ANALYZE";
                            _vm.AppStatus = "Analysis failed: " + errorMessage;
                            MessageBox.Show(
                                "Error analyzing file:\n\n" + errorMessage,
                                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }));
                    }
                });

            workerThread.IsBackground = true;
            workerThread.Start();
        }

        // ══════════════════════════════════════════════════════════════════
        //  TAB 6 — ANALYZE LOG FILE WITH GAPS
        // ══════════════════════════════════════════════════════════════════
        private FrameCounterResult AnalyzeLogFileWithGaps(string filePath)
        {
            var result = new FrameCounterResult();

            FileInfo fi = new FileInfo(filePath);
            result.FileSize = fi.Length;

            List<ushort> frameCounters = new List<ushort>();
            List<int> differences = new List<int>();
            List<GapInfo> gaps = new List<GapInfo>();

            int lineNumber = 0;
            int validFrames = 0;
            int invalidFrames = 0;

            // ── Two separate logs ─────────────────────────────────────────
            // detailedOutput = ALL lines (written to file only)
            // gapOutput      = ONLY gaps/errors (shown in GUI)
            StringBuilder detailedOutput = new StringBuilder();
            StringBuilder gapOutput = new StringBuilder();

            ushort previousFrameCounter = 0;
            bool isFirstFrame = true;
            string headerLine1 = string.Empty;
            string headerLine2 = string.Empty;
            string headerLine3 = string.Empty;

            using (StreamReader sr = new StreamReader(filePath, Encoding.UTF8, false, 65536))
            {
                string line;

                while ((line = sr.ReadLine()) != null)
                {
                    lineNumber++;
                    line = line.Trim();

                    if (string.IsNullOrWhiteSpace(line)) continue;

                    // Skip header lines
                    if (line.StartsWith("═") ||
                        line.StartsWith("─") ||
                        line.StartsWith("26 BYTES") ||
                        line.StartsWith("Created") ||
                        line.StartsWith("Mode") ||
                        line.StartsWith("Buffer") ||
                        line.StartsWith("Full Frame") ||
                        line.StartsWith("#") ||
                        line.StartsWith("//") ||
                        line.StartsWith("Docklight"))
                        continue;

                    // Parse hex bytes
                    byte[] frame26 = ParseHexLine_Old(line);

                    // Validate length
                    if (frame26.Length != 26)
                    {
                        if (frame26.Length > 0)
                        {
                            invalidFrames++;
                            gaps.Add(new GapInfo
                            {
                                Source = "?",
                                FrameCounterRange = "N/A",
                                LineNumber = lineNumber,
                                Remarks = "Invalid Frame Length"
                            });

                            string invalidLine = string.Format(
                                "Line {0,6}: Invalid frame length ({1} bytes). Skipping.",
                                lineNumber, frame26.Length);

                            // Write to file only
                            detailedOutput.AppendLine(invalidLine);
                            // Also show in GUI (it's a problem)
                            gapOutput.AppendLine(invalidLine);
                        }
                        continue;
                    }

                    // Validate markers
                    if (frame26[0] != 0x0D || frame26[1] != 0x0A || frame26[25] != 0xAB)
                    {
                        invalidFrames++;
                        gaps.Add(new GapInfo
                        {
                            Source = "?",
                            FrameCounterRange = "N/A",
                            LineNumber = lineNumber,
                            Remarks = "Invalid Frame Markers"
                        });

                        string markerLine = string.Format(
                            "Line {0,6}: Invalid Frame Markers. " +
                            "Start: 0x{1:X2} 0x{2:X2}, End: 0x{3:X2}. Skipping.",
                            lineNumber, frame26[0], frame26[1], frame26[25]);

                        // Write to file only
                        detailedOutput.AppendLine(markerLine);
                        // Also show in GUI (it's a problem)
                        gapOutput.AppendLine(markerLine);
                        continue;
                    }

                    // Extract Frame Counter bytes 3-4 (Big-endian)
                    ushort currentFrameCounter = (ushort)((frame26[3] << 8) | frame26[4]);

                    frameCounters.Add(currentFrameCounter);
                    validFrames++;

                    // MAIN / RED stats
                    bool isMain = (frame26[SOURCE_BYTE_INDEX] & SOURCE_BIT_MASK) == 0;
                    if (isMain) result.MainValid++;
                    else result.RedValid++;

                    if (isFirstFrame)
                    {
                        // ── Header lines (shown in GUI + file) ───────────
                        headerLine1 = string.Format(
                            "Line {0,6}: First Frame Counter = {1} (0x{1:X4})",
                            lineNumber, currentFrameCounter);

                        headerLine2 = string.Format(
                            "{0,-7} {1,-8} {2,-10} {3,-10} {4,-10}",
                            "Line", "Prev FC", "Curr FC", "Diff", "Status");

                        headerLine3 = new string('-', 60);

                        detailedOutput.AppendLine(headerLine1);
                        detailedOutput.AppendLine(headerLine2);
                        detailedOutput.AppendLine(headerLine3);

                        isFirstFrame = false;
                    }
                    else
                    {
                        int rawDifference = currentFrameCounter - previousFrameCounter;
                        if (rawDifference < 0) rawDifference += 65536;

                        int difference = rawDifference - 1;
                        differences.Add(difference);

                        string status;

                        if (rawDifference == 1)
                        {
                            status = "OK";

                            // ── OK lines: write to FILE ONLY ─────────────
                            detailedOutput.AppendLine(string.Format(
                                "Line {0,6}: {1,5} (0x{1:X4}) -> {2,5} (0x{2:X4}) | Diff: {3,5} {4}",
                                lineNumber,
                                previousFrameCounter,
                                currentFrameCounter,
                                rawDifference - 1,
                                status));
                        }
                        else if (rawDifference == 0)
                        {
                            status = "DUPLICATE";
                            gaps.Add(new GapInfo
                            {
                                Source = isMain ? "MAIN" : "RED",
                                FrameCounterRange = string.Format("{0}-{1}",
                                    previousFrameCounter + 1, currentFrameCounter - 1),
                                LineNumber = lineNumber,
                                Remarks = "Duplicate"
                            });

                            string gapLine = string.Format(
                                "Line {0,6}: {1,5} (0x{1:X4}) -> {2,5} (0x{2:X4}) | Diff: {3,5} {4}",
                                lineNumber,
                                previousFrameCounter,
                                currentFrameCounter,
                                rawDifference - 1,
                                status);

                            // Write to file
                            detailedOutput.AppendLine(gapLine);
                            // Show in GUI
                            gapOutput.AppendLine(gapLine);
                        }
                        else if (rawDifference > 1 && rawDifference < 100)
                        {
                            status = "GAP";
                            gaps.Add(new GapInfo
                            {
                                Source = isMain ? "MAIN" : "RED",
                                FrameCounterRange = string.Format("{0}-{1}",
                                    previousFrameCounter + 1, currentFrameCounter - 1),
                                LineNumber = lineNumber,
                                Remarks = "Frame missing"
                            });

                            string gapLine = string.Format(
                                "Line {0,6}: {1,5} (0x{1:X4}) -> {2,5} (0x{2:X4}) | Diff: {3,5} {4}",
                                lineNumber,
                                previousFrameCounter,
                                currentFrameCounter,
                                rawDifference - 1,
                                status);

                            // Write to file
                            detailedOutput.AppendLine(gapLine);
                            // Show in GUI
                            gapOutput.AppendLine(gapLine);
                        }
                        else
                        {
                            status = "ROLLOVER";
                            gaps.Add(new GapInfo
                            {
                                Source = isMain ? "MAIN" : "RED",
                                FrameCounterRange = string.Format("{0}-{1}",
                                    previousFrameCounter + 1, currentFrameCounter - 1),
                                LineNumber = lineNumber,
                                Remarks = "Rollover"
                            });

                            string gapLine = string.Format(
                                "Line {0,6}: {1,5} (0x{1:X4}) -> {2,5} (0x{2:X4}) | Diff: {3,5} {4}",
                                lineNumber,
                                previousFrameCounter,
                                currentFrameCounter,
                                rawDifference - 1,
                                status);

                            // Write to file
                            detailedOutput.AppendLine(gapLine);
                            // Show in GUI
                            gapOutput.AppendLine(gapLine);
                        }
                    }

                    previousFrameCounter = currentFrameCounter;
                }
            }

            // ── Build summary ─────────────────────────────────────────────
            StringBuilder summaryOutput = new StringBuilder();
            summaryOutput.AppendLine(new string('=', 80));
            summaryOutput.AppendLine("SUMMARY");
            summaryOutput.AppendLine(new string('=', 80));
            summaryOutput.AppendLine(string.Format(
                "Total Lines Read       : {0}", lineNumber));
            summaryOutput.AppendLine(string.Format(
                "Valid 26-byte Frames   : {0}", validFrames));
            summaryOutput.AppendLine(string.Format(
                "Invalid Frames         : {0}", invalidFrames));

            result.TotalLines = lineNumber;
            result.MainTotal = result.MainValid;
            result.RedTotal = result.RedValid;
            result.MainInvalid = invalidFrames;

            if (frameCounters.Count > 0)
            {
                result.FirstFrameCounter = frameCounters[0];
                result.LastFrameCounter = frameCounters[frameCounters.Count - 1];

                summaryOutput.AppendLine(string.Format(
                    "First Frame Counter    : {0} (0x{0:X4})", frameCounters[0]));
                summaryOutput.AppendLine(string.Format(
                    "Last Frame Counter     : {0} (0x{0:X4})",
                    frameCounters[frameCounters.Count - 1]));
            }

            if (differences.Count > 0)
            {
                int normalFrames = differences.Count(d => d == 0);
                int duplicates = differences.Count(d => d == -1);
                int gapCount = differences.Count(d => d > 0 && d < 99);
                int rollovers = differences.Count(d => d >= 99);
                int totalMissedFrames =
                    differences.Where(d => d > 0 && d < 99).Sum(d => d);

                result.GapCount = gapCount;
                result.TotalMissedFrames = totalMissedFrames;
                result.NormalIncrements = normalFrames;
                result.Duplicates = duplicates;
                result.Rollovers = rollovers;
                result.MinDifference = differences.Min();
                result.MaxDifference = differences.Max();
                result.AvgDifference = differences.Average();

                summaryOutput.AppendLine(new string('-', 80));
                summaryOutput.AppendLine(string.Format(
                    "Normal Increments (1)  : {0}", normalFrames));
                summaryOutput.AppendLine(string.Format(
                    "Duplicate Frames (0)   : {0}", duplicates));
                summaryOutput.AppendLine(string.Format(
                    "Frame Counter Gaps     : {0}", gapCount));

                if (gaps.Count > 0)
                {
                    summaryOutput.AppendLine(
                        "\t\tFRAME COUNTER No\tLine No\t\tRemarks");
                    int displayCount = Math.Min(gaps.Count, 100);
                    for (int i = 0; i < displayCount; i++)
                    {
                        summaryOutput.AppendLine(string.Format(
                            "\t{0}.\t{1}\t\t\t\t{2}\t\t\t{3}",
                            i + 1,
                            gaps[i].FrameCounterRange,
                            gaps[i].LineNumber,
                            gaps[i].Remarks));
                    }
                    if (gaps.Count > 100)
                    {
                        summaryOutput.AppendLine(
                            "\t... and " + (gaps.Count - 100) + " more gaps");
                    }
                }

                summaryOutput.AppendLine(string.Format(
                    "Rollovers              : {0}", rollovers));
                summaryOutput.AppendLine(string.Format(
                    "Total Missed Frames    : {0}", totalMissedFrames));
                summaryOutput.AppendLine(string.Format(
                    "Min Difference         : {0}", differences.Min()));
                summaryOutput.AppendLine(string.Format(
                    "Max Difference         : {0}", differences.Max()));
                summaryOutput.AppendLine(string.Format(
                    "Avg Difference         : {0:F2}", differences.Average()));

                if (validFrames > 0)
                {
                    double expectedFrames = frameCounters.Count + totalMissedFrames;
                    double lossPercentage =
                        (totalMissedFrames / expectedFrames) * 100.0;
                    result.FrameLossPercentage = lossPercentage;
                    summaryOutput.AppendLine(string.Format(
                        "Frame Loss Percentage  : {0:F2}%", lossPercentage));
                }
            }

            summaryOutput.AppendLine(new string('=', 80));

            // ── Build GUI display log ─────────────────────────────────────
            // Only: Summary + Header + GAP/DUPLICATE/ROLLOVER lines
            StringBuilder guiLog = new StringBuilder();
            guiLog.Append(summaryOutput.ToString());
            guiLog.AppendLine();

            // Add header section
            if (!string.IsNullOrEmpty(headerLine1))
            {
                guiLog.AppendLine(headerLine1);
                guiLog.AppendLine(headerLine2);
                guiLog.AppendLine(headerLine3);
            }

            // Add ONLY gap lines (no OK lines)
            if (gapOutput.Length > 0)
            {
                guiLog.Append(gapOutput.ToString());
            }
            else
            {
                guiLog.AppendLine();
                guiLog.AppendLine("No gaps detected - All frames received correctly.");
            }

            // ── Store for GUI display ─────────────────────────────────────
            // Store ONLY guiLog for display
            result.GuiLog = guiLog.ToString();
            result.SummaryLog = summaryOutput.ToString(); // kept for reference only

            // ── Save FULL log to file (summary + ALL lines) ───────────────
            string directory = Path.GetDirectoryName(filePath);
            if (string.IsNullOrEmpty(directory))
                directory = Environment.CurrentDirectory;

            string logFile = Path.Combine(
                directory,
                "FrameCounter_Analysis_" +
                DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt");

            // File = Summary + ALL detailed lines
            File.WriteAllText(
                logFile,
                summaryOutput.ToString() +
                Environment.NewLine +
                detailedOutput.ToString(),
                Encoding.UTF8);

            result.LogFilePath = logFile;

            return result;
        }
        // ══════════════════════════════════════════════════════════════════
        //  ParseHexLine_Old  (used by AnalyzeLogFileWithGaps)
        // ══════════════════════════════════════════════════════════════════
        private byte[] ParseHexLine_Old(string line)
        {
            List<byte> result = new List<byte>();
            try
            {
                // Strip timestamp prefix
                if (line.Contains(":"))
                {
                    int lastColon = line.LastIndexOf(':');
                    string after = line.Substring(lastColon + 1).Trim();
                    if (after.Length > 0) line = after;
                }

                string[] parts = line.Split(
                    new char[] { ' ', '\t', ',', '-' },
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (string part in parts)
                {
                    string clean = part.Trim();
                    if (string.IsNullOrEmpty(clean)) continue;

                    if (clean.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                        result.Add(Convert.ToByte(clean.Substring(2), 16));
                    else if (clean.Length == 2)
                        result.Add(Convert.ToByte(clean, 16));
                }
            }
            catch { }

            return result.ToArray();
        }

        // ══════════════════════════════════════════════════════════════════
        //  Open log file button
        // ══════════════════════════════════════════════════════════════════
        private void BtnOpenFrameCounterLog_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFrameCounterLogPath) ||
                !File.Exists(_currentFrameCounterLogPath))
            {
                MessageBox.Show("No log file available!",
                    "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                Process.Start("notepad.exe", _currentFrameCounterLogPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot open log:\n" + ex.Message,
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  Open log folder button
        // ══════════════════════════════════════════════════════════════════
        private void BtnOpenFrameCounterLogFolder_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFrameCounterLogPath))
            {
                MessageBox.Show("No log file available!",
                    "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                string folder = Path.GetDirectoryName(_currentFrameCounterLogPath);
                if (Directory.Exists(folder))
                    Process.Start("explorer.exe",
                        string.Format("\"{0}\"", folder));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot open folder:\n" + ex.Message,
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

     
        // ══════════════════════════════════════════════════════════════════
        //  TAB 6 — FRAME ANALYSIS RESULT CLASS
        // ══════════════════════════════════════════════════════════════════
        private class FrameCounterResult
        {
            public long MainTotal;
            public long MainValid;
            public long MainInvalid;
            public long RedTotal;
            public long RedValid;
            public long RedInvalid;
            public long TotalLines;
            public long FileSize;
            public int GapCount;
            public int TotalMissedFrames;
            public double FrameLossPercentage;
            public string LogFilePath;

            // GUI display - single string (no duplicate)
            public string GuiLog;

            // Summary only - for reference
            public string SummaryLog;

            public ushort? FirstFrameCounter;
            public ushort? LastFrameCounter;
            public int NormalIncrements;
            public int Duplicates;
            public int Rollovers;
            public int? MinDifference;
            public int? MaxDifference;
            public double AvgDifference;
        }

        private class GapInfo
        {
            public string FrameCounterRange { get; set; }
            public int LineNumber { get; set; }
            public string Remarks { get; set; }
            public string Source { get; set; } // "MAIN" or "RED"
        }


    
        // ══════════════════════════════════════════════════════════════════
        //  HELPER: DETERMINE FRAME STATUS
        // ══════════════════════════════════════════════════════════════════
        private string DetermineFrameStatus(int rawDiff)
        {
            if (rawDiff == 1)
                return "OK";
            else if (rawDiff == 0)
                return "DUPLICATE";
            else if (rawDiff > 1 && rawDiff < 100)
                return "GAP";
            else
                return "ROLLOVER";
        }

        // ══════════════════════════════════════════════════════════════════
        //  HELPER: VALIDATE FRAME
        // ══════════════════════════════════════════════════════════════════
        private bool ValidateFrame(byte[] frame)
        {
            if (frame == null || frame.Length != FRAME_SIZE) return false;
            if (frame[0] != HDR1 || frame[1] != HDR2) return false;
            if (frame[25] != FOOTER) return false;
            return true;
        }

      
    }
}