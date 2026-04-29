using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO.Ports;
using System.Threading;
using System.Windows.Threading;
using ANVESHA_TCRX_HEALTH_STATUS_GUI_V2.Models;

namespace ANVESHA_TCRX_HEALTH_STATUS_GUI_V2.ViewModels
{
    public class PortViewModel : INotifyPropertyChanged, IDisposable
    {
        // ══════════════════════════════════════════════════════════════════
        //  INFRASTRUCTURE
        // ══════════════════════════════════════════════════════════════════
        private readonly SerialPortManager _portManager;
        private readonly Dispatcher _dispatcher;
        private FileLogger _rawLogger;
        private FileLogger _payloadLogger;
        private FileLogger _rxLiveLogger;

        // ══════════════════════════════════════════════════════════════════
        //  BACKING FIELDS
        // ══════════════════════════════════════════════════════════════════
        private string _selectedPort = string.Empty;
        private int _selectedBaudRate = 230400;
        private bool _isConnected = false;
        private string _rawLog = string.Empty;
        private string _payloadLog = string.Empty;
        private string _rxLiveLog = string.Empty;
        private string _rxCompactLog = string.Empty;
        private string _statusText = "Not Connected";
        private string _portStatusColor = "#FFFF4444";
        private RxDataFields _latestFields = null;
        private long _rxFrameNumber = 0;

        // File paths (private setters to support OnPropertyChanged)
        private string _rawLogFilePath = string.Empty;
        private string _payloadLogFilePath = string.Empty;
        private string _rxLiveLogFilePath = string.Empty;

        // ══════════════════════════════════════════════════════════════════
        //  RING BUFFER LIMITS
        // ══════════════════════════════════════════════════════════════════
        private const int MAX_UI_LINES = 200;
        private const int MAX_COMPACT_LINES = 200;

        // ══════════════════════════════════════════════════════════════════
        //  UI RING BUFFER LISTS
        // ══════════════════════════════════════════════════════════════════
        private readonly List<string> _rawUiLines = new List<string>(220);
        private readonly List<string> _payloadUiLines = new List<string>(220);
        private readonly List<string> _rxCompactUiLines = new List<string>(220);
        // NOTE: NO _rxLiveUiLines — RxLiveLog stores only latest single entry

        // ══════════════════════════════════════════════════════════════════
        //  PENDING BATCH LISTS (written from background thread)
        // ══════════════════════════════════════════════════════════════════
        private readonly List<string> _pendingRawLines = new List<string>();
        private readonly List<string> _pendingPayloadLines = new List<string>();
        private readonly List<string> _pendingRxCompactLines = new List<string>();
        private string _pendingRxLiveLatest = null; // single latest
        private readonly object _pendingLock = new object();

        // ══════════════════════════════════════════════════════════════════
        //  BATCH UI UPDATE TIMER
        // ══════════════════════════════════════════════════════════════════
        private readonly DispatcherTimer _uiUpdateTimer;

        // ══════════════════════════════════════════════════════════════════
        //  EVENTS
        // ══════════════════════════════════════════════════════════════════
        public event EventHandler<StringEventArgs> PortStatusMessage;
        public event EventHandler ConnectionStateChanged;

        // ══════════════════════════════════════════════════════════════════
        //  CONSTRUCTOR
        // ══════════════════════════════════════════════════════════════════
        public PortViewModel(int portNumber, Dispatcher dispatcher)
        {
            PortNumber = portNumber;
            _dispatcher = dispatcher;

            _portManager = new SerialPortManager(portNumber);
            _portManager.FrameReceived += OnFrameReceived;
            _portManager.StatusChanged += OnStatusChanged;
            _portManager.ErrorOccurred += OnErrorOccurred;

            _selectedBaudRate = 230400;

            AvailablePorts = new ObservableCollection<string>(
                SerialPort.GetPortNames());

            if (AvailablePorts.Count > 0)
                _selectedPort = AvailablePorts[0];

            // ── Unique log paths per session ───────────────────────────────
            string baseDir = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Logs",
                DateTime.Now.ToString("yyyyMMdd_HHmmss"));

            // Use property setters (not auto-init) to trigger OnPropertyChanged
            RawLogFilePath = System.IO.Path.Combine(
                baseDir,
                string.Format("PORT{0}_RAW_FRAMES.txt", portNumber));

            PayloadLogFilePath = System.IO.Path.Combine(
                baseDir,
                string.Format("PORT{0}_PAYLOAD.txt", portNumber));

            RxLiveLogFilePath = System.IO.Path.Combine(
                baseDir,
                string.Format("PORT{0}_RXLIVE.txt", portNumber));

            // ── Batch UI update timer ──────────────────────────────────────
            // Start at 100ms, will adapt to 500ms when idle
            _uiUpdateTimer = new DispatcherTimer(
                DispatcherPriority.Background, dispatcher)
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _uiUpdateTimer.Tick += OnUiUpdateTimerTick;
        }

        // ══════════════════════════════════════════════════════════════════
        //  PUBLIC PROPERTIES
        // ══════════════════════════════════════════════════════════════════

        public int PortNumber { get; private set; }

        // ── File paths (with private setters for INPC support) ────────────
        public string RawLogFilePath
        {
            get { return _rawLogFilePath; }
            private set
            {
                _rawLogFilePath = value;
                OnPropertyChanged("RawLogFilePath");
            }
        }

        public string PayloadLogFilePath
        {
            get { return _payloadLogFilePath; }
            private set
            {
                _payloadLogFilePath = value;
                OnPropertyChanged("PayloadLogFilePath");
            }
        }

        public string RxLiveLogFilePath
        {
            get { return _rxLiveLogFilePath; }
            private set
            {
                _rxLiveLogFilePath = value;
                OnPropertyChanged("RxLiveLogFilePath");
            }
        }

        // ── Available COM ports ────────────────────────────────────────────
        public ObservableCollection<string> AvailablePorts { get; private set; }

        // ── Baud rate options ──────────────────────────────────────────────
        public int[] BaudRates
        {
            get
            {
                return new[]
                {
                    1200, 2400, 4800, 9600, 19200,
                    38400, 57600, 115200, 230400, 460800
                };
            }
        }

        // ── Selected COM port ──────────────────────────────────────────────
        public string SelectedPort
        {
            get { return _selectedPort; }
            set
            {
                if (_selectedPort == value) return;
                _selectedPort = value;
                OnPropertyChanged("SelectedPort");
            }
        }

        // ── Selected baud rate ─────────────────────────────────────────────
        public int SelectedBaudRate
        {
            get { return _selectedBaudRate; }
            set
            {
                if (_selectedBaudRate == value) return;
                _selectedBaudRate = value;
                OnPropertyChanged("SelectedBaudRate");
            }
        }

        // ── Connection state ───────────────────────────────────────────────
        public bool IsConnected
        {
            get { return _isConnected; }
            private set
            {
                if (_isConnected == value) return;
                _isConnected = value;
                PortStatusColor = value ? "#FF00FF88" : "#FFFF4444";
                OnPropertyChanged("IsConnected");
                OnPropertyChanged("IsNotConnected");
                OnPropertyChanged("ConnectionLabel");
                RaiseConnectionStateChanged();
            }
        }

        public bool IsNotConnected
        {
            get { return !_isConnected; }
        }

        public string ConnectionLabel
        {
            get { return _isConnected ? "CONNECTED" : "DISCONNECTED"; }
        }

        public string PortStatusColor
        {
            get { return _portStatusColor; }
            private set
            {
                if (_portStatusColor == value) return;
                _portStatusColor = value;
                OnPropertyChanged("PortStatusColor");
            }
        }

        // ── Frame counters ─────────────────────────────────────────────────
        public long TotalFrames
        {
            get { return _portManager != null ? _portManager.TotalFrames : 0L; }
        }

        public long ValidFrames
        {
            get { return _portManager != null ? _portManager.ValidFrames : 0L; }
        }

        public long InvalidFrames
        {
            get { return _portManager != null ? _portManager.InvalidFrames : 0L; }
        }

        // ── Status text ────────────────────────────────────────────────────
        public string StatusText
        {
            get { return _statusText; }
            set
            {
                if (_statusText == value) return;
                _statusText = value;
                OnPropertyChanged("StatusText");
            }
        }

        // ── UI Display logs ────────────────────────────────────────────────
        public string RawLog
        {
            get { return _rawLog; }
            private set
            {
                _rawLog = value;
                OnPropertyChanged("RawLog");
            }
        }

        public string PayloadLog
        {
            get { return _payloadLog; }
            private set
            {
                _payloadLog = value;
                OnPropertyChanged("PayloadLog");
            }
        }

        // ── RxLiveLog: LATEST SINGLE ENTRY only (no accumulation) ──────────
        public string RxLiveLog
        {
            get { return _rxLiveLog; }
            private set
            {
                _rxLiveLog = value;
                OnPropertyChanged("RxLiveLog");
            }
        }

        public string RxCompactLog
        {
            get { return _rxCompactLog; }
            private set
            {
                _rxCompactLog = value;
                OnPropertyChanged("RxCompactLog");
            }
        }

        // ── Latest parsed fields (for RX Compare tab) ──────────────────────
        public RxDataFields LatestFields
        {
            get { return _latestFields; }
            private set
            {
                _latestFields = value;
                OnPropertyChanged("LatestFields");
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  CONNECT
        // ══════════════════════════════════════════════════════════════════
        public bool Connect()
        {
            if (IsConnected) return true;

            if (string.IsNullOrEmpty(SelectedPort))
            {
                StatusText = string.Format(
                    "PORT {0}: No COM port selected!", PortNumber);
                return false;
            }

            // ══════════════════════════════════════════════════════════════
            // FIX 1: Reset frame counter
            // ══════════════════════════════════════════════════════════════
            Interlocked.Exchange(ref _rxFrameNumber, 0);

            // ══════════════════════════════════════════════════════════════
            // FIX 2: Clear all buffers and old data
            // ══════════════════════════════════════════════════════════════
            lock (_pendingLock)
            {
                _pendingRawLines.Clear();
                _pendingPayloadLines.Clear();
                _pendingRxCompactLines.Clear();
                _pendingRxLiveLatest = null;
            }

            _rawUiLines.Clear();
            _payloadUiLines.Clear();
            _rxCompactUiLines.Clear();

            RawLog = string.Empty;
            PayloadLog = string.Empty;
            RxLiveLog = string.Empty;
            RxCompactLog = string.Empty;
            LatestFields = null;

            // ══════════════════════════════════════════════════════════════
            // FIX 3: Create new log files with unique timestamp
            // ══════════════════════════════════════════════════════════════
            string baseDir = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Logs",
                DateTime.Now.ToString("yyyyMMdd_HHmmss"));

            RawLogFilePath = System.IO.Path.Combine(
                baseDir,
                string.Format("PORT{0}_RAW_FRAMES.txt", PortNumber));

            PayloadLogFilePath = System.IO.Path.Combine(
                baseDir,
                string.Format("PORT{0}_PAYLOAD.txt", PortNumber));

            RxLiveLogFilePath = System.IO.Path.Combine(
                baseDir,
                string.Format("PORT{0}_RXLIVE.txt", PortNumber));

            if (!OpenLoggers())
            {
                StatusText = string.Format(
                    "PORT {0}: Cannot create log files!", PortNumber);
                return false;
            }

            // ══════════════════════════════════════════════════════════════
            // FIX 4: Reset SerialPortManager counters
            // ══════════════════════════════════════════════════════════════
            if (_portManager != null)
            {
                // Reset internal counters (add this method to SerialPortManager)
                _portManager.ResetCounters();
            }

            bool ok = _portManager.Connect(
                SelectedPort, SelectedBaudRate,
                Parity.None, 8, StopBits.One);

            if (ok)
            {
                IsConnected = true;
                _uiUpdateTimer.Start();

                // ══════════════════════════════════════════════════════════
                // FIX 5: Refresh stats immediately to show zeros
                // ══════════════════════════════════════════════════════════
                RefreshStats();

                StatusText = string.Format(
                    "PORT {0}: Connected [{1}] @ {2} baud  RS422",
                    PortNumber, SelectedPort, SelectedBaudRate);
            }
            else
            {
                IsConnected = false;
                StatusText = string.Format(
                    "PORT {0}: FAILED to open [{1}]",
                    PortNumber, SelectedPort);
                CloseLoggers();
            }

            RaisePortStatusMessage(StatusText);
            return ok;
        }

        // ══════════════════════════════════════════════════════════════════
        //  DISCONNECT
        // ══════════════════════════════════════════════════════════════════
        public void Disconnect()
        {
            _uiUpdateTimer.Stop();

            if (_portManager != null)
                _portManager.Disconnect();

            CloseLoggers();
            IsConnected = false;

            StatusText = string.Format(
                "PORT {0}: Disconnected [{1}]  " +
                "Total={2}  Valid={3}  Invalid={4}",
                PortNumber, SelectedPort,
                TotalFrames, ValidFrames, InvalidFrames);

            RaisePortStatusMessage(StatusText);
        }

        // ══════════════════════════════════════════════════════════════════
        //  REFRESH PORTS
        // ══════════════════════════════════════════════════════════════════
        public void RefreshPorts()
        {
            if (IsConnected) return;

            string current = SelectedPort;
            AvailablePorts.Clear();

            foreach (string p in SerialPort.GetPortNames())
                AvailablePorts.Add(p);

            if (!string.IsNullOrEmpty(current) &&
                AvailablePorts.Contains(current))
                SelectedPort = current;
            else if (AvailablePorts.Count > 0)
                SelectedPort = AvailablePorts[0];
            else
                SelectedPort = string.Empty;
        }

        // ══════════════════════════════════════════════════════════════════
        //  CLEAR LOGS
        // ══════════════════════════════════════════════════════════════════
        public void ClearLogs()
        {
            lock (_pendingLock)
            {
                _pendingRawLines.Clear();
                _pendingPayloadLines.Clear();
                _pendingRxCompactLines.Clear();
                _pendingRxLiveLatest = null;
            }

            _rawUiLines.Clear();
            _payloadUiLines.Clear();
            _rxCompactUiLines.Clear();

            RawLog = string.Empty;
            PayloadLog = string.Empty;
            RxLiveLog = string.Empty;
            RxCompactLog = string.Empty;
            LatestFields = null;
        }

        // ══════════════════════════════════════════════════════════════════
        //  FILE LOGGER MANAGEMENT
        // ══════════════════════════════════════════════════════════════════
        private bool OpenLoggers()
        {
            CloseLoggers();
            try
            {
                _rawLogger = new FileLogger(RawLogFilePath);
                _payloadLogger = new FileLogger(PayloadLogFilePath);
                _rxLiveLogger = new FileLogger(RxLiveLogFilePath);
                return true;
            }
            catch (Exception ex)
            {
                StatusText = "Logger error: " + ex.Message;
                _rawLogger = null;
                _payloadLogger = null;
                _rxLiveLogger = null;
                return false;
            }
        }

        private void CloseLoggers()
        {
            SafeDispose(ref _rawLogger);
            SafeDispose(ref _payloadLogger);
            SafeDispose(ref _rxLiveLogger);
        }

        private static void SafeDispose(ref FileLogger logger)
        {
            if (logger == null) return;
            try
            {
                logger.Flush();
                logger.Dispose();
            }
            catch { }
            finally
            {
                logger = null;
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  FRAME RECEIVED (ThreadPool thread — NO UI ACCESS)
        // ══════════════════════════════════════════════════════════════════
        private void OnFrameReceived(object sender, FrameReceivedEventArgs e)
        {
            string rawLine = FrameValidator.BuildFrameLogLine(
                e.RawFrame, e.PortNumber, e.FrameCount);

            if (e.IsValid)
            {
                // ── File I/O on background thread ──
                if (_rawLogger != null)
                    _rawLogger.Log(rawLine);

                string payloadLine = FrameValidator.BuildPayloadLogLine(
                    e.Payload, e.PortNumber, e.FrameCount);

                if (_payloadLogger != null)
                    _payloadLogger.Log(payloadLine);

                // ── Parse frame — all string work on background thread ──
                long frameNum = Interlocked.Increment(ref _rxFrameNumber);
                RxDataFields fields = RxDataParser.ParseFrame(e.RawFrame);

                // Detailed display for RX LIVE STATS tab
                string detailedBlock = RxDataParser.GenerateDetailedDisplay(
                    fields, (int)frameNum);

                // Compact summary for RX LIVE COMPARE tab
                string compactLine = RxDataParser.FormatCompactSummary(
                    fields, (int)frameNum);

                // Log detailed to file
                if (_rxLiveLogger != null)
                    _rxLiveLogger.Log(detailedBlock);

                // ── Queue for batch UI update ──
                lock (_pendingLock)
                {
                    _pendingRawLines.Add(rawLine);
                    _pendingPayloadLines.Add(payloadLine);
                    _pendingRxCompactLines.Add(compactLine);
                    _pendingRxLiveLatest = detailedBlock; // overwrite — keep latest only
                }

                // ── Update LatestFields on UI thread (small object) ──
                RxDataFields captured = fields;
                _dispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    new Action(() =>
                    {
                        LatestFields = captured;
                        RefreshStats();
                    }));
            }
            else
            {
                // Invalid frame
                string badLine =
                    rawLine + "  << INVALID: " + e.FailReason + " >>";

                if (_rawLogger != null)
                    _rawLogger.Log(badLine);

                lock (_pendingLock)
                {
                    _pendingRawLines.Add(badLine);
                }

                _dispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    new Action(RefreshStats));
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  BATCH UI UPDATE TIMER (100ms, Background priority)
        //  Prevents COM context timeout by minimizing UI thread work
        // ══════════════════════════════════════════════════════════════════
        private void OnUiUpdateTimerTick(object sender, EventArgs e)
        {
            // STEP 1: Snapshot pending data under lock (fast)
            List<string> rawBatch = null;
            List<string> payloadBatch = null;
            List<string> rxCompactBatch = null;
            string rxLiveLatest = null;

            lock (_pendingLock)
            {
                if (_pendingRawLines.Count > 0)
                {
                    rawBatch = new List<string>(_pendingRawLines);
                    _pendingRawLines.Clear();
                }
                if (_pendingPayloadLines.Count > 0)
                {
                    payloadBatch = new List<string>(_pendingPayloadLines);
                    _pendingPayloadLines.Clear();
                }
                if (_pendingRxCompactLines.Count > 0)
                {
                    rxCompactBatch = new List<string>(_pendingRxCompactLines);
                    _pendingRxCompactLines.Clear();
                }
                if (_pendingRxLiveLatest != null)
                {
                    rxLiveLatest = _pendingRxLiveLatest;
                    _pendingRxLiveLatest = null;
                }
            }

            // STEP 2: Merge into ring buffers (fast list operations)
            bool anyChange = false;

            if (rawBatch != null && rawBatch.Count > 0)
            {
                MergeIntoRingBuffer(_rawUiLines, rawBatch, MAX_UI_LINES);
                anyChange = true;
            }

            if (payloadBatch != null && payloadBatch.Count > 0)
            {
                MergeIntoRingBuffer(_payloadUiLines, payloadBatch, MAX_UI_LINES);
                anyChange = true;
            }

            if (rxCompactBatch != null && rxCompactBatch.Count > 0)
            {
                MergeIntoRingBuffer(_rxCompactUiLines, rxCompactBatch, MAX_COMPACT_LINES);
                anyChange = true;
            }

            // STEP 3: Rebuild strings ONLY if changed
            if (anyChange)
            {
                // string.Join is still O(n) but only when data actually changed
                RawLog = string.Join(Environment.NewLine, _rawUiLines);
                PayloadLog = string.Join(Environment.NewLine, _payloadUiLines);
                RxCompactLog = string.Join(Environment.NewLine, _rxCompactUiLines);

                RefreshStats();
            }

            // STEP 4: Update latest RxLive (single string — no join)
            if (rxLiveLatest != null)
                RxLiveLog = rxLiveLatest;

            // STEP 5: Adaptive timer — slow down when idle
            if (!anyChange && rxLiveLatest == null)
            {
                // No activity — reduce frequency to save CPU
                if (_uiUpdateTimer.Interval.TotalMilliseconds < 500)
                    _uiUpdateTimer.Interval = TimeSpan.FromMilliseconds(500);
            }
            else
            {
                // Data flowing — keep responsive
                if (_uiUpdateTimer.Interval.TotalMilliseconds > 100)
                    _uiUpdateTimer.Interval = TimeSpan.FromMilliseconds(100);
            }
        }

        // ── Helper: Merge new lines into ring buffer ──────────────────────
        private static void MergeIntoRingBuffer(List<string> ringBuffer,
                                                 List<string> newItems,
                                                 int maxItems)
        {
            // Add all new items
            ringBuffer.AddRange(newItems);

            // Trim excess with single RemoveRange (fast)
            if (ringBuffer.Count > maxItems)
            {
                int excess = ringBuffer.Count - maxItems;
                ringBuffer.RemoveRange(0, excess);
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  SERIALPORT MANAGER CALLBACKS
        // ══════════════════════════════════════════════════════════════════

        private void OnStatusChanged(object sender, StringEventArgs e)
        {
            string msg = e.Message;
            _dispatcher.BeginInvoke(
                DispatcherPriority.Normal,
                new Action(() => StatusText = msg));
        }

        private void OnErrorOccurred(object sender, StringEventArgs e)
        {
            string msg = "[ERROR] " + e.Message;
            _dispatcher.BeginInvoke(
                DispatcherPriority.Normal,
                new Action(() =>
                {
                    StatusText = msg;
                    lock (_pendingLock)
                    {
                        _pendingRawLines.Add(msg);
                    }
                }));
        }

        // ══════════════════════════════════════════════════════════════════
        //  HELPER: REFRESH STATS COUNTERS
        // ══════════════════════════════════════════════════════════════════
        private void RefreshStats()
        {
            OnPropertyChanged("TotalFrames");
            OnPropertyChanged("ValidFrames");
            OnPropertyChanged("InvalidFrames");
        }

        // ══════════════════════════════════════════════════════════════════
        //  EVENT RAISERS
        // ══════════════════════════════════════════════════════════════════
        private void RaisePortStatusMessage(string msg)
        {
            EventHandler<StringEventArgs> h = PortStatusMessage;
            if (h != null)
                h(this, new StringEventArgs(msg));
        }

        private void RaiseConnectionStateChanged()
        {
            EventHandler h = ConnectionStateChanged;
            if (h != null)
                h(this, EventArgs.Empty);
        }

        // ══════════════════════════════════════════════════════════════════
        //  INotifyPropertyChanged
        // ══════════════════════════════════════════════════════════════════
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler h = PropertyChanged;
            if (h != null)
                h(this, new PropertyChangedEventArgs(name));
        }

        // ══════════════════════════════════════════════════════════════════
        //  IDisposable
        // ══════════════════════════════════════════════════════════════════
        public void Dispose()
        {
            _uiUpdateTimer.Stop();
            Disconnect();
        }
    }
}