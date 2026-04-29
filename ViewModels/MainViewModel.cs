using System;
using System.ComponentModel;
using System.Windows.Threading;
using ANVESHA_TCRX_HEALTH_STATUS_GUI_V2.Models;

namespace ANVESHA_TCRX_HEALTH_STATUS_GUI_V2.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged, IDisposable
    {
        // ── Uptime timer ───────────────────────────────────────────────────
        private readonly DispatcherTimer _uiTimer;
        private readonly DateTime _startTime;

        // ── Port ViewModels ────────────────────────────────────────────────
        public PortViewModel Port1 { get; private set; }
        public PortViewModel Port2 { get; private set; }

        // ── Backing fields ─────────────────────────────────────────────────
        private string _uptimeText = "00:00:00";
        private string _appStatus = "Ready — Configure ports then click CONNECT BOTH";
        private bool _isAcquiring = false;
        private string _acquisitionStatus = "IDLE";

        // ── Properties ────────────────────────────────────────────────────
        public string UptimeText
        {
            get { return _uptimeText; }
            private set
            {
                _uptimeText = value;
                OnPropertyChanged("UptimeText");
            }
        }

        public string AppStatus
        {
            get { return _appStatus; }
            set
            {
                _appStatus = value;
                OnPropertyChanged("AppStatus");
            }
        }

        public bool IsAcquiring
        {
            get { return _isAcquiring; }
            set
            {
                if (_isAcquiring == value) return;
                _isAcquiring = value;
                OnPropertyChanged("IsAcquiring");
                OnPropertyChanged("IsNotAcquiring");
            }
        }

        public bool IsNotAcquiring
        {
            get { return !_isAcquiring; }
        }

        public string AcquisitionStatus
        {
            get { return _acquisitionStatus; }
            set
            {
                _acquisitionStatus = value;
                OnPropertyChanged("AcquisitionStatus");
            }
        }

        // ── Constructor ────────────────────────────────────────────────────
        public MainViewModel(Dispatcher dispatcher)
        {
            Port1 = new PortViewModel(1, dispatcher);
            Port2 = new PortViewModel(2, dispatcher);

            _startTime = DateTime.Now;

            // Subscribe port events
            Port1.PortStatusMessage += OnPortStatusMessage;
            Port2.PortStatusMessage += OnPortStatusMessage;
            Port1.ConnectionStateChanged += OnPortConnectionChanged;
            Port2.ConnectionStateChanged += OnPortConnectionChanged;

            // Uptime ticker
            _uiTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _uiTimer.Tick += (s, e) =>
            {
                TimeSpan up = DateTime.Now - _startTime;
                UptimeText = up.ToString(@"hh\:mm\:ss");
            };
            _uiTimer.Start();
        }

        // ── Port event handlers ────────────────────────────────────────────
        private void OnPortConnectionChanged(object sender, EventArgs e)
        {
            UpdateAcquisitionState();
        }

        private void OnPortStatusMessage(object sender, StringEventArgs e)
        {
            AppStatus = e.Message;
        }

        // ── Centralized acquisition state ──────────────────────────────────
        public void UpdateAcquisitionState()
        {
            bool p1 = Port1.IsConnected;
            bool p2 = Port2.IsConnected;

            if (p1 && p2)
            {
                IsAcquiring = true;
                AcquisitionStatus = "ACQUIRING";
                AppStatus = string.Format(
                    "BOTH PORTS ACTIVE  |  PORT1=[{0}]@{1}  PORT2=[{2}]@{3}",
                    Port1.SelectedPort, Port1.SelectedBaudRate,
                    Port2.SelectedPort, Port2.SelectedBaudRate);
            }
            else if (p1 || p2)
            {
                IsAcquiring = true;
                AcquisitionStatus = "PARTIAL";
                AppStatus = p1
                    ? string.Format("PORT 1 ACTIVE [{0}]  —  PORT 2 IDLE",
                                    Port1.SelectedPort)
                    : string.Format("PORT 2 ACTIVE [{0}]  —  PORT 1 IDLE",
                                    Port2.SelectedPort);
            }
            else
            {
                IsAcquiring = false;
                AcquisitionStatus = "IDLE";
                AppStatus = "All ports disconnected.";
            }
        }

        // ── INotifyPropertyChanged ─────────────────────────────────────────
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler h = PropertyChanged;
            if (h != null)
                h(this, new PropertyChangedEventArgs(name));
        }

        // ── IDisposable ────────────────────────────────────────────────────
        public void Dispose()
        {
            _uiTimer.Stop();

            Port1.ConnectionStateChanged -= OnPortConnectionChanged;
            Port2.ConnectionStateChanged -= OnPortConnectionChanged;
            Port1.PortStatusMessage -= OnPortStatusMessage;
            Port2.PortStatusMessage -= OnPortStatusMessage;

            Port1.Dispose();
            Port2.Dispose();
        }
    }
}