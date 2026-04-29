using System;
using System.IO.Ports;
using System.Threading;

namespace ANVESHA_TCRX_HEALTH_STATUS_GUI_V2.Models
{
    // ── Event argument carrying one validated frame ──────────────────────────
    public class FrameReceivedEventArgs : EventArgs
    {
        public byte[] RawFrame { get; private set; }
        public byte[] Payload { get; private set; }
        public int PortNumber { get; private set; }
        public long FrameCount { get; private set; }
        public bool IsValid { get; private set; }
        public string FailReason { get; private set; }

        public FrameReceivedEventArgs(
            byte[] raw, byte[] payload, int portNumber,
            long frameCount, bool isValid, string failReason)
        {
            RawFrame = raw;
            Payload = payload;
            PortNumber = portNumber;
            FrameCount = frameCount;
            IsValid = isValid;
            FailReason = failReason;
        }
    }

    // ── Per-port manager ────────────────────────────────────────────────────
    public class SerialPortManager : IDisposable
    {
        // ── Serial port ────────────────────────────────────────────────────
        private SerialPort _serialPort;
        private readonly object _lock = new object();

        // ── Frame accumulation buffer ──────────────────────────────────────
        private byte[] _accumBuffer = new byte[FrameValidator.FrameSize * 8];
        private int _accumCount = 0;

        // ── Statistics ─────────────────────────────────────────────────────
        private long _totalFrames = 0;
        private long _validFrames = 0;
        private long _invalidFrames = 0;

        // ── Public properties ──────────────────────────────────────────────
        public int PortNumber { get; private set; }
        public string PortName { get; private set; }
        public bool IsConnected { get { return _serialPort != null && _serialPort.IsOpen; } }
        public long TotalFrames { get { return _totalFrames; } }
        public long ValidFrames { get { return _validFrames; } }
        public long InvalidFrames { get { return _invalidFrames; } }

        // ── Events (using custom StringEventArgs — fixes .NET 4.0 error) ──
        public event EventHandler<FrameReceivedEventArgs> FrameReceived;
        public event EventHandler<StringEventArgs> StatusChanged;
        public event EventHandler<StringEventArgs> ErrorOccurred;

        // ── Constructor ────────────────────────────────────────────────────
        public SerialPortManager(int portNumber)
        {
            PortNumber = portNumber;
        }

        // ── Connect ────────────────────────────────────────────────────────
        public bool Connect(string portName, int baudRate,
                            Parity parity, int dataBits, StopBits stopBits)
        {
            if (IsConnected)
                Disconnect();

            try
            {
                PortName = portName;
                _serialPort = new SerialPort(portName, baudRate, parity, dataBits, stopBits)
                {
                    // RS422 optimised settings
                    ReadTimeout = 500,
                    WriteTimeout = 500,
                    ReadBufferSize = 65536,   // Large buffer for 230400 baud
                    WriteBufferSize = 4096,
                    ReceivedBytesThreshold = 26 // Wake on each frame
                };

                _serialPort.DataReceived += OnDataReceived;
                _serialPort.ErrorReceived += OnErrorReceived;
                _serialPort.Open();

                ResetBuffer();

                RaiseStatus(string.Format(
                    "PORT {0}: Connected to {1} @ {2} baud  [RS422]",
                    PortNumber, portName, baudRate));
                return true;
            }
            catch (Exception ex)
            {
                RaiseError(string.Format(
                    "PORT {0}: Connect FAILED — {1}", PortNumber, ex.Message));
                _serialPort = null;
                return false;
            }
        }

        // ── Disconnect ─────────────────────────────────────────────────────
        public void Disconnect()
        {
            try
            {
                if (_serialPort != null)
                {
                    _serialPort.DataReceived -= OnDataReceived;
                    _serialPort.ErrorReceived -= OnErrorReceived;

                    if (_serialPort.IsOpen)
                        _serialPort.Close();

                    _serialPort.Dispose();
                    _serialPort = null;
                }

                ResetBuffer();
                RaiseStatus(string.Format("PORT {0}: Disconnected", PortNumber));
            }
            catch (Exception ex)
            {
                RaiseError(string.Format(
                    "PORT {0}: Disconnect error — {1}", PortNumber, ex.Message));
            }
        }

        // ── Raw data received (fires on SerialPort thread) ─────────────────
        private void OnDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                SerialPort sp = sender as SerialPort;
                if (sp == null || !sp.IsOpen) return;

                int bytesToRead = sp.BytesToRead;
                if (bytesToRead <= 0) return;

                byte[] incoming = new byte[bytesToRead];
                int read = sp.Read(incoming, 0, bytesToRead);
                if (read <= 0) return;

                lock (_lock)
                {
                    // Grow accumulation buffer if necessary
                    if (_accumCount + read > _accumBuffer.Length)
                    {
                        int newSize = Math.Max(_accumBuffer.Length * 2,
                                                  _accumCount + read + 256);
                        byte[] newBuf = new byte[newSize];
                        Array.Copy(_accumBuffer, newBuf, _accumCount);
                        _accumBuffer = newBuf;
                    }

                    Array.Copy(incoming, 0, _accumBuffer, _accumCount, read);
                    _accumCount += read;

                    ProcessAccumulatedBytes();
                }
            }
            catch (Exception ex)
            {
                RaiseError(string.Format(
                    "PORT {0}: DataReceived error — {1}", PortNumber, ex.Message));
            }
        }

        public void ResetCounters()
        {
            Interlocked.Exchange(ref _totalFrames, 0);
            Interlocked.Exchange(ref _validFrames, 0);
            Interlocked.Exchange(ref _invalidFrames, 0);
        }

        // ── Frame extraction ───────────────────────────────────────────────
        private void ProcessAccumulatedBytes()
        {
            int frameSize = FrameValidator.FrameSize;   // 26

            while (true)
            {
                // 1. Find next 0x0D 0x0A header pair
                int headerIdx = FindHeader();
                if (headerIdx < 0)
                {
                    // No header found — keep last byte (partial header possible)
                    if (_accumCount > 1)
                        ShiftBuffer(_accumCount - 1);
                    break;
                }

                // 2. Discard bytes before header
                if (headerIdx > 0)
                    ShiftBuffer(headerIdx);

                // 3. Need a full frame
                if (_accumCount < frameSize) break;

                // 4. Candidate frame
                byte[] candidate = new byte[frameSize];
                Array.Copy(_accumBuffer, 0, candidate, 0, frameSize);

                string reason;
                bool valid = FrameValidator.ValidateFrame(candidate, out reason);

                Interlocked.Increment(ref _totalFrames);

                byte[] payload = null;
                if (valid)
                {
                    Interlocked.Increment(ref _validFrames);
                    payload = FrameValidator.ExtractPayload(candidate);
                    // Consume full frame
                    ShiftBuffer(frameSize);
                }
                else
                {
                    Interlocked.Increment(ref _invalidFrames);
                    // Skip only header byte 1 to re-sync
                    ShiftBuffer(1);
                }

                // 5. Raise on thread-pool (not holding the lock)
                FrameReceivedEventArgs args = new FrameReceivedEventArgs(
                    candidate, payload, PortNumber,
                    _totalFrames, valid, reason);

                ThreadPool.QueueUserWorkItem(RaiseFrameReceivedCallback, args);
            }
        }

        // ── Header search: 0x0D followed by 0x0A ──────────────────────────
        private int FindHeader()
        {
            for (int i = 0; i < _accumCount - 1; i++)
            {
                if (_accumBuffer[i] == 0x0D &&
                    _accumBuffer[i + 1] == 0x0A)
                    return i;
            }
            return -1;
        }

        // ── Shift accumulated buffer left by count bytes ───────────────────
        private void ShiftBuffer(int count)
        {
            if (count <= 0) return;
            if (count >= _accumCount)
            {
                _accumCount = 0;
                return;
            }
            Buffer.BlockCopy(_accumBuffer, count, _accumBuffer, 0, _accumCount - count);
            _accumCount -= count;
        }

        private void ResetBuffer()
        {
            lock (_lock) { _accumCount = 0; }
        }

        // ── Serial error ───────────────────────────────────────────────────
        private void OnErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            RaiseError(string.Format(
                "PORT {0}: Serial HW error — {1}", PortNumber, e.EventType));
        }

        // ── Event helpers ──────────────────────────────────────────────────
        private void RaiseFrameReceivedCallback(object state)
        {
            EventHandler<FrameReceivedEventArgs> h = FrameReceived;
            if (h != null)
                h(this, (FrameReceivedEventArgs)state);
        }

        private void RaiseStatus(string msg)
        {
            EventHandler<StringEventArgs> h = StatusChanged;
            if (h != null) h(this, new StringEventArgs(msg));
        }

        private void RaiseError(string msg)
        {
            EventHandler<StringEventArgs> h = ErrorOccurred;
            if (h != null) h(this, new StringEventArgs(msg));
        }

        // ── IDisposable ────────────────────────────────────────────────────
        public void Dispose()
        {
            Disconnect();
        }
    }
}