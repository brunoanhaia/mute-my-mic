using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace mute_it
{
    class MuteItContext : ApplicationContext
    {
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public static readonly int MUTE_CODE = 123;
        public static readonly int UNMUTE_CODE = 234;

        private static readonly double REFRESH_INTERVAL = 1000;
        private readonly NotifyIcon tbIcon;
        private readonly System.Timers.Timer refreshTimer;
        private readonly HotkeyManager hkManager;

        public MuteItContext()
        {
            tbIcon = CreateIcon();
            UpdateMicStatus();
            refreshTimer = StartTimer();
            hkManager = RegisterHotkeys();
        }

        private HotkeyManager RegisterHotkeys()
        {
            var manager = new HotkeyManager(this);

            RegisterHotKey(manager.Handle, MUTE_CODE, Constants.ALT + Constants.SHIFT, (int)Keys.P);
            RegisterHotKey(manager.Handle, UNMUTE_CODE, Constants.ALT + Constants.SHIFT, (int)Keys.O);

            return manager;
        }

        private NotifyIcon CreateIcon()
        {
            var exitMenuItem = new MenuItem("Exit", new EventHandler(HandleExit));

            var icon = new NotifyIcon
            {
                Icon = Properties.Resources.mic_off,
                ContextMenu = new ContextMenu(new MenuItem[] { exitMenuItem })
            };
            icon.DoubleClick += (source, e) => ToggleMicStatus();

            icon.Visible = true;

            return icon;
        }

        private System.Timers.Timer StartTimer()
        {
            var timer = new System.Timers.Timer(REFRESH_INTERVAL);
            timer.Elapsed += (source, e) => UpdateMicStatus();
            timer.Start();
            return timer;
        }

        public void MuteMic()
        {
            SetMicMuteStatus(true);
        }

        public void UnmuteMic()
        {
            SetMicMuteStatus(false);
        }

        private void SetMicMuteStatus(bool doMute)
        {
            var device = GetPrimaryMicDevice();

            if (device != null)
            {
                device.AudioEndpointVolume.Mute = doMute;
                UpdateMicStatus(device);
            }
            else
            {
                UpdateMicStatus(null);
            }
        }

        private void ToggleMicStatus()
        {
            var device = GetPrimaryMicDevice();

            if (device != null)
            {
                device.AudioEndpointVolume.Mute = !device.AudioEndpointVolume.Mute;
                UpdateMicStatus(device);
            }
            else
            {
                UpdateMicStatus(null);
            }
        }

        private void UpdateMicStatus()
        {
            var device = GetPrimaryMicDevice();
            UpdateMicStatus(device);
        }

        private void UpdateMicStatus(MMDevice device)
        {
            if (device == null || device.AudioEndpointVolume.Mute)
            {
                tbIcon.Icon = Properties.Resources.mic_off;
            }
            else
            {
                tbIcon.Icon = Properties.Resources.mic_on;
            }

            DisposeDevice(device);
        }

        private MMDevice GetPrimaryMicDevice()
        {
            var enumerator = new MMDeviceEnumerator();
            var result = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);
            enumerator.Dispose();

            tbIcon.Text = result.DeviceFriendlyName;

            return result;
        }

        private void DisposeDevice(MMDevice device)
        {
            if (device != null)
            {
                device.AudioEndpointVolume.Dispose();
                device.Dispose();
            }
        }

        private void HandleExit(object sender, EventArgs e)
        {
            tbIcon.Visible = false;
            refreshTimer.Stop();
            Application.Exit();
        }
    }
}
