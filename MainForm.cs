using AudioSwitcher.AudioApi;
using AudioSwitcher.AudioApi.CoreAudio;
using Microsoft.Win32;
using Shortcut;
using System;
using System.Drawing;
using System.Reactive;
using System.Windows.Forms;
using Colore;
using Colore.Effects.Keyboard;
using Colore.Effects.Headset;

namespace MicMute
{
    public partial class MainForm : Form
    {
        public CoreAudioController AudioController = new CoreAudioController();
        private readonly HotkeyBinder hotkeyBinder = new HotkeyBinder();
        private readonly RegistryKey registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\MicMute");
        private readonly string registryKeyName = "Hotkey";

        private Hotkey hotkey;
        private IChroma chroma;
        private KeyboardCustom muted;
        private bool myVisible;
        public bool MyVisible
        {
            get { return myVisible; }
            set { myVisible = value; Visible = value; }
        }

        public MainForm()
        {
            InitializeComponent();

            initChromaSkd();

        }

        private void initChromaSkd()
        {
            var t = ColoreProvider.CreateNativeAsync();
            t.Wait();
            this.chroma = t.Result;

            this.muted = KeyboardCustom.Create();
            this.muted.Set(Colore.Data.Color.Red);
            this.muted[Key.Macro4] = Colore.Data.Color.Green;

            this.unmuted = KeyboardCustom.Create();
            this.unmuted.Set(Colore.Data.Color.Green);
            this.unmuted[Key.Macro4] = Colore.Data.Color.Red;


        }

        private void OnNextDevice(DeviceChangedArgs next)
        {
            UpdateDevice(AudioController.DefaultCaptureDevice);
        }

        private void MyHide()
        {
            ShowInTaskbar = false;
            Location = new Point(-10000, -10000);
            MyVisible = false;
        }

        private void MyShow()
        {
            MyVisible = true;
            ShowInTaskbar = true;
            CenterToScreen();
        }


        private void MainForm_Load(object sender, EventArgs e)
        {
            MyHide();
            UpdateDevice(AudioController.DefaultCaptureDevice);
            AudioController.AudioDeviceChanged.Subscribe(OnNextDevice);

            var timer = new Timer() { Interval = 500, Enabled = true };

            timer.Tick += Timer_Tick;

            timer.Start();


            var hotkeyValue = registryKey.GetValue(registryKeyName);
            if (hotkeyValue != null)
            {
                var converter = new Shortcut.Forms.HotkeyConverter();
                hotkey = (Hotkey)converter.ConvertFromString(hotkeyValue.ToString());
                hotkeyBinder.Bind(hotkey).To(ToggleMicStatus);
            }

            //AudioController.AudioDeviceChanged.Subscribe(x =>
            //{
            //    Debug.WriteLine("{ 0} - {1}", x.Device.Name, x.ChangedType.ToString());
            //});
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (!isDeviceInUse())
            {
                if (chroma.Initialized)
                    chroma.UninitializeAsync();
            }
            else
            {
                if (!chroma.Initialized)
                {
                    initChromaSkd();
                    updateKeyboardStatus(AudioController.DefaultCaptureDevice);
                }
                else
                    updateKeyboardStatus(AudioController.DefaultCaptureDevice);
            }
        }

        private static bool isDeviceInUse()
        {
            var subkeys = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone");
            foreach (var sub in subkeys.GetSubKeyNames())
            {
                var subsub = subkeys.OpenSubKey(sub);
                if (subsub.SubKeyCount == 0)
                {
                    // LastUsedTimeStart
                    // LastUsedTimeStop
                    var lastStart = subsub.GetValue("LastUsedTimeStart");
                    var lastStop = subsub.GetValue("LastUsedTimeStop");
                    var startDate = new DateTime((long)lastStart);
                    var stopDate = new DateTime((long)lastStop);
                    if (startDate > stopDate)
                        return true;

                }
                else
                {
                    foreach (var key in subsub.GetSubKeyNames())
                    {
                        var np = subsub.OpenSubKey(key);
                        var lastStart = np.GetValue("LastUsedTimeStart");
                        var lastStop = np.GetValue("LastUsedTimeStop");
                        var startDate = new DateTime((long)lastStart);
                        var stopDate = new DateTime((long)lastStop);
                        if (startDate > stopDate)
                        {
                            //Console.WriteLine(key);
                            //Console.WriteLine(startDate);
                            //Console.WriteLine(stopDate);
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private void OnMuteChanged(DeviceMuteChangedArgs next)
        {
            UpdateStatus(next.Device);
        }

        IDisposable muteChangedSubscription;
        public void UpdateDevice(AudioSwitcher.AudioApi.IDevice device)
        {
            muteChangedSubscription?.Dispose();
            muteChangedSubscription = device?.MuteChanged.Subscribe(OnMuteChanged);
            UpdateStatus(device);
        }

        Icon iconOff = Properties.Resources.off;
        Icon iconOn = Properties.Resources.on;
        Icon iconError = Properties.Resources.error;
        private KeyboardCustom unmuted;
        private bool lastMutedState;

        public void UpdateStatus(AudioSwitcher.AudioApi.IDevice device)
        {
            if (device != null)
            {
                UpdateIcon(device.IsMuted ? iconOff : iconOn, device.FullName);
                updateKeyboardStatus(device);

            }
            else
            {
                UpdateIcon(iconError, "< No device >");
            }
        }

        private void updateKeyboardStatus(AudioSwitcher.AudioApi.IDevice device)
        {
            if (chroma.Initialized)
            {
                chroma.Keyboard.SetCustomAsync(device.IsMuted ? this.muted : this.unmuted);
            }
        }

        private void UpdateIcon(Icon icon, string tooltipText)
        {
            this.icon.Icon = icon;
            this.icon.Text = tooltipText;
        }

        public async void ToggleMicStatus()
        {
            await AudioController.DefaultCaptureDevice?.ToggleMuteAsync();
        }

        public void UpdateStatus()
        {
            var device = AudioController.DefaultCaptureDevice;

            if (device != null)
            {
                UpdateIcon(device.IsMuted ? Properties.Resources.off : Properties.Resources.on, device.FullName);
            }
            else
            {
                UpdateIcon(Properties.Resources.error, "< No device >");
            }
        }

        private void Icon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ToggleMicStatus();
            }
        }

        private void HotkeyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (hotkey != null)
            {
                hotkeyTextBox.Hotkey = hotkey;
                hotkeyBinder.Unbind(hotkey);
            }
            MyShow();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MyVisible)
            {
                MyHide();
                e.Cancel = true;

                hotkey = hotkeyTextBox.Hotkey;

                if (hotkey == null)
                {
                    registryKey.DeleteValue(registryKeyName, false);
                }
                else
                {
                    if (!hotkeyBinder.IsHotkeyAlreadyBound(hotkey))
                    {
                        registryKey.SetValue(registryKeyName, hotkey);
                        hotkeyBinder.Bind(hotkey).To(ToggleMicStatus);
                    }
                }
            }
        }

        private void ButtonReset_Click(object sender, EventArgs e)
        {
            hotkeyTextBox.Hotkey = null;
            hotkeyTextBox.Text = "None";
        }
        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
