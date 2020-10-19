using AudioSwitcher.AudioApi.CoreAudio;
using NAudio.CoreAudioApi;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Resources;
using IWshRuntimeLibrary;
using System.Security.Principal;

namespace AudioShortcut
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<DisplayAudioDevice> audioDevices = new ObservableCollection<DisplayAudioDevice>();
        public MainWindow()
        {
            InitializeComponent();

            //get args here

            initializeActiveAudioDevices();
            audioDeviceListView.ItemsSource = audioDevices;

        }

       private void initializeActiveAudioDevices()
        {
            MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
            MMDeviceCollection activeAudioDevices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
            foreach (MMDevice device in activeAudioDevices)
            {
                audioDevices.Add(new DisplayAudioDevice(device));
            }
        }

        private void switchAudioDevice(String deviceName)
        {
            string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "nircmd.exe");
            System.IO.File.WriteAllBytes(path, AudioShortcut.Properties.Resources.nircmd);
            Process.Start(path, $"setdefaultsounddevice \"{deviceName}\" 1");
        }

        private void elevatePermissions()
        {
            if (!IsElevated())
            {
                var path = Assembly.GetExecutingAssembly().Location;
                using (var process = Process.Start(new ProcessStartInfo(path, "/run_elevated_action")
                {
                    Verb = "runas"
                }))
                {
                    System.Environment.Exit(0);
                }
            }
        }

        private static bool IsElevated()
        {
            using (var identity = WindowsIdentity.GetCurrent())
            {
                var principal = new WindowsPrincipal(identity);

                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private void createShortcut(String deviceName)
        {
            elevatePermissions();
            var startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var shell = new WshShell();
            var shortCutLinkFilePath = System.IO.Path.Combine(startupFolderPath, "CreateShortcutSample.lnk");
            var windowsApplicationShortcut = (IWshShortcut)shell.CreateShortcut(shortCutLinkFilePath);
            windowsApplicationShortcut.Description = "AudioShortcut";
            windowsApplicationShortcut.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            windowsApplicationShortcut.TargetPath = System.Windows.Forms.Application.ExecutablePath;
            windowsApplicationShortcut.Save();
        }

        private void button_MouseEnter(object sender, MouseEventArgs e)
        {
            ((Button) sender).Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2c2f33"));
        }

        private void button_MouseLeave(object sender, MouseEventArgs e)
        {
            ((Button)sender).Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffffff"));
        }

        private void testShortcutButton_Click(object sender, RoutedEventArgs e)
        {
            switchAudioDevice(((DisplayAudioDevice)audioDeviceListView.SelectedItem).FriendlyName);
        }

        private void createShortcutButton_Click(object sender, RoutedEventArgs e)
        {
            createShortcut("test");
        }
    }

    public class DisplayAudioDevice
    {
        public MMDevice device { get; }
        public String DeviceFriendlyName { get { return this.device.DeviceFriendlyName;} }
        public String FriendlyName { get { return _FriendlyName; } }
        private String _FriendlyName;

        public DisplayAudioDevice(MMDevice device)
        {
            this.device = device;
            this._FriendlyName = removeDeviceFriendlyName(this.device.FriendlyName);
        }

        private String removeDeviceFriendlyName(String friendlyName)
        {
            String value = friendlyName;
            while (value[value.Length - 1] != '(')
            {
                value = value.Remove(value.Length - 1, 1);
            }

            value = value.Remove(value.Length - 1, 1);

            if (value[value.Length - 1] == ' ')
            {
                value = value.Remove(value.Length - 1, 1);
            }

            return value;
        }
    }
}
