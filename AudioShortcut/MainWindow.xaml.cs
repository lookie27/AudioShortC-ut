using IWshRuntimeLibrary;
using Microsoft.Toolkit.Uwp.Notifications;
using NAudio.CoreAudioApi;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Windows.UI.Notifications;
using static AudioShortcut.DisplayAudioDevice;

namespace AudioShortcut
{
    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<DisplayAudioDevice> audioDevices = new ObservableCollection<DisplayAudioDevice>();

        private static class ArgumentCommands
        {
            public const string CREATE = "CREATE";
            public const string SWITCH = "SWITCH";
        }
        public MainWindow()
        {
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length > 2)
            {
                switch (args[1])
                {
                    case ArgumentCommands.CREATE:
                        {
                            createShortcut(args[2], args[3]);
                            break;
                        }
                    case ArgumentCommands.SWITCH:
                        {
                            switchAudioDevice(args[2]);
                            System.Environment.Exit(0);
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }
            }
            InitializeComponent();
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
            shownNotification(deviceName);
            Process.Start(path, $"setdefaultsounddevice \"{deviceName}\" 1");
        }

        private void elevatePermissions(String deviceName, String iconName)
        {
            if (!IsElevated())
            {
                var path = Assembly.GetExecutingAssembly().Location;
                using (var process = Process.Start(new ProcessStartInfo(path, $"{ArgumentCommands.CREATE} {deviceName} {iconName} /run_elevated_action")
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

        private void createShortcut(String deviceName, String iconLocation)
        {
            char[] charactersToRemove = { '<', '>', ':', '"', '/', '\\', '|', '?', '*' };
            String trimmedDeviceName = deviceName.Trim(charactersToRemove);
            elevatePermissions(trimmedDeviceName, iconLocation);
            var startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var shell = new WshShell();
            var shortCutLinkFilePath = System.IO.Path.Combine(startupFolderPath, $"{trimmedDeviceName}.lnk");
            var windowsApplicationShortcut = (IWshShortcut)shell.CreateShortcut(shortCutLinkFilePath);
            windowsApplicationShortcut.IconLocation = $"{iconLocation}";
            windowsApplicationShortcut.Description = "AudioShortcut";
            windowsApplicationShortcut.Arguments = $"{ArgumentCommands.SWITCH} \"{deviceName}\"";
            windowsApplicationShortcut.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            windowsApplicationShortcut.TargetPath = System.Windows.Forms.Application.ExecutablePath;
            windowsApplicationShortcut.Save();
        }

        private void button_MouseEnter(object sender, MouseEventArgs e)
        {
            ((Button)sender).Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2c2f33"));
        }

        private void button_MouseLeave(object sender, MouseEventArgs e)
        {
            ((Button)sender).Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ffffff"));
        }

        private void createShortcutButton_Click(object sender, RoutedEventArgs e)
        {
            if (audioDeviceListView.SelectedIndex != -1)
            {
                DisplayAudioDevice selectedItem = (DisplayAudioDevice)audioDeviceListView.SelectedItem;
                createShortcut(selectedItem.FriendlyName, selectedItem.IconLocation);
            }
        }

        private void audioDeviceListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (audioDeviceListView.SelectedIndex != -1)
            {
                DisplayAudioDevice selectedDevice = (DisplayAudioDevice)audioDeviceListView.SelectedItem;
                MMType speakerType = selectedDevice.SpeakerType;
                MMType newSpeakerType = speakerType == MMType.SPEAKER ? MMType.HEADPHONE : MMType.SPEAKER;
                selectedDevice.SpeakerType = newSpeakerType;
                audioDeviceListView.Items.Refresh();
            }
        }

        private void shownNotification(String deviceName)
        {
            DesktopNotificationManagerCompat.RegisterAumidAndComServer<MyNotificationActivator>("LucasBuccilli.AudioShortcut");
            DesktopNotificationManagerCompat.RegisterActivator<MyNotificationActivator>();

            ToastContent toastContent = new ToastContentBuilder()
                .AddToastActivationInfo("action=viewConversation&conversationId=5", ToastActivationType.Foreground)
                .AddText($"Audio device set to {deviceName}")
                .GetToastContent();

            var toast = new ToastNotification(toastContent.GetXml());
            DesktopNotificationManagerCompat.CreateToastNotifier().Show(toast);
        }

        private void switchButton_Click(object sender, RoutedEventArgs e)
        {
            if (audioDeviceListView.SelectedIndex != -1)
            {
                switchAudioDevice(((DisplayAudioDevice)audioDeviceListView.SelectedItem).FriendlyName);
            }
        }
    }

    public class DisplayAudioDevice
    {
        public enum MMType
        {
            SPEAKER,
            HEADPHONE
        }
        public MMDevice device { get; }
        public String DeviceFriendlyName { get { return this.device.DeviceFriendlyName; } }
        public String DisplayIcon { get { return SpeakerType == MMType.SPEAKER ? "Resources/speakerIcon.png" : "Resources/headphoneIcon.png"; } }
        public String IconLocation { get { return SpeakerType == MMType.SPEAKER ? "%systemroot%\\system32\\ddores.dll,88" : "%systemroot%\\system32\\ddores.dll,89"; } }
        public MMType SpeakerType { get; set; }
        public String FriendlyName { get { return _FriendlyName; } }
        private String _FriendlyName;

        public DisplayAudioDevice(MMDevice device)
        {
            this.device = device;
            this.SpeakerType = MMType.SPEAKER;
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

    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(INotificationActivationCallback))]
    [Guid("45824dff-aee5-4ceb-a6f3-ec52c67b64ea"), ComVisible(true)]
    public class MyNotificationActivator : NotificationActivator
    {
        public override void OnActivated(string invokedArgs, NotificationUserInput userInput, string appUserModelId)
        {
        }
    }

}
