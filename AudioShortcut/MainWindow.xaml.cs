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

        private void switchAudioDevice(String deviceId)
        {
            Console.WriteLine(deviceId);
            //CoreAudioController is slow to new up
            CoreAudioController controller = new CoreAudioController();
            IEnumerable<CoreAudioDevice> audioDevices = controller.GetPlaybackDevices();
            foreach(CoreAudioDevice device in audioDevices)
            {
                Console.WriteLine(device.RealId);
                if (device.RealId == deviceId)
                {
                    controller.SetDefaultDevice(device);
                    break;
                }
            }

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
            switchAudioDevice(((DisplayAudioDevice)audioDeviceListView.SelectedItem).device.ID);
        }
    }

    public class DisplayAudioDevice
    {
        public MMDevice device { get; }
        public String DeviceFriendlyName { get { return this.device.DeviceFriendlyName;} }
        public String FriendlyName { get { return removeDeviceFriendlyName(this.device.FriendlyName);} }

        public DisplayAudioDevice(MMDevice device)
        {
            this.device = device;
        }

        private String removeDeviceFriendlyName(String friendlyName)
        {
            String value = friendlyName;
            while (value[value.Length - 1] != '(')
            {
                value = value.Remove(value.Length - 1, 1);
            }
            return value.Remove(value.Length - 1, 1); ;
        }
    }
}
