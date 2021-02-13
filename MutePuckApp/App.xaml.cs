using System;
using System.Collections.Generic;
using System.Configuration;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using CoreAudioApi;
using UsbLibrary;
using System.Windows.Interop;
using System.Runtime.InteropServices;

namespace MutePuckApp
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {

        private IntPtr m_hNotifyDevNode; // Needed for USB Events   

        private System.Windows.Forms.NotifyIcon _notifyIcon;
        private bool _isExit;
        private mpAudioController _mic;
        private string _iconPath;
        private bool _isMuted;
        private List<MMDevice> _mmDevices;
        //USB
        private HIDDevice m_usb;
        private List<HIDInfoSet> m_connList;
        private DeviceEvents _deviceEvents;
        private bool RedGreenState;
        enum Color
        {
            Red = 0,
            Green = 1,
            Blue = 2,
            White = 3,
            Yellow = 4,
            Turquiose = 5,
            Black = 6
        };

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            MainWindow = new MainWindow();
            MainWindow.Closing += MainWindow_Closing;

            _notifyIcon = new System.Windows.Forms.NotifyIcon();
            _notifyIcon.DoubleClick += (s, args) => ShowMainWindow();
            _notifyIcon.Icon = MutePuckApp.Properties.Resources.puck_gray; 
            _notifyIcon.Visible = true;

            CreateContextMenu();
            _mic = new mpAudioController();
            _deviceEvents = new DeviceEvents();
            
            _mic.OnVolumeNotification += _mic_OnVolumeNotification;
            _deviceEvents.DevicesChanged += _mic_RefreshMicList;
            _deviceEvents.OnDeviceArrived += usb_OnDeviceArrived;
            _deviceEvents.OnDeviceRemoved += usb_OnDeviceRemoved;
            _mmDevices = _mic.GetDevices();
            string test = "";
            foreach (MMDevice m in _mmDevices)
            {

                test += m.FriendlyName +"\n";
            }

            //HID USB Stuff
            m_usb = new HIDDevice();
            //m_cmds = new List<Command>();
            m_connList = new List<HIDInfoSet>();
            m_usb.OnDeviceRemoved += usb_OnDeviceRemoved;
            m_usb.OnDeviceArrived += usb_OnDeviceArrived;
            m_usb.OnDataRecieved += usb_OnDataReceived;

            ScanConnection();
            Connect();

        }


        #region Audio Functions
        private void _mic_OnVolumeNotification(AudioVolumeNotificationData data)
        {
            IsMuted = data.Muted;
        }

        public bool IsMuted
        {
            get
            {
                _notifyIcon.Icon = _isMuted ? MutePuckApp.Properties.Resources.puck_red : MutePuckApp.Properties.Resources.puck_green;
                _mic.SetMicStateTo(_isMuted ? MpMicStates.Muted : MpMicStates.Unmuted);
                return _isMuted;
            }
            set
            {
                _isMuted = value;
                //IconPath = Settings.IsMuted ? @"res\microphone_muted.ico" : @"res\microphone_unmuted.ico";
                _notifyIcon.Icon = _isMuted ? MutePuckApp.Properties.Resources.puck_red : MutePuckApp.Properties.Resources.puck_green;
                _mic.SetMicStateTo(_isMuted ? MpMicStates.Muted : MpMicStates.Unmuted);
                if (value)
                { 
                    SetColor(Color.Red);
                }
                else
                { 
                    SetColor(Color.Green);
                }

                //ToDo: Handle Settings
                //SerializeStatic.Save(typeof(Settings));
                OnPropertyChanged("IsMuted");
            }
        }

        public void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void _mic_RefreshMicList(object sender, EventArgs e)
        {
            _mic.Dispose();
            _mic = new mpAudioController();
            _mic.OnVolumeNotification += _mic_OnVolumeNotification;
            IsMuted = _isMuted; // reassert the muted state on the new set of mics
            

        }


        #endregion

        #region USB Event Handlers
        private void usb_OnDeviceRemoved(object sender, EventArgs e)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new EventHandler(usb_OnDeviceRemoved), new object[] { sender, e });
            }
            else
            {
                if (m_usb.IsConnected)
                {
                    if (HIDDevice.GetInfoSets(m_usb.VendorID, m_usb.ProductID, m_usb.SerialNumberString).Count == 0)
                    {
                        Disconnect();
                    }
                }

                ScanConnection();
            }
        }

        private void usb_OnDeviceArrived(object sender, EventArgs e)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new EventHandler(usb_OnDeviceArrived), new object[] { sender, e });
            }
            else
            {
                System.Threading.Thread.Sleep(600);
                ScanConnection();

                Connect();

            }
        }

        private void usb_OnDataReceived(object sender, HIDReport args)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new DataReceivedEventHandler(usb_OnDataReceived), new object[] { sender, args });
            }
            else
            {
                //ToDO Handle Data
                //MessageBox.Show("Data received: " + args.ToArray().ToString());
                IsMuted = !_isMuted;

                
                //ShowData("RX", 0x00, args);
            }
        }
        #endregion

        #region USB General Methods
        private void ScanConnection()
        {
            
            int vid = int.Parse(MutePuckApp.Properties.Resources.VID, System.Globalization.NumberStyles.HexNumber);
            int pid = int.Parse(MutePuckApp.Properties.Resources.PID, System.Globalization.NumberStyles.HexNumber);
            m_connList = HIDDevice.GetInfoSets(vid, pid);
            
            /* cbConnectionChooser.Items.Clear();
            
            foreach (var hidInfoSet in m_connList)
            {
                if (hidInfoSet.ProductString == "")
                {
                    cbConnectionChooser.Items.Add("Unknown HID device");
                }
                else
                {
                    cbConnectionChooser.Items.Add(hidInfoSet.GetInfo());
                }
            }

            if (cbConnectionChooser.Items.Count > 0)
            {
                cbConnectionChooser.SelectedIndex = 0;

            } */


        }

        private void Disconnect()
        {
            m_usb.Disconnect();
            //lbConnectionText.Text = "Disconnected";

            //pbLED.Image = MutePuckApp.Properties.Resources.ledred;
            _notifyIcon.Icon = MutePuckApp.Properties.Resources.puck_gray;

        }

        private void Connect()
        {
            HIDInfoSet device;

            if (m_connList.Count == 0) return;

            device = m_connList[0];
            if (HIDDevice.GetInfoSets(device.VendorID, device.ProductID, device.SerialNumberString).Count > 0)
            {
                if (m_usb.Connect(device.DevicePath, true))
                {
                    //lbConnectionText.Text = "Connected";

                    //pbLED.Image = Properties.Resources.ledgreen;
                    
                    IsMuted = false;
                    /*m_ToolStripStatusIcon.ToolTipText = "Manufacturer: " + m_usb.ManufacturerString + "\n" +
                                                        "Product Name: " + m_usb.ProductString + "\n" +
                                                        "Serial Number: " + m_usb.SerialNumberString + "\n" +
                                                        "SW Version: " + m_usb.VersionInBCD; */
                }
            }
            else
            {
                ScanConnection();
                MessageBox.Show(String.Format("DEVICE: {0} is used in other program !", device.ProductString));
            }
        }
        #endregion

        private byte[] GetColor(Color color)
        {
            switch (color)
            {
                case Color.Red:
                    return new byte[] {0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF,
                                      0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00, 0x00,
                                      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
                   
                case Color.Green:
                    return new byte[] { 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00,
                                       0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
                case Color.Black:
                    return new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
                    
                default:
                    return new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
                    

            }

        }

        private void SetColor(Color color)
        {
            switch (color)
            {
                case Color.Red:
                    sendMPMessage( new byte[] {0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF,
                                      0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00, 0x00,
                                      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 });
                    break;

                case Color.Green:
                    sendMPMessage(new byte[] { 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00,
                                       0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 });
                    break;
                case Color.Black:
                    sendMPMessage(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 });
                    break;

                default:
                    sendMPMessage(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 });
                    break;


            }

        }

        private void sendMPMessage (byte[] m)
        {
            

            HIDReport hidReport = new HIDReport(0, m);
            m_usb.Write(hidReport);
            
        }


        private void CreateContextMenu()
        {
            _notifyIcon.ContextMenuStrip =
              new System.Windows.Forms.ContextMenuStrip();
            _notifyIcon.ContextMenuStrip.Items.Add("MainWindow...").Click += (s, e) => ShowMainWindow();
            _notifyIcon.ContextMenuStrip.Items.Add("Exit").Click += (s, e) => ExitApplication();
        }

        private void ExitApplication()
        {
            _isExit = true;
            MainWindow.Close();
            _notifyIcon.Dispose();
            _notifyIcon = null;
        }

        private void ShowMainWindow()
        {
            if (MainWindow.IsVisible)
            {
                if (MainWindow.WindowState == WindowState.Minimized)
                {
                    MainWindow.WindowState = WindowState.Normal;
                }
                MainWindow.Activate();
            }
            else
            {
                MainWindow.Show();
            }
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (!_isExit)
            {
                e.Cancel = true;
                MainWindow.Hide(); // A hidden window can be shown again, a closed one not
            }
        }

        
    }
}
