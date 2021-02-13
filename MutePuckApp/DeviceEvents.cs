using System;
using System.Windows.Forms;



namespace MutePuckApp
{
    /// <summary>
    /// Creates a window to receive and dispatch WM_DEVICECHANGE messages to detect when mics are added/removed.
    /// Also reacts to USB Device changes
    /// </summary>
    public sealed class DeviceEvents : IDisposable
    {
        /// <summary>
        /// Represents the window that is used internally to get the messages.
        /// </summary>
        private class Window : NativeWindow, IDisposable
        {
            private static int WM_DEVICECHANGE = 0x0219;
            /// <summary>WParam for above : A device was inserted</summary>
            public const int DEVICE_ARRIVAL = 0x8000;
            /// <summary>WParam for above : A device was removed</summary>
            public const int DEVICE_REMOVECOMPLETE = 0x8004;

            public Window()
            {
                // create the handle for the window.
                this.CreateHandle(new CreateParams());
            }

            /// <summary>
            /// Overridden to get the notifications.
            /// </summary>
            /// <param name="m"></param>
            protected override void WndProc(ref Message m)
            {
                base.WndProc(ref m);

                // check if we got a DeviceChange.
                if (m.Msg == WM_DEVICECHANGE)
				{
                    
                    if (DevicesChanged != null)
					{
                        DevicesChanged(this, null);
					}
                    switch (m.WParam.ToInt32()) // Check the W parameter to see if a device was inserted or removed
                    {
                        case DEVICE_ARRIVAL:    // inserted
                            OnDeviceArrived(this, new EventArgs());
                    
                            break;

                        case DEVICE_REMOVECOMPLETE: // removed
                            OnDeviceRemoved(this, new EventArgs());
                    
                            break;
                    }
                }
                
            }

            
            public event EventHandler DevicesChanged;
            public event EventHandler OnDeviceArrived;
            public event EventHandler OnDeviceRemoved;

            #region IDisposable Members

            public void Dispose()
            {
                this.DestroyHandle();
            }

            #endregion
        }

        private Window _window = new Window();
        
        public DeviceEvents()
        {
            // register the event of the inner native window.
            _window.DevicesChanged += delegate (object sender, EventArgs args)
            {
                if (DevicesChanged != null)
                    DevicesChanged(this, args);
            };
            _window.OnDeviceArrived += delegate (object sender, EventArgs args)
            {
                if (OnDeviceArrived != null)
                    OnDeviceArrived(this, args);
            };
            _window.OnDeviceRemoved += delegate (object sender, EventArgs args)
            {
                if (OnDeviceRemoved != null)
                    OnDeviceRemoved(this, args);
            };


        }

        /// <summary>
        /// A device has been changed.
        /// </summary>
        
        public event EventHandler DevicesChanged;
        public event EventHandler OnDeviceArrived;
        public event EventHandler OnDeviceRemoved;

        #region IDisposable Members

        public void Dispose()
        {
            
            // dispose the inner native window.
            _window.Dispose();
        }

        

        #endregion
    }

   
}
