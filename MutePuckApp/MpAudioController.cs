using System;
using System.Collections.Generic;
using System.Linq;
using CoreAudioApi;

namespace MutePuckApp
{
    public delegate void VolumeNotificationEvent(AudioVolumeNotificationData data);
    

    /// <summary>
    /// Starting point: http://msdn.microsoft.com/en-us/library/dd370805(v=vs.85).aspx
    /// </summary>
    public class mpAudioController: IDisposable
    {
        private readonly MMDeviceCollection _devices;
        private readonly int _count;
        public event VolumeNotificationEvent OnVolumeNotification = delegate { };
        //public DeviceEvents Hook = new DeviceEvents();

        public mpAudioController()
        {
            MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
            _devices = enumerator.EnumerateAudioEndPoints(EDataFlow.eCapture,
                EDeviceState.DEVICE_STATE_ACTIVE);
            _count = _devices.Count;

            for (int i = 0; i < _count; i++)
            {
                _devices[i].AudioEndpointVolume.OnVolumeNotification += AudioEndpointVolume_OnVolumeNotification;
            }
        }

        void AudioEndpointVolume_OnVolumeNotification(AudioVolumeNotificationData data)
        {
            OnVolumeNotification(data);
        }

        public List<MMDevice> GetDevices()
        {
            List<MMDevice> _mDev;
            MMDeviceCollection _devices;
            _mDev = new List<MMDevice>();
            MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
            _devices = enumerator.EnumerateAudioEndPoints(EDataFlow.eCapture,
                EDeviceState.DEVICE_STATE_ACTIVE);
            int _count = _devices.Count;
            for (int i = 0; i < _count; i++)
            {
                _mDev.Add(_devices[i]);
            }
            return _mDev;
        }




        public void SetMicStateTo(MpMicStates state)
        {
            // Removed this check so the current state remains the responsibility of the caller (i.e. NotifyIconViewModel).
            // This class now just does whatever it is told to do.
            //if (_oldState == state) return;

            for (int i = 0; i < _count; i++)
            {
                try
                {
                    _devices[i].AudioEndpointVolume.Mute = state == MpMicStates.Muted;
                }
                catch
                {
                    //We don't care about it beeing set or not.
                    //Sometimes, it doesn't work.
                }
            }
        }

        

		private bool disposedValue;

		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing && _devices != null)
				{
					for(int i = 0; i < _count; i++)
					{
                        _devices[i].AudioEndpointVolume.OnVolumeNotification -= AudioEndpointVolume_OnVolumeNotification;
                    }
				}
				disposedValue = true;
			}
		}

		public void Dispose()
		{
			// Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
			Dispose(disposing: true);
			GC.SuppressFinalize(this);
		}
	}
}