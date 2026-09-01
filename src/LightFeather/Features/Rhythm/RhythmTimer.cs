using System;
using System.Diagnostics;
using Timer = System.Windows.Forms.Timer;

namespace LightFeather.Features.Rhythm {
	public class RhythmTimer {
		private static Timer _timer;

		public RhythmTimer() {
			if (_timer is null) {
				_timer = new Timer {
					Interval = 200
				};
			}
		}

		public void Start(EventHandler onTick) {
			if (_timer.Enabled) {
				Debug.WriteLine("Timer already started. Skipping.", "[Timer]");
				return;
			}

			Debug.WriteLine("Timer started.", "[Timer]");
			_timer.Tick -= onTick;
			_timer.Start();
			_timer.Tick += onTick;
		}

		public void Stop() {
			if (_timer == null) {
				Debug.WriteLine("Timer was null.", "[Timer]");
				return;
			}

			Debug.WriteLine("Timer stopped.", "[Timer]");
			_timer.Stop();
		}
	}
}