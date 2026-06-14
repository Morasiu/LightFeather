using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Word;

namespace LightFeather.Features.Rhythm {
	public class RhythmTimer {
		private static Timer _timer;

		public RhythmTimer() {
			_timer = new Timer {
				Interval = 500
			};
		}

		public void Start(EventHandler onTick) {
			_timer.Tick -= onTick;
			_timer.Start();
			_timer.Tick += onTick;
		}

		public void Stop() {
			if (_timer == null) return;

			_timer.Stop();
			_timer.Dispose();
		}
	}
}