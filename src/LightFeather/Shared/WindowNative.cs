using System;
using System.Runtime.InteropServices;

namespace LightFeather.Shared {
	/// <summary>
	/// Win32 potrzebny, żeby nasze okna były satelitami Worda, a nie zwykłymi formami
	/// gubiącymi się pod dokumentem.
	/// </summary>
	internal static class WindowNative {
		internal const int GWL_HWNDPARENT = -8;

		/// <summary>GA_ROOT - okno najwyższego poziomu (OpusApp).</summary>
		internal const uint GA_ROOT = 2;

		[StructLayout(LayoutKind.Sequential)]
		internal struct RECT {
			public int Left;
			public int Top;
			public int Right;
			public int Bottom;
		}

		[DllImport("user32.dll")]
		internal static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);

		[DllImport("user32.dll")]
		internal static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

		[DllImport("user32.dll", EntryPoint = "SetWindowLong")]
		private static extern int SetWindowLong32(IntPtr hWnd, int index, int newLong);

		[DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
		private static extern IntPtr SetWindowLong64(IntPtr hWnd, int index, IntPtr newLong);

		[DllImport("user32.dll")]
		private static extern int GetDpiForWindow(IntPtr hWnd);

		internal static void SetWindowLong(IntPtr hWnd, int index, IntPtr value) {
			if (IntPtr.Size == 8)
				SetWindowLong64(hWnd, index, value);
			else
				SetWindowLong32(hWnd, index, value.ToInt32());
		}

		internal static float GetScale(IntPtr hWnd) {
			try {
				var dpi = GetDpiForWindow(hWnd);

				if (dpi > 0)
					return dpi / 96f;
			}
			catch (EntryPointNotFoundException) {
				// Windows < 10 1607.
			}

			return 1f;
		}
	}
}
