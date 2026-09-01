using System;
using System.Runtime.InteropServices;
using LightFeather.Shared;
using Microsoft.Office.Interop.Word;

namespace LightFeather.Extensions {
	public static class WindowExtensions {
		public static IntPtr GetOwnerHandle(this Window window) {
			try {
				if (window == null)
					return IntPtr.Zero;

				var hwnd = new IntPtr(window.Hwnd);
				var root = WindowNative.GetAncestor(hwnd, WindowNative.GA_ROOT);

				return root != IntPtr.Zero ? root : hwnd;
			}
			catch (COMException) {
				return IntPtr.Zero;
			}
		}
	}
}