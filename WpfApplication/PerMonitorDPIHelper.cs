using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace WpfApplicationDPI
{
	class PerMonitorDPIHelper
	{
		public static bool SetPerMonitorDPIAware()
		{
			return SetProcessDpiAwareness(PROCESS_DPI_AWARENESS.PROCESS_PER_MONITOR_DPI_AWARE) == 0;
		}

		[DllImport("Shcore")]
		private static extern int SetProcessDpiAwareness(PROCESS_DPI_AWARENESS awareness);

		public static PROCESS_DPI_AWARENESS GetPerMonitorDPIAware()
		{
			PROCESS_DPI_AWARENESS awareness;
			int result = GetProcessDpiAwareness(IntPtr.Zero, out awareness);

			if (result != 0)
			{
				throw new Exception("Unable to read process DPI level");
			}
			return awareness;
			
		}

		[DllImport("Shcore")]
		private static extern int GetProcessDpiAwareness(IntPtr hprocess, out PROCESS_DPI_AWARENESS awareness);

		public static int GetDpiForWindow(IntPtr hwnd)
		{
			IntPtr hmonitor = MonitorFromWindow(hwnd, MonitorFromWindowFlags.DefaultToNearest);
			int newDpiX;
			int newDpiY;
			if (GetDpiForMonitor(hmonitor, MonitorDpiTypes.EffectiveDPI, out newDpiX, out newDpiY) != 0)
			{
				return 96;
			}
			return newDpiX;
			
		}

		[DllImport("Shcore")]
		private static extern int GetDpiForMonitor(IntPtr hmonitor, MonitorDpiTypes dpiType, out int dpiX, out int dpiY);

		private enum MonitorDpiTypes
		{
			EffectiveDPI = 0,
			AngularDPI = 1,
			RawDPI = 2,
		}

		public static int GetSystemDPI()
		{
			int newDpiX = 0;
			IntPtr dc = GetDC(IntPtr.Zero);
			newDpiX = GetDeviceCaps(dc, LOGPIXELSX);
			ReleaseDC(IntPtr.Zero, dc);
			return newDpiX;
			
		}

		[DllImport("user32.dll")]
		private static extern IntPtr GetDC(IntPtr hWnd);

		[DllImport ("user32")]
		private static extern int ReleaseDC (IntPtr hWnd, IntPtr hdc);

		[DllImport("gdi32.dll")]
		private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

		private const int LOGPIXELSX = 88;

		public static IntPtr GetMonitorFromWindow(Window wnd)
		{
			var source = (HwndSource)PresentationSource.FromVisual(wnd);
			return MonitorFromWindow(source.Handle, MonitorFromWindowFlags.DefaultToNearest);
		}

		[DllImport("user32.dll")]
		private static extern IntPtr MonitorFromWindow(IntPtr hwnd, MonitorFromWindowFlags dwFlags);

		[Flags]
		private enum MonitorFromWindowFlags
		{
			DefaultToNull = 0,
			DefaultToPrimary = 1,
			DefaultToNearest = 2,			
		}
	}
}
