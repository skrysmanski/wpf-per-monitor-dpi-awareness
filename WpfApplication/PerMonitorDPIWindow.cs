using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace WpfApplicationDPI
{
	public class PerMonitorDPIWindow : Window
	{
		private const int WM_DPICHANGED = 0x02E0;

		private readonly bool m_perMonitorEnabled;
		
		private int m_currentDPI;		
		
		private double m_systemDPI;
		
		private double m_wpfDPI;
		
		private double m_scaleFactor;		
		
		private HwndSource m_source;

		public event EventHandler DPIChanged;

		public int CurrentDPI
		{
			get { return m_currentDPI; }
		}

		public double WpfDPI
		{
			get { return m_wpfDPI; }
			set { m_wpfDPI = value; }
		}

		public double ScaleFactor
		{
			get { return m_scaleFactor; }
		}

		public PerMonitorDPIWindow()
		{
			Loaded += OnLoaded;
			if (PerMonitorDPIHelper.SetPerMonitorDPIAware())
			{
				m_perMonitorEnabled = true;
			}
			else
			{
				throw new Exception("Enabling Per-monitor DPI Failed.  Do you have [assembly: DisableDpiAwareness] in your assembly manifest [AssemblyInfo.cs]?");
			}
			
		}

		public string GetCurrentDpiConfiguration()
		{
			StringBuilder stringBuilder = new StringBuilder();		

			var awareness = PerMonitorDPIHelper.GetPerMonitorDPIAware();

			var systemDpi = PerMonitorDPIHelper.GetSystemDPI();

			switch (awareness)
			{
			case PROCESS_DPI_AWARENESS.PROCESS_DPI_UNAWARE:
				stringBuilder.AppendFormat("Application is DPI Unaware.  Using {0} DPI.", systemDpi);
				break;
			case PROCESS_DPI_AWARENESS.PROCESS_SYSTEM_DPI_AWARE:
				stringBuilder.AppendFormat("Application is System DPI Aware.  Using System DPI: {0}.", systemDpi);
				break;
			case PROCESS_DPI_AWARENESS.PROCESS_PER_MONITOR_DPI_AWARE:			
				stringBuilder.AppendFormat("Application is Per-Monitor DPI Aware.  Using \tmonitor DPI = {0}  \t(System DPI = {1}).", m_currentDPI, systemDpi);
				break;
			}

			return stringBuilder.ToString();
		}

		//OnLoaded Handler: Adjusts the window size and graphics and text size based on current DPI of the Window
		[EnvironmentPermission(SecurityAction.LinkDemand, Unrestricted = true)]
		protected void OnLoaded(Object sender, RoutedEventArgs args)
		{
			// WPF has already scaled window size, graphics and text based on system DPI. In order to scale the window based on monitor DPI, update the 
			// window size, graphics and text based on monitor DPI. For example consider an application with size 600 x 400 in device independent pixels
			//		- Size in device independent pixels = 600 x 400 
			//		- Size calculated by WPF based on system/WPF DPI = 192 (scale factor = 2)
			//		- Expected size based on monitor DPI = 144 (scale factor = 1.5)

			// Similarly the graphics and text are updated updated by applying appropriate scale transform to the top level node of the WPF application

			// Important Note: This method overwrites the size of the window and the scale transform of the root node of the WPF Window. Hence, 
			// this sample may not work "as is" if 
			//	- The size of the window impacts other portions of the application like this WPF  Window being hosted inside another application. 
			//  - The WPF application that is extending this class is setting some other transform on the root visual; the sample may 
			//     overwrite some other transform that is being applied by the WPF application itself.
		
			if (m_perMonitorEnabled)
			{
				m_source = (HwndSource) PresentationSource.FromVisual(this);
				m_source.AddHook(this.HandleMessages);	
			
			
				//Calculate the DPI used by WPF; this is same as the system DPI. 
					
				m_wpfDPI = 96.0 *  m_source.CompositionTarget.TransformToDevice.M11; 

				//Get the Current DPI of the monitor of the window. 
					
				m_currentDPI = PerMonitorDPIHelper.GetDpiForWindow(m_source.Handle);

				//Calculate the scale factor used to modify window size, graphics and text
				m_scaleFactor = m_currentDPI / m_wpfDPI; 
		
				//Update Width and Height based on the on the current DPI of the monitor
			
				Width = Width * m_scaleFactor;
				Height = Height * m_scaleFactor;

				//Update graphics and text based on the current DPI of the monitor
			
				UpdateLayoutTransform(m_scaleFactor);
			}			
		}

		//Called when the DPI of the window changes. This method adjusts the graphics and text size based on the new DPI of the window
		protected void OnDPIChanged()
		{
			m_scaleFactor = m_currentDPI / m_wpfDPI;
			UpdateLayoutTransform(m_scaleFactor);
			DPIChanged(this, EventArgs.Empty);
		}

		protected void UpdateLayoutTransform(double scaleFactor)
		{
			// Adjust the rendering graphics and text size by applying the scale transform to the top level visual node of the Window		
			if (m_perMonitorEnabled) 
			{		
				var child = GetVisualChild(0);
				if (m_scaleFactor != 1.0) {
					ScaleTransform dpiScale = new ScaleTransform(scaleFactor, scaleFactor);
					child.SetValue(LayoutTransformProperty, dpiScale);
				}
				else 
				{
					child.SetValue(LayoutTransformProperty, null);
				}			
			}
		}

		// Message handler of the Per_Monitor_DPI_Aware window. The handles the WM_DPICHANGED message and adjusts window size, graphics and text
		// based on the DPI of the monitor. The window message provides the new window size (lparam) and new DPI (wparam)
		protected IntPtr HandleMessages(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
		{
			double oldDpi;

			switch (msg)
			{
				case WM_DPICHANGED:
					RECT lprNewRect = (RECT)Marshal.PtrToStructure(lParam, typeof(RECT));
					SetWindowPos(hwnd, IntPtr.Zero, lprNewRect.Left, lprNewRect.Top, lprNewRect.Width, lprNewRect.Height,
								 SetWindowPosFlags.NOZORDER | SetWindowPosFlags.NOOWNERZORDER | SetWindowPosFlags.NOACTIVATE);
					oldDpi = m_currentDPI;
					m_currentDPI = wParam.ToInt32() & 0xFFFF;
					if (oldDpi != m_currentDPI) 
					{
						OnDPIChanged();
					}
				break;
			}

			return IntPtr.Zero;
		}

		[DllImport("user32.dll", SetLastError=true)]
		static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int width, int height, SetWindowPosFlags uFlags);

		[Flags]
		private enum SetWindowPosFlags
		{
		   NOSIZE = 0x0001,
		   NOMOVE = 0x0002,
		   NOZORDER = 0x0004,
		   NOREDRAW = 0x0008,
		   NOACTIVATE = 0x0010,
		   DRAWFRAME = 0x0020,
		   FRAMECHANGED = 0x0020,
		   SHOWWINDOW = 0x0040,
		   HIDEWINDOW = 0x0080,
		   NOCOPYBITS = 0x0100,
		   NOOWNERZORDER = 0x0200,
		   NOREPOSITION = 0x0200,
		   NOSENDCHANGING = 0x0400,
		   DEFERERASE = 0x2000,
		   ASYNCWINDOWPOS = 0x4000,
		}

		[StructLayout(LayoutKind.Sequential)]
		public struct RECT
		{
			public int Left, Top, Right, Bottom;

			public RECT(int left, int top, int right, int bottom)
			{
				Left = left;
				Top = top;
				Right = right;
				Bottom = bottom;
			}

			public int Height
			{
				get { return Bottom - Top; }
			}

			public int Width
			{
				get { return Right - Left; }
			}
		}
	};
}
