using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace WindowsLayoutSnapshot {

    internal class Snapshot {

        private Dictionary<IntPtr, WINDOWPLACEMENT> m_placements = new Dictionary<IntPtr, WINDOWPLACEMENT>();
		private Dictionary<IntPtr, string> m_winText = new Dictionary<IntPtr, string>();
		private List<IntPtr> m_windowsBackToTop = new List<IntPtr>();
		private string name = null;
		
		private Snapshot(bool userInitiated, string name)
		{
			this.name = name;
			EnumWindows(EvalWindow, 0);

			TimeTaken = DateTime.UtcNow;
			UserInitiated = userInitiated;

			var pixels = new List<long>();
			foreach (var screen in Screen.AllScreens)
			{
				pixels.Add(screen.Bounds.Width * screen.Bounds.Height);
			}
			MonitorPixelCounts = pixels.ToArray();
			NumMonitors = pixels.Count;
		}
		private Snapshot(string str)
		{
			string[] strList = str.Split(new string[] { "|||" }, StringSplitOptions.None);
			this.name = strList[0];
			NumMonitors = Convert.ToInt32(strList[1]);
			UserInitiated = Convert.ToBoolean(strList[2]);

			string[] sTemp = strList[3].Split(new string[] { "###" }, StringSplitOptions.None);
			var pixels = new List<long>();
			foreach (var s in sTemp)
			{
				if (!s.Equals(""))
				{
					pixels.Add(Convert.ToInt64(s));
				}
			}
			MonitorPixelCounts = pixels.ToArray();

			sTemp = strList[4].Split(new string[] { "###" }, StringSplitOptions.None);
			foreach (var s in sTemp)
			{
				if (!s.Equals(""))
				{
					m_windowsBackToTop.Add(new IntPtr(Convert.ToInt32(s)));
				}
			}

			sTemp = strList[5].Split(new string[] { "###" }, StringSplitOptions.None);
			foreach (var s in sTemp)
			{
				if (!s.Equals(""))
				{
					string[] sTemp2 = s.Split(new string[] { "~~~" }, StringSplitOptions.None);
					m_winText.Add(new IntPtr(Convert.ToInt32(sTemp2[0])), sTemp2[1]);
				}
			}

			sTemp = strList[6].Split(new string[] { "###" }, StringSplitOptions.None);
			foreach (var s in sTemp)
			{
				if (!s.Equals(""))
				{
					string[] sTemp2 = s.Split(new string[] { "~~~" }, StringSplitOptions.None);
					string[] minXY = sTemp2[4].Split(',');
					string[] maxXY = sTemp2[5].Split(',');
					string[] rectDim = sTemp2[6].Split(',');
					RECT tempRect = new RECT();
					tempRect.Left = Convert.ToInt32(rectDim[0]);
					tempRect.Top = Convert.ToInt32(rectDim[1]);
					tempRect.Right = Convert.ToInt32(rectDim[2]);
					tempRect.Bottom = Convert.ToInt32(rectDim[3]);

					WINDOWPLACEMENT tempPlacement = new WINDOWPLACEMENT();
					tempPlacement.length = Convert.ToInt32(sTemp2[1]);
					tempPlacement.flags = Convert.ToInt32(sTemp2[2]);
					tempPlacement.showCmd = Convert.ToInt32(sTemp2[3]);
					tempPlacement.ptMinPosition = new Point(Convert.ToInt32(minXY[0]), Convert.ToInt32(minXY[1]));
					tempPlacement.ptMaxPosition = new Point(Convert.ToInt32(maxXY[0]), Convert.ToInt32(maxXY[1]));
					tempPlacement.rcNormalPosition = tempRect;

					m_placements.Add(new IntPtr(Convert.ToInt32(sTemp2[0])), tempPlacement);
				}
			}
		}
		internal static Snapshot TakeSnapshot(bool userInitiated) {
            return new Snapshot(userInitiated, null);
        }
		internal static Snapshot TakeSnapshot(string name)
		{
			return new Snapshot(true, name);
		}
		internal static Snapshot LoadSnapshot(string info)
		{
			return new Snapshot(info);
		}
		public string ConvertToString()
		{
			string s = "";
			s += (name == null ? "" : name) + "|||";
			s += (NumMonitors == 0 ? "" : NumMonitors.ToString()) + "|||";
			s += UserInitiated + "|||";

			for (int i = 0; i < MonitorPixelCounts.Length; i++)
			{
				s += MonitorPixelCounts[i] + "###";
			}
			s += "|||";
			
			foreach (var ptr in m_windowsBackToTop)
			{
				s += ptr + "###";
			}
			s += "|||";

			foreach (var ptr in m_winText.Keys)
			{
				s += ptr + "~~~" + m_winText[ptr]+ "###";
			}
			s += "|||";


			foreach (var ptr in m_placements.Keys)
			{
				s += ptr + "~~~";
				WINDOWPLACEMENT a = m_placements[ptr];
				s += a.length + "~~~";
				s += a.flags + "~~~";
				s += a.showCmd + "~~~";
				s += a.ptMinPosition.X + "," + a.ptMinPosition.Y + "~~~";
				s += a.ptMaxPosition.X + "," + a.ptMinPosition.Y + "~~~";
				s += a.rcNormalPosition.Left + "," + a.rcNormalPosition.Top + "," + a.rcNormalPosition.Right + "," + a.rcNormalPosition.Bottom + "###";
			}
			s += "|||";

			return s;
		}
		private bool EvalWindow(int hwndInt, int lParam) {
            var hwnd = new IntPtr(hwndInt);

            if (!IsAltTabWindow(hwnd)) {
                return true;
            }

            // EnumWindows returns windows in Z order from back to front
            m_windowsBackToTop.Add(hwnd);

            var placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(typeof(WINDOWPLACEMENT));
            if (!GetWindowPlacement(hwnd, ref placement)) {
                throw new Exception("Error getting window placement");
            }
            m_placements.Add(hwnd, placement);
			m_winText.Add(hwnd, GetWindowText(hwnd)); 

			return true;
        }
        internal DateTime TimeTaken { get; private set; }
        internal bool UserInitiated { get; private set; }
        internal long[] MonitorPixelCounts { get; private set; }
        internal int NumMonitors { get; private set; }

        internal TimeSpan Age {
            get { return DateTime.UtcNow.Subtract(TimeTaken); }
        }

		internal string Name {
			get { return name; }
		}

		internal void RestoreAndPreserveMenu(object sender, EventArgs e) { // ignore extra params
            // We save and restore the current foreground window because it's our tray menu
            // I couldn't find a way to get this handle straight from the tray menu's properties;
            //   the ContextMenuStrip.Handle isn't the right one, so I'm using win32
            // More info RE the restore is below, where we do it
            var currentForegroundWindow = GetForegroundWindow();

            try {
                Restore(sender, e);
            } finally {
                // A combination of SetForegroundWindow + SetWindowPos (via set_Visible) seems to be needed
                // This was determined by trying a bunch of stuff
                // This prevents the tray menu from closing, and makes sure it's still on top
                SetForegroundWindow(currentForegroundWindow);
                TrayIconForm.me.Visible = true;
            }
        }

        internal void Restore(object sender, EventArgs e) { // ignore extra params
            // first, restore the window rectangles and normal/maximized/minimized states
            foreach (var placement in m_placements) {
                // this might error out if the window no longer exists
                var placementValue = placement.Value;

				// make sure points and rects will be inside monitor
				IntPtr extendedStyles = GetWindowLongPtr(placement.Key, (-20)); // GWL_EXSTYLE
                placementValue.ptMaxPosition = GetUpperLeftCornerOfNearestMonitor(extendedStyles, placementValue.ptMaxPosition);
                placementValue.ptMinPosition = GetUpperLeftCornerOfNearestMonitor(extendedStyles, placementValue.ptMinPosition);
                placementValue.rcNormalPosition = GetRectInsideNearestMonitor(extendedStyles, placementValue.rcNormalPosition);

				if (SetWindowPlacement(placement.Key, ref placementValue) == false)
				{
					var winList = FindWindowsWithText(m_winText[placement.Key]);
					foreach (var hWndInt in winList)
					{
						IntPtr temp = new IntPtr(hWndInt);
						if (m_placements.ContainsKey(temp))
						{
							continue;
						}
						SetWindowPlacement(temp, ref placementValue);
						break;
					}
				}
            }

            // now update the z-orders
            m_windowsBackToTop = m_windowsBackToTop.FindAll(IsWindowVisible);
            IntPtr positionStructure = BeginDeferWindowPos(m_windowsBackToTop.Count);
            for (int i = 0; i < m_windowsBackToTop.Count; i++) {
                positionStructure = DeferWindowPos(positionStructure, m_windowsBackToTop[i], i == 0 ? IntPtr.Zero : m_windowsBackToTop[i - 1],
                    0, 0, 0, 0, DeferWindowPosCommands.SWP_NOMOVE | DeferWindowPosCommands.SWP_NOSIZE | DeferWindowPosCommands.SWP_NOACTIVATE);
            }
            EndDeferWindowPos(positionStructure);
        }

		/// <summary> Get the text for the window pointed to by hWnd </summary>
		public static string GetWindowText(IntPtr hWnd)
		{
			int size = GetWindowTextLength(hWnd);
			if (size > 0)
			{
				var builder = new StringBuilder(size + 1);
				GetWindowText(hWnd, builder, builder.Capacity);
				return builder.ToString();
			}

			return String.Empty;
		}

		
		/// <summary> Find all windows that match the given filter </summary>
		/// <param name="filter"> A delegate that returns true for windows
		///    that should be returned and false for windows that should
		///    not be returned </param>
		private static IEnumerable<int> FindWindows(EnumWindowsProc filter)
		{
			List<int> windows = new List<int>();

			EnumWindows(delegate (int wnd, int param)
			{
				if (filter(wnd, param))
				{
					// only add the windows that pass the filter
					windows.Add(wnd);
				}

				// but return true here so that we iterate all windows
				return true;
			}, 0);

			return windows;
		}

		/// <summary> Find all windows that contain the given title text </summary>
		/// <param name="titleText"> The text that the window title must contain. </param>
		public static IEnumerable<int> FindWindowsWithText(string titleText)
		{
			return FindWindows(delegate (int wnd, int param)
			{
				var hwnd = new IntPtr(wnd);
				return GetWindowText(hwnd).Contains(titleText);
			});
		}

		private static Point GetUpperLeftCornerOfNearestMonitor(IntPtr windowExtendedStyles, Point point) {
            if ((windowExtendedStyles.ToInt64() & 0x00000080) > 0) { // WS_EX_TOOLWINDOW
                return Screen.GetBounds(point).Location; // use screen coordinates
            } else {
                return Screen.GetWorkingArea(point).Location; // use workspace coordinates
            }
        }

        private static RECT GetRectInsideNearestMonitor(IntPtr windowExtendedStyles, RECT rect) {
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            Rectangle rectAsRectangle = new Rectangle(rect.Left, rect.Top, width, height);
            Rectangle monitorRect;
            if ((windowExtendedStyles.ToInt64() & 0x00000080) > 0) { // WS_EX_TOOLWINDOW
                monitorRect = Screen.GetBounds(rectAsRectangle); // use screen coordinates
            } else {
                monitorRect = Screen.GetWorkingArea(rectAsRectangle); // use workspace coordinates
            }

            var y = new RECT();
            y.Left = Math.Max(monitorRect.Left, Math.Min(monitorRect.Right - width, rect.Left));
            y.Top = Math.Max(monitorRect.Top, Math.Min(monitorRect.Bottom - height, rect.Top));
            y.Right = y.Left + Math.Min(monitorRect.Width, width);
            y.Bottom = y.Top + Math.Min(monitorRect.Height, height);
            return y;
        }

        private static bool IsAltTabWindow(IntPtr hwnd) {
            if (!IsWindowVisible(hwnd)) {
                return false;
            }

            IntPtr extendedStyles = GetWindowLongPtr(hwnd, (-20)); // GWL_EXSTYLE
            if ((extendedStyles.ToInt64() & 0x00040000) > 0) { // WS_EX_APPWINDOW
                return true;
            }
            if ((extendedStyles.ToInt64() & 0x00000080) > 0) { // WS_EX_TOOLWINDOW
                return false;
            }

            IntPtr hwndTry = GetAncestor(hwnd, GetAncestor_Flags.GetRootOwner);
            IntPtr hwndWalk = IntPtr.Zero;
            while (hwndTry != hwndWalk) {
                hwndWalk = hwndTry;
                hwndTry = GetLastActivePopup(hwndWalk);
                if (IsWindowVisible(hwndTry)) {
                    break;
                }
            }
            if (hwndWalk != hwnd) {
                return false;
            }

            return true;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr BeginDeferWindowPos(int nNumWindows);

        [DllImport("user32.dll")]
        private static extern IntPtr DeferWindowPos(IntPtr hWinPosInfo, IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy,
            [MarshalAs(UnmanagedType.U4)]DeferWindowPosCommands uFlags);

        private enum DeferWindowPosCommands : uint {
            SWP_DRAWFRAME = 0x0020,
            SWP_FRAMECHANGED = 0x0020,
            SWP_HIDEWINDOW = 0x0080,
            SWP_NOACTIVATE = 0x0010,
            SWP_NOCOPYBITS = 0x0100,
            SWP_NOMOVE = 0x0002,
            SWP_NOOWNERZORDER = 0x0200,
            SWP_NOREDRAW = 0x0008,
            SWP_NOREPOSITION = 0x0200,
            SWP_NOSENDCHANGING = 0x0400,
            SWP_NOSIZE = 0x0001,
            SWP_NOZORDER = 0x0004,
            SWP_SHOWWINDOW = 0x0040
        };

        [DllImport("user32.dll")]
        private static extern bool EndDeferWindowPos(IntPtr hWinPosInfo);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        private static extern IntPtr GetWindowLongPtr32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        private static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex) {
            if (IntPtr.Size == 8) {
                return GetWindowLongPtr64(hWnd, nIndex);
            }
            return GetWindowLongPtr32(hWnd, nIndex);
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetLastActivePopup(IntPtr hWnd);

        private enum GetAncestor_Flags {
            GetParent = 1,
            GetRoot = 2,
            GetRootOwner = 3
        }

        [DllImport("user32.dll", ExactSpelling = true)]
        private static extern IntPtr GetAncestor(IntPtr hwnd, GetAncestor_Flags gaFlags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPlacement(IntPtr hWnd, [In] ref WINDOWPLACEMENT lpwndpl);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern int GetWindowTextLength(IntPtr hWnd);
		[StructLayout(LayoutKind.Sequential)]

        private struct WINDOWPLACEMENT {
            public int length;
            public int flags;
            public int showCmd;
            public Point ptMinPosition;
            public Point ptMaxPosition;
            public RECT rcNormalPosition;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int EnumWindows(EnumWindowsProc ewp, int lParam);
        private delegate bool EnumWindowsProc(int hWnd, int lParam);
    }
}
