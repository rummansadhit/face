Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    using System.Text;

    public class SessionLock {
        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern IntPtr WTSOpenServer([MarshalAs(UnmanagedType.LPStr)] string pServerName);

        [DllImport("wtsapi32.dll")]
        public static extern void WTSCloseServer(IntPtr hServer);

        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSEnumerateSessions(
            IntPtr hServer,
            [MarshalAs(UnmanagedType.U4)] int reserved,
            [MarshalAs(UnmanagedType.U4)] int version,
            ref IntPtr ppSessionInfo,
            [MarshalAs(UnmanagedType.U4)] ref int pCount);

        [DllImport("wtsapi32.dll")]
        public static extern void WTSFreeMemory(IntPtr pMemory);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool LockWorkStation();

        [StructLayout(LayoutKind.Sequential)]
        public struct WTS_SESSION_INFO {
            public Int32 SessionID;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pWinStationName;
            public WTS_CONNECTSTATE_CLASS State;
        }

        public enum WTS_CONNECTSTATE_CLASS {
            WTSActive,
            WTSConnected,
            WTSConnectQuery,
            WTSShadow,
            WTSDisconnected,
            WTSIdle,
            WTSListen,
            WTSReset,
            WTSDown,
            WTSInit
        }

        public static void LockAllSessions() {
            IntPtr serverHandle = WTSOpenServer(Environment.MachineName);
            try {
                IntPtr ppSessionInfo = IntPtr.Zero;
                int count = 0;
                int dataSize = Marshal.SizeOf(typeof(WTS_SESSION_INFO));
                if (WTSEnumerateSessions(serverHandle, 0, 1, ref ppSessionInfo, ref count)) {
                    IntPtr current = ppSessionInfo;
                    for (int i = 0; i < count; i++) {
                        WTS_SESSION_INFO si = (WTS_SESSION_INFO)Marshal.PtrToStructure(current, typeof(WTS_SESSION_INFO));
                        if (si.State == WTS_CONNECTSTATE_CLASS.WTSActive) {
                            LockWorkStation();
                        }
                        current = IntPtr.Add(current, dataSize);
                    }
                    WTSFreeMemory(ppSessionInfo);
                }
            } finally {
                WTSCloseServer(serverHandle);
            }
        }
    }
"@

[SessionLock]::LockAllSessions()
