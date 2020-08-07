
using System.Diagnostics;
using System.Text;

namespace System.Runtime.InteropServices
{
    /// <summary>
    /// This is used to get the name of the current window in focus as well as the process of the window. 
    /// </summary>
    class WindowData
    {
        public string WindowTitle { get; set; }
        public string WindowProcess { get; set; }

        [DllImport("user32.dll")]
        static extern int GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(int hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(int hWnd, out uint lpdwProcessId);
        public void GetActiveWindow()
        {
            const int nChars = 256;

            int handle = 0;

            StringBuilder Buff = new StringBuilder(nChars);

            handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                WindowTitle = Buff.ToString();

                uint lpdwProcessId;

                GetWindowThreadProcessId(handle, out lpdwProcessId);

                WindowProcess = Process.GetProcessById((int)lpdwProcessId).ProcessName;
            }

        }


    }
}