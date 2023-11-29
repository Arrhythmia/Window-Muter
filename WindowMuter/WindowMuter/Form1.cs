using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using AudioSwitcher.AudioApi.CoreAudio;
using AudioSwitcher.AudioApi.Session;
using System.Reflection;

namespace WindowMuter
{
    public partial class Form1 : Form
    {
        // Constants for hotkey modifiers
        private const int MOD_CONTROL = 0x0002;

        // Windows API function imports
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, Keys vk);
        // Constants for hotkey IDs
        private const int HOTKEY_ID = 1;

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private CoreAudioController audioController;
        public Form1()
        {
            InitializeComponent();
            InitializeAudioDevice();

            this.Resize += Form1_Resize;
            this.Activated += Form1_Activated;

            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;

            notifyIcon.Visible = true;
            notifyIcon.MouseClick += notifyIcon_MouseClick;
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem exitMenuItem = new ToolStripMenuItem("Exit", null, exitMenuItem_Click);
            contextMenu.Items.Add(exitMenuItem);
            notifyIcon.ContextMenuStrip = contextMenu;

            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, Keys.M);
        }
        private void notifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(notifyIcon, null);
            }
            else if (e.Button == MouseButtons.Left)
            {
                base.WindowState = FormWindowState.Normal;
                base.ShowInTaskbar = true;
            }
        }

        private void exitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void InitializeAudioDevice()
        {
            audioController = new CoreAudioController();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, Keys.M);
        }
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 0x0312 && m.WParam.ToInt32() == HOTKEY_ID)
            {
                ToggleMuteForActiveWindow();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Register Ctrl+M as the global hotkey
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, Keys.M);
        }



        private void ToggleMuteForActiveWindow()
        {
            GetWindowThreadProcessId(GetForegroundWindow(), out var processId);
            string processName = Process.GetProcessById((int)processId).ProcessName;
            Console.WriteLine("Attempting to mute: " + Process.GetProcessById((int)processId).ProcessName);
            foreach (IAudioSession session in audioController.DefaultPlaybackDevice.SessionController.ActiveSessions())
            {
                if (session.ProcessId == processId || processName == Process.GetProcessById(session.ProcessId).ProcessName)
                {
                    session.IsMuted = !session.IsMuted;
                    break;
                }
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.ShowInTaskbar = false;
                notifyIcon.Visible = true;
                RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, Keys.M);
            }
        }
    }
}
