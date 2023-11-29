using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AudioSwitcher.AudioApi;
using System.Drawing;
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

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
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
            // Handle exit menu item click
            Application.Exit();
        }

        private void InitializeAudioDevice()
        {
            audioController = new CoreAudioController();
        }


        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            // Handle hotkey message
            if (m.Msg == 0x0312 && m.WParam.ToInt32() == HOTKEY_ID)
            {
                // Call your method to mute/unmute the focused application
                ToggleMuteForActiveWindow();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Register Ctrl+M as the global hotkey
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, Keys.M);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Unregister the hotkey when the form is closing
            UnregisterHotKey(this.Handle, HOTKEY_ID);
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
                    if (session.IsMuted)
                    {
                        statusLabel.Text = "Application Muted";
                    }
                    else
                    {
                        statusLabel.Text = "Application Unmuted";
                    }
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
            }
        }

        private void notifyIcon_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            notifyIcon.Visible = false;
        }

        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            // Restore the form when double-clicking the tray icon
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            notifyIcon.Visible = false;
        }
    }
}
