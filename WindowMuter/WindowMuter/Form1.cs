using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AudioSwitcher.AudioApi.CoreAudio;
using System.Drawing;

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

        private NotifyIcon notifyIcon;
        public Form1()
        {
            InitializeComponent();
            InitializeAudioDevice();




            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;


            notifyIcon = new NotifyIcon();
            notifyIcon.Icon = SystemIcons.Application;
            notifyIcon.Visible = true;
            notifyIcon.Click += notifyIcon_Click;

            // Set up a context menu
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem exitMenuItem = new ToolStripMenuItem("Exit", null, exitMenuItem_Click);
            contextMenu.Items.Add(exitMenuItem);
            notifyIcon.ContextMenuStrip = contextMenu;




            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, Keys.M);


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
            IntPtr foregroundWindow = GetForegroundWindow();
            uint processId;
            GetWindowThreadProcessId(foregroundWindow, out processId);

            foreach (var session in audioController.DefaultPlaybackDevice.SessionController.ActiveSessions())
            {
                if (session.ProcessId == processId)
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

                    return;
                }
            }
        }




        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.ShowInTaskbar = false;
                notifyIcon2.Visible = true;
            }
        }

        private void notifyIcon_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            notifyIcon2.Visible = false;
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
