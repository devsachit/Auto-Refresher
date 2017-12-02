using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

namespace Auto_refresher
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        [System.Runtime.InteropServices.DllImport("Shell32.dll")]
        private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);

        public static void RefreshWindowsExplorer()
        {
            // Refresh the desktop
            SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);                                             
            Guid CLSID_ShellApplication = new Guid("13709620-C279-11CE-A49E-444553540000");
            Type shellApplicationType = Type.GetTypeFromCLSID(CLSID_ShellApplication, true);

            object shellApplication = Activator.CreateInstance(shellApplicationType);
            object windows = shellApplicationType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shellApplication, new object[] { });

            Type windowsType = windows.GetType();
            object count = windowsType.InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windows, null);
            for (int i = 0; i < (int)count; i++)
            {
                object item = windowsType.InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windows, new object[] { i });
                Type itemType = item.GetType();

                // Only refresh Windows Explorer, without checking for the name this could refresh open IE windows
                string itemName = (string)itemType.InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, item, null);
                if (itemName == "Windows Explorer")
                {
                    itemType.InvokeMember("Refresh", System.Reflection.BindingFlags.InvokeMethod, null, item, null);
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {    
            RefreshWindowsExplorer();
            ClearRAM();
        }

        private void ClearRAM()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.WaitForFullGCComplete();
            GC.WaitForFullGCApproach();
            GC.Collect();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            notifyIcon1.Text = "Auto Refresher";
            notifyIcon1.ShowBalloonTip(1000, "Auto Refresher", "Auto refresh in every " + timer1.Interval + " miliseconds", ToolTipIcon.Info);
                
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            notifyIcon1.Dispose();
            ClearRAM();
            this.Hide();
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            toolStripTextBox1.Text = timer1.Interval.ToString();
        }

        private void toolStripTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == 9)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void contextMenuStrip1_Closing(object sender, ToolStripDropDownClosingEventArgs e)
        {
            timer1.Interval =int.Parse( toolStripTextBox1.Text);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(1);
        }

        private void label2_Click(object sender, EventArgs e)
        {
            Process.Start(@"http:\\www.facebook.com\sachit.devkota");
        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Opacity = 100;
            this.Show();
        }

        private void label3_Click(object sender, EventArgs e)
        {
            this.Opacity = 0;
            this.Hide();
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            RefreshWindowsExplorer();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RefreshWindowsExplorer();
        }
    }
}
