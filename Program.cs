using System;
using System.Threading;
using System.Windows.Forms;

namespace ClipboardImageLogger
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            bool createdNew;

            using var mutex = new Mutex(true, "ClipboardImageLogger_SingleInstance_Mutex", out createdNew);

            if (!createdNew)
            {
                MessageBox.Show(
                    "Clipboard Image Logger is already running.",
                    "Clipboard Image Logger",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
