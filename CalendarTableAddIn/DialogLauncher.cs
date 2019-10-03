using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CalendarTableAddIn
{
    public partial class DialogLauncher : Form
    {
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

        private const int MCM_SETCURRENTVIEW = 0x1000 + 32;

        public DialogLauncher(Action<DateTime> dateTimeSetter)
        {
            InitializeComponent();
            monthPicker1.Initialize(this, dateTimeSetter);

            Load += DialogLauncher_Load;
        }

        private void DialogLauncher_Load(object sender, EventArgs e)
        {
            monthPicker1.SetDate(DateTime.Now);
            SendMessage(monthPicker1.Handle, MCM_SETCURRENTVIEW, IntPtr.Zero, (IntPtr)1);
            SetDesktopLocation(Cursor.Position.X, Cursor.Position.Y);
        }
    }
}
