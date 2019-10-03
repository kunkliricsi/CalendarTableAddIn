using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CalendarTableAddIn
{
    public class MonthPicker : MonthCalendar
    {
        private Action<DateTime> _dateTimeSetter;
        private Form _parent;

        public void Initialize(Form parent, Action<DateTime> dateTimeSetter)
        {
            _parent = parent;
            _dateTimeSetter = dateTimeSetter;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.HWnd == Handle && m.Msg == WM_REFLECT + WM_NOTIFY)
            {
                var nmhdr = (NMHDR)Marshal.PtrToStructure(m.LParam, typeof(NMHDR));
                if (nmhdr.code == MCN_VIEWCHANGE)
                {
                    var nmviewchange = (NMVIEWCHANGE)Marshal.PtrToStructure(m.LParam, typeof(NMVIEWCHANGE));
                    if (nmviewchange.dwOldView == 1 && nmviewchange.dwNewView == 0)
                    {
                        _dateTimeSetter(SelectionStart);
                        _parent.DialogResult = DialogResult.OK;
                        _parent.Close();
                        return;
                    }
                }
            }

            base.WndProc(ref m);
        }

        private const int WM_USER = 0x0400;
        private const int WM_REFLECT = WM_USER + 0x1C00;
        private const int WM_NOTIFY = 0x004E;
        private const int MCN_VIEWCHANGE = -750;

        [StructLayout(LayoutKind.Sequential)]
        private struct NMHDR
        {
            public IntPtr hwndFrom;
            public IntPtr idFrom;
            public int code;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct NMVIEWCHANGE
        {
            public NMHDR nmhdr;
            public uint dwOldView;
            public uint dwNewView;
        }
    }
}
