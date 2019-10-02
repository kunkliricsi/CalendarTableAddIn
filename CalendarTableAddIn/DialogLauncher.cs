using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalendarTableAddIn
{
    public partial class DialogLauncher : Form
    {
        private DateTime _selectedMonth;

        public DialogLauncher(ref DateTime selectedMonth)
        {
            InitializeComponent();

            _selectedMonth = selectedMonth;
        }
    }
}
