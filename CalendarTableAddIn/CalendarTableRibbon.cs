using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace CalendarTableAddIn
{
    public partial class CalendarTableRibbon
    {
        private ICalendarTableFactory _calendarTableFactory;
        private Word.Document _document;

        private DialogLauncher _dialogLauncher;
        private DateTime _selectedMonth;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            _dialogLauncher = new DialogLauncher(d => _selectedMonth = d);

            var app = Globals.ThisAddIn.Application;
            var inspector = app.ActiveInspector();
            var mail = inspector.CurrentItem as Outlook.MailItem;
            _document = inspector.WordEditor as Word.Document;

            _calendarTableFactory = new CalendarTableFactory(_document);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            _calendarTableFactory.Create(DateTime.Now);
        }

        private void GroupInsertCalendarTables_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            var result = _dialogLauncher.ShowDialog();
            if (result == DialogResult.OK)
            {
                _calendarTableFactory.Create(_selectedMonth);
            }
        }
    }
}
