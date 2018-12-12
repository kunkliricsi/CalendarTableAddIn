using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace CalendarTableAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;

            var inspector = app.ActiveInspector();
            var mail = inspector.CurrentItem as Outlook.MailItem;
            if (mail != null)
            {
                var document = inspector.WordEditor as Word.Document;
                var range = document.Application.Selection.Range;
                new CalendarTable(document, range).Create();
            }
        }
    }
}
