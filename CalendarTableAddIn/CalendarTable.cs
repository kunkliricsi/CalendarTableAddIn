using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using Nager.Date;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace CalendarTableAddIn
{
    public class CalendarTable
    {
        private Word.Range Range { get; set; }
        private Word.Document Document { get; set; }
        private Word.Table Table { get; set; }

        private int Rows { get; set; }
        private int Columns { get; set; }

        private Random Random { get; set; }
        private HashSet<Word.WdColor> TopRowColors { get; set; }
        private Word.WdColor RandomColor => TopRowColors.ElementAt(Random.Next(TopRowColors.Count));

        private Dictionary<DateTime, (int row, int column)> DaysToCells { get; set; }
        
        public CalendarTable(Word.Document document, Word.Range range)
        {
            this.DaysToCells = new Dictionary<DateTime, (int, int)>();

            this.Document = document;
            this.Range = range;

            this.Rows = 8;
            this.Columns = 7;

            this.TopRowColors = new HashSet<Word.WdColor>()
            {
                Word.WdColor.wdColorDarkBlue,
                Word.WdColor.wdColorDarkGreen,
                Word.WdColor.wdColorDarkRed,
                Word.WdColor.wdColorDarkTeal,
                Word.WdColor.wdColorDarkYellow,
            };

            this.Random = new Random();
        }

        public void Create()
        {
            try
            {
                this.Table = this.Document.Tables.Add(
                    this.Range,
                    this.Rows,
                    this.Columns);

                this.InitializeTable(this.Table);

                this.MakeFirstRow();
                this.MakeSecondRow();

                this.FillTable();
                
                Task.Run(async () => FillGoogleWorkdays(
                    await GoogleCalendar.UpdateWorkdaysAsync(
                        DaysToCells.First().Key,
                        DaysToCells.Last().Key)));
                
                this.SetBorders(this.Table);
                
                this.AddInstructions();
            }
            catch { }
        }

        private void InitializeTable(Word.Table table)
        {
            // Set table text alignment to center.
            table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // Set table widths
            table.AllowAutoFit = true;
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
        }

        private void MakeFirstRow()
        {
            var monthNumber = DateTime.Now.ToString("MM");
            var monnth = DateTime.Now.ToString("MMMM");

            this.Table.Rows[1].Range.Text = String.Format("{0} ({1})", monnth, monthNumber);
            this.Table.Rows[1].Range.Font.Color = Word.WdColor.wdColorWhite;
            this.Table.Rows[1].Cells.Merge();
            this.Table.Rows[1].Shading.BackgroundPatternColor = RandomColor;
        }

        private void MakeSecondRow()
        {
            this.Table.Rows[2].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
            this.Table.Rows[2].Range.Font.Bold = 1;
            this.Table.Cell(2, 1).Range.Text = "V";
            this.Table.Cell(2, 2).Range.Text = "H";
            this.Table.Cell(2, 3).Range.Text = "K";
            this.Table.Cell(2, 4).Range.Text = "Sz";
            this.Table.Cell(2, 5).Range.Text = "Cs";
            this.Table.Cell(2, 6).Range.Text = "P";
            this.Table.Cell(2, 7).Range.Text = "Sz";
        }

        // Returns the last day in the table.
        private void FillTable()
        {
            var now = DateTime.Now;

            var firstDay = new DateTime(now.Year, now.Month, 1);
            var daysToRemoveFromFirstDay = this.GetDaysToRemove(firstDay);

            var currentDay = firstDay.AddDays(-daysToRemoveFromFirstDay);
            var firstDayOfNextMonth = firstDay.AddMonths(1);
            var daysToRemoveFromLastDay = this.GetDaysToRemove(firstDayOfNextMonth);
            if (daysToRemoveFromLastDay <= 2)
            {
                firstDayOfNextMonth = firstDayOfNextMonth.AddDays(-daysToRemoveFromLastDay);
            }

            var currentMonthEnded = false;

            for (int r = 3; r <= this.Rows; r++)
            {
                for (int c = 1; c <= this.Columns; c++)
                {
                    currentDay = currentDay.AddDays(1);

                    this.DaysToCells.Add(currentDay, (r, c));
                    this.Table.Cell(r, c).Range.Text = currentDay.Day.ToString();

                    if (DateSystem.IsWeekend(currentDay, CountryCode.HU) || 
                        DateSystem.IsPublicHoliday(currentDay, CountryCode.HU))
                    {
                        this.Table.Cell(r, c).Range.Font.Color = Word.WdColor.wdColorGray25;
                    }

                    if (!currentMonthEnded && currentDay >= firstDayOfNextMonth)
                    {
                        currentMonthEnded = true;
                    }
                }    
                
                if (currentMonthEnded)
                {
                    DeleteEmptyRows(r);
                    return;
                }
            }
        }

        private void DeleteEmptyRows(int fromRow)
        {
            for (int i = this.Rows; i > fromRow; i--)
                this.Table.Rows[i].Delete();

            this.Rows = fromRow;
        }

        private int GetDaysToRemove(DateTime date)
        {
            var dayOfWeek = (int)date.DayOfWeek;

            return dayOfWeek == 6 ? 0 : ++dayOfWeek;
        }

        private void FillGoogleWorkdays(GoogleCalendarUpdateResult calendarUpdateResult)
        {
            // Updating Holidays in Table
            foreach (var day in calendarUpdateResult.holidays)
            {
                var pair = DaysToCells[day];
                this.Table.Cell(pair.row, pair.column).Range.Font.Color = Word.WdColor.wdColorGray25;
            }

            // Updating Workdays in Table
            foreach (var day in calendarUpdateResult.workdays)
            {
                var pair = DaysToCells[day];
                this.Table.Cell(pair.row, pair.column).Range.Font.Color = Word.WdColor.wdColorBlack;
            }
        }

        private void AddInstructions()
        {
            var selection = this.Document.Application.Selection;
            selection.MoveDown(Count: this.Rows);
            selection.HomeKey();
            selection.TypeText("\n");
            selection.MoveUp(Count: 1);

            for (int i = 0; i < 3; i++)
            {
                if (i != 0)
                {
                    selection.MoveEnd(Word.WdUnits.wdTable, 1);
                    selection.MoveRight(Count: 1);
                }

                selection.InsertParagraphAfter();
                selection.MoveDown(Count: 1);
                var table = this.Document.Tables.Add(selection.Range, 1, 2);
                table.Rows.SetLeftIndent(20, Word.WdRulerStyle.wdAdjustFirstColumn);
                this.InitializeTable(table);
                table.Cell(1, 1).Range.Text = "x";
                table.Cell(1, 2).Range.Text = ": availability";
                table.Cell(1, 2).Range.Font.Italic = 1;
                table.Cell(1, 1).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Cell(1, 1).Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth050pt;
            }
            
            selection.MoveEnd(Word.WdUnits.wdTable, 1);
            selection.MoveRight(Count: 1);
            selection.TypeText("\n");
        }

        private void SetBorders(Word.Table table)
        {
            try
            {
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth050pt;

                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth050pt;

                table.Borders.Shadow = true;
            }
            catch { }
        }
    }
}
