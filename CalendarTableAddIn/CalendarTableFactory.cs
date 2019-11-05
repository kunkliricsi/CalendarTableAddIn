using Nager.Date;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace CalendarTableAddIn
{
    public class CalendarTableFactory : ICalendarTableFactory
    {
        private readonly Word.Document _document;
        private Word.Table _table;

        private int _rows;
        private readonly int _columns;

        private readonly Random _random;
        private readonly HashSet<Word.WdColor> _topRowColors;
        private Word.WdColor _randomColor => _topRowColors.ElementAt(_random.Next(_topRowColors.Count));

        private readonly Dictionary<DateTime, (int row, int column)> _daysToCells;

        public CalendarTableFactory(Word.Document document)
        {
            _daysToCells = new Dictionary<DateTime, (int, int)>();

            _document = document;

            _rows = 8;
            _columns = 7;

            _topRowColors = new HashSet<Word.WdColor>()
            {
                Word.WdColor.wdColorDarkBlue,
                Word.WdColor.wdColorDarkGreen,
                Word.WdColor.wdColorDarkRed,
                Word.WdColor.wdColorDarkTeal,
                Word.WdColor.wdColorOrange,
            };

            _random = new Random();
        }

        public void Create(DateTime month)
        {
            try
            {
                _daysToCells.Clear();
                _rows = 8;

                var currentSelection = _document.Application.Selection.Range;

                _table = _document.Tables.Add(
                    currentSelection,
                    _rows,
                    _columns);

                InitializeTable(_table);

                MakeFirstRow(month);
                MakeSecondRow();

                FillTable(month);

                var getWorkdaysTask = GoogleCalendar.GetWorkdaysAsync(_daysToCells.First().Key, _daysToCells.Last().Key);
                Task.Run(() => UpdateCalendarAsync(getWorkdaysTask));

                SetBorders(_table);

                AddInstructions();
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

        private void MakeFirstRow(DateTime month)
        {
            var year = month.ToString("yyyy");
            var monthNumber = month.ToString("MM");
            var monthString = month.ToString("MMMM");

            _table.Rows[1].Range.Text = string.Format("{0} {1} ({2})", year, monthString, monthNumber);
            _table.Rows[1].Range.Font.Color = Word.WdColor.wdColorWhite;
            _table.Rows[1].Cells.Merge();
            _table.Rows[1].Shading.BackgroundPatternColor = _randomColor;
        }

        private void MakeSecondRow()
        {
            _table.Rows[2].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
            _table.Rows[2].Range.Font.Bold = 1;
            _table.Cell(2, 1).Range.Text = "V";
            _table.Cell(2, 2).Range.Text = "H";
            _table.Cell(2, 3).Range.Text = "K";
            _table.Cell(2, 4).Range.Text = "Sz";
            _table.Cell(2, 5).Range.Text = "Cs";
            _table.Cell(2, 6).Range.Text = "P";
            _table.Cell(2, 7).Range.Text = "Sz";
        }

        // Returns the last day in the table.
        private void FillTable(DateTime month)
        {
            var firstDay = new DateTime(month.Year, month.Month, 1);

            var currentDay = firstDay;
            while (DateSystem.IsPublicHoliday(currentDay, CountryCode.HU) ||
                DateSystem.IsWeekend(currentDay, CountryCode.HU))
            {
                currentDay = currentDay.AddDays(1);
            }

            while (currentDay.DayOfWeek != DayOfWeek.Saturday)
            {
                currentDay = currentDay.AddDays(-1);
            }

            var lastDay = firstDay.AddMonths(1);
            lastDay = lastDay.AddDays(-1);
            while (DateSystem.IsPublicHoliday(lastDay, CountryCode.HU) ||
                DateSystem.IsWeekend(lastDay, CountryCode.HU))
            {
                lastDay = lastDay.AddDays(-1);
            }

            while (lastDay.DayOfWeek != DayOfWeek.Saturday)
            {
                lastDay = lastDay.AddDays(1);
            }

            for (int r = 3; r <= _rows; r++)
            {
                for (int c = 1; c <= _columns; c++)
                {
                    currentDay = currentDay.AddDays(1);

                    _daysToCells.Add(currentDay, (r, c));
                    _table.Cell(r, c).Range.Text = currentDay.Day.ToString();

                    if (DateSystem.IsWeekend(currentDay, CountryCode.HU) ||
                        DateSystem.IsPublicHoliday(currentDay, CountryCode.HU))
                    {
                        _table.Cell(r, c).Range.Font.Color = Word.WdColor.wdColorGray25;
                    }

                    if (lastDay == currentDay)
                    {
                        DeleteEmptyRows();
                        return;
                    }
                }
            }
        }

        private void DeleteEmptyRows()
        {
            for (int i = _rows; i > 2; i--)
            {
                if (!string.IsNullOrWhiteSpace(_table.Cell(i, 1).Range.Text.Trim('\r', '\a')))
                {
                    _rows = i;
                    break;
                }

                _table.Rows[i].Delete();
            }
        }

        private async Task UpdateCalendarAsync(Task<CalendarUpdateResult> calendarUpdateTask)
        {
            var result = await calendarUpdateTask.ConfigureAwait(false);

            // Updating Holidays in Table
            foreach (var day in result.Holidays)
            {
                var (row, column) = _daysToCells[day];
                _table.Cell(row, column).Range.Font.Color = Word.WdColor.wdColorGray25;
            }

            // Updating Workdays in Table
            foreach (var day in result.Workdays)
            {
                var (row, column) = _daysToCells[day];
                _table.Cell(row, column).Range.Font.Color = Word.WdColor.wdColorBlack;
            }
        }

        private void AddInstructions()
        {
            var selection = _document.Application.Selection;
            selection.MoveDown(Count: _rows);
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
                var table = _document.Tables.Add(selection.Range, 1, 2);
                table.Rows.SetLeftIndent(20, Word.WdRulerStyle.wdAdjustFirstColumn);
                InitializeTable(table);
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
