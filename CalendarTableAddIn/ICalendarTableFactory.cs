using System;

namespace CalendarTableAddIn
{
    public interface ICalendarTableFactory
    {
        void Create(DateTime month);
    }
}
