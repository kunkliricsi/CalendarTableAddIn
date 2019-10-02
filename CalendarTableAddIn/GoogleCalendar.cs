using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace CalendarTableAddIn
{
    public class CalendarUpdateResult
    {
        public List<DateTime> Holidays { get; set; } = new List<DateTime>();
        public List<DateTime> Workdays { get; set; } = new List<DateTime>();
    }

    public static class GoogleCalendar
    {
        private const string APPLICATION_NAME = "Calendar Table API";

        public static Task<CalendarUpdateResult> GetWorkdaysAsync(DateTime from, DateTime to)
        {
            return Task.Run(() => GetWorkdays(from, to));
        }

        public static CalendarUpdateResult GetWorkdays(DateTime from, DateTime to)
        {
            var result = new CalendarUpdateResult();

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                ApiKey = Properties.Settings.Default.ApiKey,
                ApplicationName = APPLICATION_NAME,
            });

            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("en.hungarian#holiday@group.v.calendar.google.com");

            // Make TimeMin point to first day of current month.
            request.TimeMin = from;

            // Make TimeMax point to last day of current month.
            request.TimeMax = to;

            // List events.
            Events events = request.Execute();
            if (events.Items?.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    // eg. Extra Work Day for December 15
                    if (eventItem.Summary.Contains("Extra Work Day"))
                    {
                        try
                        {
                            if (eventItem.Start.DateTime.HasValue)
                            {
                                var eventDate = eventItem.Start.DateTime.Value;

                                var summary = eventItem.Summary.Split(' ');

                                var monthName = summary[summary.Length - 1];
                                var month = DateTime.ParseExact(monthName, "MMMM", CultureInfo.CurrentCulture).Month; // This converts December to 12
                                var day = int.Parse(summary.Last()); // This gets the '15' from the summary.
                                var holiday = new DateTime(eventDate.Year, month, day);

                                if (eventDate > holiday)
                                {
                                    holiday.AddYears(1);
                                }

                                result.Holidays.Add(holiday);

                                // Adding the eventDay because its a workday.
                                result.Workdays.Add(eventDate);
                            }
                        }
                        catch { }
                    }
                }
            }

            return result;
        }
    }
}
