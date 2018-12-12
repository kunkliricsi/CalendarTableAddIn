using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CalendarTableAddIn
{
    public static class GoogleCalendar
    {
        private static bool initialized = false;
        private static List<int> holidays = new List<int>();
        private static List<int> workdays = new List<int>();
        public static List<int> Holidays
        {
            get
            {
                if (!initialized) throw new Exception("GoogleCalendar has not been initialized");

                return holidays;
            }
        }

        public static List<int> Workdays
        {
            get
            {
                if (!initialized) throw new Exception("GoogleCalendar has not been initialized");

                return workdays;
            }
        }

        public static bool IsWorkday(int day)
        {
            return Workdays.Contains(day);
        }

        public static bool IsHoliday(int day)
        {
            return Holidays.Contains(day);
        }

        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Calendar Table API";

        public static async Task UpdateWorkdaysAsync()
        {
            await Task.Run(() =>
            {
                if (initialized)
                    return;

                // Create Google Calendar API service.
                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    ApiKey = Properties.Settings.Default.ApiKey,
                    ApplicationName = ApplicationName,
                });

                // Define parameters of request.
                EventsResource.ListRequest request = service.Events.List("en.hungarian#holiday@group.v.calendar.google.com");

                // Make TimeMin point to first day of current month.
                var now = DateTime.Now;
                var start = new DateTime(now.Year, now.Month, 1);
                request.TimeMin = start;

                // Make TimeMax point to last day of current month.
                var end = new DateTime(now.Year, now.Month, DateTime.DaysInMonth(now.Year, now.Month));
                request.TimeMax = end;

                // List events.
                Events events = request.Execute();
                if (events.Items != null && events.Items.Count > 0)
                {
                    foreach (var eventItem in events.Items)
                    {
                        // eg. Extra Work Day for December 15
                        if (eventItem.Summary.Contains("Extra Work Day"))
                        {
                            try
                            {
                                string[] summary = eventItem.Summary.Split(' ');

                                // This gets the '15' from the above summary example.
                                holidays.Add(int.Parse(summary.Last()));

                                // This gets the date's day eg. '2018-12-15' to '15'
                                workdays.Add(int.Parse(eventItem.Start.Date.Split('-').Last()));
                            }
                            catch { }
                        }
                    }
                }

                initialized = true;
            });
        }
    }
}
