using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CalendarTableAddIn
{
    public class GoogleCalendarUpdateResult
    {
        public List<DateTime> holidays { get; set; } = new List<DateTime>();
        public List<DateTime> workdays { get; set; } = new List<DateTime>();
    }

    public static class GoogleCalendar
    {
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Calendar Table API";

        public static async Task<GoogleCalendarUpdateResult> UpdateWorkdaysAsync(DateTime from, DateTime to)
        {
            return await Task.Run(() =>
            {
               var result = new GoogleCalendarUpdateResult();

                // Create Google Calendar API service.
                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    ApiKey = Properties.Settings.Default.ApiKey,
                    ApplicationName = ApplicationName,
                });

                // Define parameters of request.
                EventsResource.ListRequest request = service.Events.List("en.hungarian#holiday@group.v.calendar.google.com");

                // Make TimeMin point to first day of current month.
                request.TimeMin = from;

                // Make TimeMax point to last day of current month.
                request.TimeMax = to;

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
                                if (eventItem.Start.DateTime.HasValue)
                                {
                                    var eventDate = eventItem.Start.DateTime.Value;

                                    var summary = eventItem.Summary.Split(' ');

                                    var monthName = summary.ElementAt(summary.Length - 1);
                                    var month = DateTime.ParseExact(monthName, "MMMM", CultureInfo.CurrentCulture).Month; // This converts December to 12
                                    var day = int.Parse(summary.Last()); // This gets the '15' from the summary.
                                    var holiday = new DateTime(eventDate.Year, month, day);

                                    if (eventDate > holiday)
                                    {
                                        holiday.AddYears(1);
                                    }
                                    
                                    result.holidays.Add(holiday);

                                    // Adding the eventDay because its a workday.
                                    result.workdays.Add(eventDate);
                                }
                            }
                            catch { }
                        }
                    }
                }

                return result;
            });
        }
    }
}
