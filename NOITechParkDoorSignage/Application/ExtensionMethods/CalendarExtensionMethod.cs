// SPDX-FileCopyrightText: NOI Techpark <digital@noi.bz.it>
//
// SPDX-License-Identifier: AGPL-3.0-or-later

using Microsoft.Graph.Education.Me.User;
using NOITechParkDoorSignage.Application.Models.REST;
using NOITechParkDoorSignage.Domain.Room;

namespace NOITechParkDoorSignage.Application.ExtensionMethods
{
    public static class CalendarExtensionMethod
    {
        private static TimeZoneInfo italianTimezone => TimeZoneInfo.FindSystemTimeZoneById("Europe/Rome");
        private static DateTime getItalianTimezoneDate(DateTime utcDate)
        {
            DateTime kindDate = new DateTime(utcDate.Year, utcDate.Month, utcDate.Day, utcDate.Hour, utcDate.Minute, utcDate.Second, DateTimeKind.Utc);
            return TimeZoneInfo.ConvertTimeFromUtc(kindDate, italianTimezone);
        }

        public static int TotalMinutesFromNow(this DateTime dateTime)
        {
            var diff = (dateTime - DateTime.Now.ToUniversalTime()).TotalMinutes;
            return diff >= 0 && diff < 1 ? 1 : (int)Math.Round(diff);
        }

        public static CalendarEventViewModel GetEventViewModel(this Event calendarEvent)
        {
            if (calendarEvent == null)
            {
                return null;
            }

            //  At the NOI tech park, the title is compose as follow
            //  [Room],[Organizer],<Title>
            var titleTokens = calendarEvent.Title.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
            var cem = new CalendarEventViewModel
            {
                Title = titleTokens.Length > 3 ? titleTokens[2] : calendarEvent.Title,
                Organizer = calendarEvent.Organizer,
                StartAt = getItalianTimezoneDate(calendarEvent.StartDate).ToString("HH:mm"),
                EndAt = getItalianTimezoneDate(calendarEvent.EndDate).ToString("HH:mm"),
                BookedByLabel = calendarEvent.Organizer == calendarEvent.Room.Email
            };

            if (titleTokens.Length == 3)
            {
                cem.Organizer = titleTokens[1];
                cem.Title = titleTokens[2];
            }
            else if (titleTokens.Length == 2)
            {
                cem.Organizer = titleTokens[0];
                cem.Title = titleTokens[1];
            }
            else
            {
                cem.Title = calendarEvent.Title;
            }

            if (cem.Title.Length > 25)
                cem.Title = cem.Title.Substring(0, 25) + "..";

            return cem;
        }
    }
}