using System;
using CsvHelper.Configuration.Attributes;
using Ical.Net.Interfaces.DataTypes;

namespace ICalParser
{
    public class CalEvent
    {
        [Ignore]
        public DateTime Start { get; set; }
        [Ignore]
        public DateTime End { get; set; }
        public string StartDate { get; set; }
        public string StartTime { get; set; }
        public string EndDate { get; set; }
        public string EndTime { get; set; }
        public string Text { get; set; }

        public CalEvent(DateTime start, DateTime end, string text)
        {
            Start = start;
            End = end;
            StartDate = start.Day + "/" + start.Month + "/" + start.Year;
            StartTime = start.Hour + ":" + start.Minute;
            EndDate = end.Day + "/" + end.Month + "/" + end.Year;
            EndTime = end.Hour + ":" + end.Minute;
            Text = text;
        }
    }
}