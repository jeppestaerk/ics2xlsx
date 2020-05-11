using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using CsvHelper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ical.Net;
using Calendar = Ical.Net.Calendar;

namespace ICalParser
{
    class Program
    {
        static void Main(string[] args)
        {
            var name = "FileName";
            var startTime = new DateTime(2019,1,1);
            var endTime = new DateTime(2020,5,1);
            
            try
            {
                var calEvents = new List<CalEvent>();
                using (StreamReader sr = new StreamReader(name + ".ics"))
                {
                    var calendar = Calendar.LoadFromStream(sr);
                    var events = calendar
                        .GetOccurrences(startTime, endTime)
                        .Select(o => o.Source)
                        .Cast<CalendarEvent>()
                        .Distinct()
                        .ToList();

                    foreach (var calendarEvent in events)
                    {
                        var eventStart = calendarEvent.DtStart.AsSystemLocal;
                        var eventEnd = calendarEvent.DtEnd.AsSystemLocal;
                        var eventSummary = Regex.Replace(calendarEvent.Summary, @"\t|\n|\r", "");
                        calEvents.Add(
                            new CalEvent(
                                eventStart,
                                eventEnd,
                                eventSummary));
                    }

                    calEvents.Sort((x, y) => string.Compare(x.Start.ToString("s"), y.Start.ToString("s"), StringComparison.Ordinal));
                    
                    using (var writer = new StreamWriter(name + ".csv"))
                    using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                    {    
                        csv.WriteRecords(calEvents);
                    }

                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(name + ".xlsx", SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        var sheetData = new SheetData();
                        worksheetPart.Worksheet = new Worksheet(sheetData);

                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet()
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart), 
                            SheetId = 1, 
                            Name = "Sheet1"
                        };

                        sheets.Append(sheet);

                        Row headerRow = new Row();
                        
                        Cell cell1h = new Cell();
                        cell1h.DataType = CellValues.String;
                        cell1h.CellValue = new CellValue("Start");
                        headerRow.Append(cell1h);
                            
                        Cell cell2h = new Cell();
                        cell2h.DataType = CellValues.String;
                        cell2h.CellValue = new CellValue("End");
                        headerRow.Append(cell2h);
                            
                        Cell cell3h = new Cell();
                        cell3h.DataType = CellValues.String;
                        cell3h.CellValue = new CellValue("Text");
                        headerRow.AppendChild(cell3h);
                        
                        sheetData.AppendChild(headerRow);

                        foreach (var calEvent in calEvents)
                        {
                            Row row = new Row();
                            
                            Cell cell1 = new Cell();
                            cell1.DataType = CellValues.String;
                            cell1.CellValue = new CellValue(calEvent.Start.ToString("dd/MM/yyyy HH:mm"));
                            row.AppendChild(cell1);
                            
                            Cell cell2 = new Cell();
                            cell2.DataType = CellValues.String;
                            cell2.CellValue = new CellValue(calEvent.End.ToString("dd/MM/yyyy HH:mm"));
                            row.AppendChild(cell2);
                            
                            Cell cell3 = new Cell();
                            cell3.DataType = CellValues.String;
                            cell3.CellValue = new CellValue(calEvent.Text);
                            row.AppendChild(cell3);

                            sheetData.AppendChild(row);
                        }
                        
                        workbookPart.Workbook.Save();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}