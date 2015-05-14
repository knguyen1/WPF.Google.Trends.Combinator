using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;

namespace GoogleTrendsCombinator
{
    public class ExcelClient
    {
        //event class
        //public delegate void EventHandler(object sender, EventArgs e);
        public event EventHandler TaskUpdate;

        private readonly GoogleTrendsCsvParser _dailyParser;
        private readonly GoogleTrendsCsvParser _weeklyParser;
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;
        private FileStream _fileStream;

        private Dictionary<string, int> _dailyDict;
        private Dictionary<string, int> _weeklyDict;

        //parser variables
        private int row = 1;
        private int column = 1;

        public ExcelClient(GoogleTrendsCsvParser dailyParser, GoogleTrendsCsvParser weeklyParser, ExcelPackage package, FileStream fileStream)
        {
            if (dailyParser == null)
                throw new ArgumentNullException("Daily files cannot be null!");

            if (weeklyParser == null)
                throw new ArgumentNullException("Weekly files cannot be null!");

            if (package == null)
                _package = new ExcelPackage(new FileInfo(Settings1.Default.defaultExcel));

            //constructor arguments
            _fileStream = fileStream;
            _dailyParser = dailyParser;
            _weeklyParser = weeklyParser;
            _package = package;
            _dailyDict = new Dictionary<string, int>();
            _weeklyDict = new Dictionary<string, int>();

            //create a new sheet based on the search term fed
            string searchTerm = _dailyParser.GetSearchTerm();
            _sheet = _package.Workbook.Worksheets.Add(searchTerm.ToUpper());

            var dailyLines = _dailyParser.GetAllSectionsAsLines("Interest over time");
            var weeklyLines = _weeklyParser.GetAllSectionsAsLines("Interest over time");

            //form the weekly index dictionary
            foreach (var line in weeklyLines)
            {
                string[] dateIndex = GetDateAndIndex(line);
                string dates = dateIndex[0];
                int index = int.Parse(dateIndex[1]);

                //add to dictionary
                _weeklyDict.Add(dates, index);
            }

            //form the daily index dictionary
            foreach (var line in dailyLines)
            {
                string[] dateIndex = GetDateAndIndex(line);
                string date = dateIndex[0];
                int index = int.Parse(dateIndex[1]);

                //add to daily dictionary
                _dailyDict.Add(date, index);
            }
        }

        public void ProcessAndSave()
        {
            var combinedTrends = (from daily in _dailyDict
                                  from weekly in _weeklyDict
                                  where (DateTime.Parse(daily.Key) >= DateTime.Parse(weekly.Key.Substring(0, 10)) && DateTime.Parse(daily.Key) <= DateTime.Parse(weekly.Key.Substring(13, 10)))
                                  select new GoogleTrends
                                  {
                                      Date = DateTime.Parse(daily.Key),
                                      WeekStart = DateTime.Parse(weekly.Key.Substring(0, 10)),
                                      WeekEnd = DateTime.Parse(weekly.Key.Substring(13, 10)),
                                      DailyIndex = daily.Value,
                                      WeeklyIndex = weekly.Value
                                  }).OrderBy(x => x.Date);

            //UpdateTask(EventArgs.Empty);

            //headers
            _sheet.Cells[row, column].Value = "Date";
            _sheet.Cells[row, column + 1].Value = "WeekStart";
            _sheet.Cells[row, column + 2].Value = "WeekEnd";
            _sheet.Cells[row, column + 3].Value = "DailyIndex";
            _sheet.Cells[row, column + 4].Value = "WeeklyIndex";
            _sheet.Cells[row, column + 5].Value = "ReIndexedCoeff";
            _sheet.Cells[row, column + 6].Value = "ReCalcedIndex";

            var props = typeof(GoogleTrends).GetProperties();

            //UpdateTask(EventArgs.Empty);

            foreach (var trend in combinedTrends)
            {
                //increment row
                row++;

                _sheet.Cells[row, column].Value = trend.Date;
                _sheet.Cells[row, column + 1].Value = trend.WeekStart;
                _sheet.Cells[row, column + 2].Value = trend.WeekEnd;
                _sheet.Cells[row, column + 3].Value = trend.DailyIndex;
                _sheet.Cells[row, column + 4].Value = trend.WeeklyIndex;

                //re-indexed coefficient formulae
                int curr = row;
                int prev = row - 1;
                string coefFormula = String.Format("IF(B{0}=B{1},F{1},E{0}/D{0})", curr, prev);
                string indexFormula = String.Format("D{0}*F{0}", curr);

                //formulae
                _sheet.Cells[row, column + 5].Formula = coefFormula;
                _sheet.Cells[row, column + 6].Formula = indexFormula;
            }

            //UpdateTask(EventArgs.Empty);

            //calculate the formulae
            _sheet.Cells["F:F,G:G"].Calculate();

            //format the columns

            //decimals
            _sheet.Cells["F:F,G:G"].Style.Numberformat.Format = "0.00000";

            //int
            _sheet.Cells["D:D,E:E"].Style.Numberformat.Format = "0";
            _sheet.Cells["D:D,E:E,F:F,G:G"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            //date
            _sheet.Cells["A:A,B:B,C:C"].Style.Numberformat.Format = "MM-dd-yyyy";
            _sheet.Cells["A:A,B:B,C:C"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            //freeze top row and autofit
            _sheet.Cells[_sheet.Dimension.Address].AutoFitColumns();
            _sheet.View.FreezePanes(2, 1);
            _sheet.Row(1).Style.Font.Bold = true;

            _package.SaveAs(_fileStream);

            //UpdateTask(EventArgs.Empty);
        }

        public string GetSavedFileLocation()
        {
            return _package.File.FullName;
        }

        private static string[] GetDateAndIndex(string input)
        {
            return input.Split(new char[] { ',' }, StringSplitOptions.None);
        }

        protected void UpdateTask(EventArgs e)
        {
            EventHandler handler = TaskUpdate;
            if (handler != null)
                handler(this, e);
        }
    }
}
