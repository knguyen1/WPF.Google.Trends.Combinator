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
        private int _row = 1;

        /// <summary>
        /// Takes two types of parser objects (weekly and daily) and processes them into a combined file.
        /// </summary>
        /// <param name="dailyParser">The daily parser object</param>
        /// <param name="weeklyParser">The weekly parser object</param>
        /// <param name="package">The Excel package object</param>
        /// <param name="fileStream">The FileStream object</param>
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

        /// <summary>
        /// Gets the current row of the Excel client.
        /// </summary>
        /// <returns>The int of the current row.</returns>
        public int GetCurrentRow()
        {
            return _row;
        }

        private void DecreaseRow()
        {
            _row--;
        }

        private void IncreaseRow()
        {
            _row++;
        }

        public void Process()
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

            //form the headers using the enum
            var columns = Enum.GetValues(typeof(ExcelColumns));
            foreach (var column in columns)
            {
                _sheet.Cells[_row, (int)column].Value = column.ToString();
            }

            foreach (var trend in combinedTrends)
            {
                //increment row
                IncreaseRow();

                _sheet.Cells[_row, (int)ExcelColumns.Date].Value = trend.Date;
                _sheet.Cells[_row, (int)ExcelColumns.WeekStart].Value = trend.WeekStart;
                _sheet.Cells[_row, (int)ExcelColumns.WeekEnd].Value = trend.WeekEnd;
                _sheet.Cells[_row, (int)ExcelColumns.DailyIndex].Value = trend.DailyIndex;
                _sheet.Cells[_row, (int)ExcelColumns.WeeklyIndex].Value = trend.WeeklyIndex;

                //re-indexed coefficient formulae
                int curr = _row;
                int prev = _row - 1;
                string coefFormula = String.Format("IF(B{0}=B{1},F{1},E{0}/D{0})", curr, prev);
                string indexFormula = String.Format("D{0}*F{0}", curr);

                //formulae
                _sheet.Cells[_row, (int)ExcelColumns.ReIndexCoeff].Formula = coefFormula;
                _sheet.Cells[_row, (int)ExcelColumns.ReCalcedIndex].Formula = indexFormula;
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

            //_package.SaveAs(_fileStream);
            //Save();

            //UpdateTask(EventArgs.Empty);
        }

        /// <summary>
        /// Saves the package to a filestream.
        /// </summary>
        public void Save()
        {
            _package.SaveAs(_fileStream);
        }

        /// <summary>
        /// Returns the full location of the saved object
        /// </summary>
        /// <returns>Returns the full location of the object</returns>
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
