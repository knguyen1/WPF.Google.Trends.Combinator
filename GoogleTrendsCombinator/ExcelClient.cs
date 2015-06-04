using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Data;
using OfficeOpenXml.Drawing.Chart;

namespace GoogleTrendsCombinator
{
    public class ExcelClient
    {
        //event class
        //public delegate void EventHandler(object sender, EventArgs e);
        //public event EventHandler TaskUpdate;

        private readonly GoogleTrendsCsvParser _dailyParser;
        private readonly GoogleTrendsCsvParser _weeklyParser;
        private ExcelPackage _package;
        private ExcelWorksheet _dataSheet;
        private ExcelWorksheet _chartSheet;
        private FileStream _fileStream;

        //chart member variables
        private ExcelChart _chart0;
        private ExcelChart _chart1;
        private ExcelChart _chart2;
        private ExcelChart _chart3;

        private Dictionary<string, int> _dailyDict;
        private Dictionary<string, int> _weeklyDict;

        //private DataTable _dataTable;

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
            _dataSheet = _package.Workbook.Worksheets.Add(searchTerm.ToUpper());

            var dailyLines = _dailyParser.GetAllSectionsGrouped("Interest over time");
            var weeklyLines = _weeklyParser.GetAllSectionsAsLines("Interest over time");

            AddToDict(_dailyDict, dailyLines);
            AddToDict(_weeklyDict, weeklyLines);
        }

        private void AddToDict(Dictionary<string, int> dict, IEnumerable<string> lines)
        {
            foreach (string line in lines)
            {
                string[] dateIndex = GetDateAndIndex(line);
                string date = dateIndex[0];
                int index = int.Parse(dateIndex[1]);

                if(!dict.ContainsKey(date))
                    dict.Add(date, index);
            }
        }

        private static string[] GetDateAndIndex(string input)
        {
            return input.Split(new char[] { ',' }, StringSplitOptions.None);
        }

        private void DecreaseRow()
        {
            _row--;
        }

        private void IncreaseRow()
        {
            _row++;
        }

        /// <summary>
        /// Gets the current row of the Excel client.
        /// </summary>
        /// <returns>The int of the current row.</returns>
        public int GetCurrentRow()
        {
            return _row;
        }

        public void SetDailyCsv(GoogleTrendsCsvParser g)
        {
            //clear the dictionary first
            _dailyDict.Clear();

            //now add
            AddToDict(_dailyDict, g.GetAllSectionsAsLines("Interest over time"));
        }

        public void SetWeeklyCsv(GoogleTrendsCsvParser g)
        {
            //clear the dictionary first
            _weeklyDict.Clear();

            //now add
            AddToDict(_weeklyDict, g.GetAllSectionsAsLines("Interest over time"));
        }

        /// <summary>
        /// Processes and makes calculations and then will insert into the Excel package in memory.
        /// </summary>
        public void Process()
        {
            //LINQ
            var combinedTrends = from daily in _dailyDict
                                 from weekly in _weeklyDict
                                 where (DateTime.Parse(daily.Key.Substring(daily.Key.IndexOf("Þ") + 1, 10)) >= DateTime.Parse(weekly.Key.Substring(0, 10)) && DateTime.Parse(daily.Key.Substring(daily.Key.IndexOf("Þ") + 1, 10)) <= DateTime.Parse(weekly.Key.Substring(13, 10)))
                                 select new GoogleTrends
                                 {
                                     Group = int.Parse(daily.Key.Substring(0, daily.Key.IndexOf("Þ"))),
                                     Date = DateTime.Parse(daily.Key.Substring(daily.Key.IndexOf("Þ") + 1, 10)),
                                     WeekStart = DateTime.Parse(weekly.Key.Substring(0, 10)),
                                     WeekEnd = DateTime.Parse(weekly.Key.Substring(13, 10)),
                                     DailyIndex = daily.Value,
                                     WeeklyIndex = weekly.Value
                                 };

            //find the mins and maxes
            var groupedMaxMin = from trend in combinedTrends
                                group trend by trend.Group into g
                                select new
                                {
                                    Group = g.Key,
                                    MaxDailyIndex = g.Max(m => m.DailyIndex),
                                    MinDailyIndex = g.Min(m => m.DailyIndex),
                                    MaxWeeklyIndex = g.Max(m => m.WeeklyIndex),
                                    MinWeeklyIndex = g.Min(m => m.WeeklyIndex)
                                };

            //get the min and maxes by joinging with groupedMaxMin
            var trendsWithMaxMin = from trends in combinedTrends
                                   join maxMin in groupedMaxMin on trends.Group equals maxMin.Group
                                   select new GoogleTrendsWithMaxMin
                                   {
                                       Group = trends.Group,
                                       Date = trends.Date,
                                       WeekStart = trends.WeekStart,
                                       WeekEnd = trends.WeekEnd,
                                       DailyIndex = trends.DailyIndex,
                                       WeeklyIndex = trends.WeeklyIndex,
                                       MaxDailyIndex = maxMin.MaxDailyIndex,
                                       MinDailyIndex = maxMin.MinDailyIndex,
                                       MaxWeeklyIndex = maxMin.MaxWeeklyIndex,
                                       MinWeeklyIndex = maxMin.MinWeeklyIndex
                                   };

            //form the headers using the enum
            var columns = Enum.GetValues(typeof(ExcelColumns));
            foreach (var column in columns)
            {
                _dataSheet.Cells[_row, (int)column].Value = column.ToString();
            }

            foreach (var trend in trendsWithMaxMin.OrderBy(d => d.Date))
            {
                //increment row
                IncreaseRow();

                _dataSheet.Cells[_row, (int)ExcelColumns.Group].Value = trend.Group;
                _dataSheet.Cells[_row, (int)ExcelColumns.Date].Value = trend.Date;
                _dataSheet.Cells[_row, (int)ExcelColumns.WeekStart].Value = trend.WeekStart;
                _dataSheet.Cells[_row, (int)ExcelColumns.WeekEnd].Value = trend.WeekEnd;
                _dataSheet.Cells[_row, (int)ExcelColumns.DailyIndex].Value = trend.DailyIndex;
                _dataSheet.Cells[_row, (int)ExcelColumns.WeeklyIndex].Value = trend.WeeklyIndex;

                //min maxes used in normalization
                _dataSheet.Cells[_row, (int)ExcelColumns.MaxDailyIndex].Value = trend.MaxDailyIndex;
                _dataSheet.Cells[_row, (int)ExcelColumns.MinDailyIndex].Value = trend.MinDailyIndex;
                _dataSheet.Cells[_row, (int)ExcelColumns.MaxWeeklyIndex].Value = trend.MaxWeeklyIndex;
                _dataSheet.Cells[_row, (int)ExcelColumns.MinWeeklyIndex].Value = trend.MinWeeklyIndex;

                //re-indexed coefficient formulae
                int curr = _row;
                int prev = _row - 1;
                string coefFormula = String.Format("IFERROR(IF(C{0}=C{1},G{1},F{0}/E{0}),0)", curr, prev);
                string indexFormula = String.Format("IFERROR(E{0}*G{0},0)", curr);

                //normalizing formulae: (dailyIndex - MinDailyIndex)/(MaxDailyIndex - MinDailyIndex)
                string normFormula = String.Format("((E{0}-J{0})/(I{0}-J{0}))*(K{0}-L{0})+L{0}", curr);

                //formulae
                _dataSheet.Cells[_row, (int)ExcelColumns.ReIndexCoeff].Formula = coefFormula;
                _dataSheet.Cells[_row, (int)ExcelColumns.ReCalcedIndex].Formula = indexFormula;
                _dataSheet.Cells[_row, (int)ExcelColumns.NormalizedIndices].Formula = normFormula;
            }

            //UpdateTask(EventArgs.Empty);

            //calculate the formulae
            _dataSheet.Cells["G:G,H:H,M:M"].Calculate();

            //format the columns

            //decimals
            _dataSheet.Cells["G:G,H:H,M:M"].Style.Numberformat.Format = "0.00";

            //int
            _dataSheet.Cells["A:A,E:E,F:F"].Style.Numberformat.Format = "0";
            _dataSheet.Cells["A:A,E:E,F:F,G:G,H:H,M:M"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            //date
            _dataSheet.Cells["B:B,C:C,D:D"].Style.Numberformat.Format = "MM-dd-yyyy";
            _dataSheet.Cells["B:B,C:C,D:D"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            //freeze top row and autofit
            _dataSheet.Cells[_dataSheet.Dimension.Address].AutoFitColumns();
            _dataSheet.View.FreezePanes(2, 1);
            _dataSheet.Row(1).Style.Font.Bold = true;
        }

        /// <summary>
        /// Adds three four different charts to the "Charts and Graphs" sheet: Normalized Daily Index, Weekly Index raw, etc...
        /// </summary>
        public void AddCharts()
        {
            //add chart to a separate sheet
            _chartSheet = _package.Workbook.Worksheets.Add("Charts and Graphs");

            //get the last possible row
            int lastRow = _dataSheet.Dimension.End.Row;

            //format the ExcelBaseRange addresses to be used in forming r2 and r1
            string dateDimension = String.Format("B2:B{0}", lastRow.ToString());
            string normalizedIndex = String.Format("M2:M{0}", lastRow.ToString());
            string unNormalizedIndex = String.Format("H2:H{0}", lastRow.ToString());
            string weeklyIndex = String.Format("F2:F{0}", lastRow.ToString());
            string dailyIndex = String.Format("E2:E{0}", lastRow.ToString());

            //get r1 and r2 from the _dataSheet cells so we can add it to the separated sheet
            var dateX = _dataSheet.Cells[dateDimension];
            var normalizedY = _dataSheet.Cells[normalizedIndex];
            var unNormalizedY = _dataSheet.Cells[unNormalizedIndex];
            var dailyY = _dataSheet.Cells[dailyIndex];
            var weeklyY = _dataSheet.Cells[weeklyIndex];

            //add and format chart on the _chartSheet object
            _chart0 = _chartSheet.Drawings.AddChart("Chart1", eChartType.Line);
            _chart0.Title.Text = "Normalized Daily Trends";
            _chart0.Title.Font.Bold = true;
            _chart0.SetPosition(2, 5, 1, 5);
            _chart0.SetSize(550, 230);
            _chart0.Series.Add(normalizedY, dateX);
            _chart0.Series[0].Header = "NormalizedIndices";
            _chart0.Style = eChartStyle.Style4;
            _chart0.XAxis.Title.Text = "Date (Daily)";
            _chart0.YAxis.Title.Text = "Trend Index";
            _chart0.XAxis.Title.Font.Size = 10;
            _chart0.YAxis.Title.Font.Size = 10;

            _chart1 = _chartSheet.Drawings.AddChart("Chart4", eChartType.Line);
            _chart1.Title.Text = "Un-Normalized Daily Trends";
            _chart1.Title.Font.Bold = true;
            _chart1.SetPosition(2, 5, 10, 5); //row, row pixel offset, column, column pixeloffset
            _chart1.SetSize(550, 230);
            _chart1.Series.Add(unNormalizedY, dateX);
            _chart1.Series[0].Header = "ReCalcedIndex";
            _chart1.Style = eChartStyle.Style5;
            _chart1.XAxis.Title.Text = "Date (Daily)";
            _chart1.YAxis.Title.Text = "Trend Index";
            _chart1.XAxis.Title.Font.Size = 10;
            _chart1.YAxis.Title.Font.Size = 10;

            _chart2 = _chartSheet.Drawings.AddChart("Chart2", eChartType.Line);
            _chart2.Title.Text = "Daily Trends (Raw Data)";
            _chart2.Title.Font.Bold = true;
            _chart2.SetPosition(15, 5, 10, 5);
            _chart2.SetSize(550, 230);
            _chart2.Series.Add(dailyY, dateX);
            _chart2.Series[0].Header = "DailyIndex";
            _chart2.Style = eChartStyle.Style3;
            _chart2.XAxis.Title.Text = "Date (Daily)";
            _chart2.YAxis.Title.Text = "Trend Index";
            _chart2.XAxis.Title.Font.Size = 10;
            _chart2.YAxis.Title.Font.Size = 10;

            _chart3 = _chartSheet.Drawings.AddChart("Chart3", eChartType.Line);
            _chart3.Title.Text = "Weekly Trends (Raw Data)";
            _chart3.Title.Font.Bold = true;
            _chart3.SetPosition(15, 5, 1, 5);
            _chart3.SetSize(550, 230);
            _chart3.Series.Add(weeklyY, dateX);
            _chart3.Series[0].Header = "WeeklyIndex";
            _chart3.Style = eChartStyle.Style3;
            _chart3.XAxis.Title.Text = "Date (Weekly)";
            _chart3.YAxis.Title.Text = "Trend Index";
            _chart3.XAxis.Title.Font.Size = 10;
            _chart3.YAxis.Title.Font.Size = 10;
        }

        /// <summary>
        /// Saves the package to a filestream.
        /// </summary>
        public void Save()
        {
            try
            {
                _package.SaveAs(_fileStream);
            }
            catch (InvalidOperationException exc)
            {
                throw (exc);
            }
            catch (Exception exc)
            {
                throw (exc);
            }
        }

        public void Dispose()
        {
            //dispose of the charts
            _chart0.Dispose();
            _chart1.Dispose();
            _chart2.Dispose();
            _chart3.Dispose();

            //dispose of the package
            _package.Dispose();
        }

        /// <summary>
        /// Returns the full location of the saved object
        /// </summary>
        /// <returns>Returns the full location of the object</returns>
        public string GetSavedFileLocation()
        {
            return _package.File.FullName;
        }

        /// <summary>
        /// Sets the selected worksheet as default active when opened.
        /// </summary>
        /// <param name="index">The index position of the sheet.  1-based.</param>
        public void SetSheetAsActive(int index)
        {
            _package.Workbook.Worksheets[index].View.TabSelected = true;
        }

        //protected void UpdateTask(EventArgs e)
        //{
        //    EventHandler handler = TaskUpdate;
        //    if (handler != null)
        //        handler(this, e);
        //}
    }
}
