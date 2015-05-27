# WPF.Google.Trends.Combinator
Accepts weekly and daily Google Trends (csv) files and reindexes the daily trends

Please see blog post at: http://erikjohansson.blogspot.com/2013/04/how-to-get-daily-google-trends-data-for.html

### TL;DR

Google Trends lacks an API that delivers daily trends data for sets more than 90 days.  This program will take in the weekly csv files and the daily csv files and re-indexes the trend data.  You will need to manually download the CSV's and feed the WPF UI.

### Usage

    //from ObservableCollection, you need to make List<T> and pass to the parser
    List<string> dailyFiles = new List<string>(DailyList);
    List<string> weeklyFiles = new List<string>(WeeklyList);
    GoogleTrendsCsvParser dailyParser = new GoogleTrendsCsvParser(dailyFiles);
    GoogleTrendsCsvParser weeklyParser = new GoogleTrendsCsvParser(weeklyFiles);

    //ExcelClient takes the daily/weekly parser, the excel package, and the fileStream
    using (ExcelPackage package = new ExcelPackage())
    using (FileStream fileStream = new FileStream("fileName", FileMode.Create))
    {
        //processing of daily indices are encapsulated inside ExcelClient class
        ExcelClient client = new ExcelClient(dailyParser, weeklyParser, package, fileStream);
        
        //make calculations, reindexes the daily values, then normalize to weekly values
        client.Process();
        
        //make some pretty graphs and charts
        client.AddCharts();
        
        //set sheet active to the charts sheet
        client.SetSheetAsActive(2);
        
        //finally, save the package
        client.Save();
        
        //close and dispose
        client.Dispose();
    }
