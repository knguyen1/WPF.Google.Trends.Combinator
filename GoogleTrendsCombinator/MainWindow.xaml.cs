using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Microsoft.Win32;
using OfficeOpenXml;
using System.IO;
//using System.WindowsFor
//using System.Windows.Threading;

namespace GoogleTrendsCombinator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //ObservableCollection<string> dailyList = new ObservableCollection<string>();
        //ObservableCollection<string> weeklyList = new ObservableCollection<string>();

        #region Dependency property

        public static readonly DependencyProperty DailyListProperty = DependencyProperty.Register("DailyList", typeof(ObservableCollection<string>), typeof(MainWindow));
        public static readonly DependencyProperty WeeklyListProperty = DependencyProperty.Register("WeeklyList", typeof(ObservableCollection<string>), typeof(MainWindow));
        public static readonly DependencyProperty FileDirCheckBoxProperty = DependencyProperty.Register("IsSameDirBoxChecked", typeof(bool), typeof(MainWindow), new UIPropertyMetadata(true));
        public static readonly DependencyProperty StatusBarTextProperty = DependencyProperty.Register("StatusText", typeof(string), typeof(MainWindow), new UIPropertyMetadata(""));

        public ObservableCollection<string> DailyList
        {
            get { return (ObservableCollection<string>)GetValue(DailyListProperty); }
            set { SetValue(DailyListProperty, value); }
        }

        public ObservableCollection<string> WeeklyList
        {
            get { return (ObservableCollection<string>)GetValue(WeeklyListProperty); }
            set { SetValue(WeeklyListProperty, value); }
        }

        public bool IsSameDirBoxChecked
        {
            get { return (bool)GetValue(FileDirCheckBoxProperty); }
            set { SetValue(FileDirCheckBoxProperty, value); }
        }

        //public string StatusText { get; set; }

        public string StatusText
        {
            get { return (string)GetValue(StatusBarTextProperty); }
            set { SetValue(StatusBarTextProperty, value); }
        }

        #endregion

        private Random random = new Random();

        public MainWindow() //constructor
        {
            DailyList = new ObservableCollection<string>();
            WeeklyList = new ObservableCollection<string>();

            InitializeComponent();

            RefreshUI(null, null);
        }

        private void button12_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;

            ObservableCollection<string> currentList;

            if (button.Name == "btnSelectDaily")
            {
                currentList = DailyList;
            }
            else
            {
                currentList = WeeklyList;
            }

            ////clear the list
            //if (currentList.Count > 0)
            //    currentList.Clear();

            //create open file dialog
            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.Title = "Select Daily Files";

            //set filter for the file extension and the default file extension
            fileDialog.Filter = "CSV Files (*.csv)|*.csv";
            fileDialog.DefaultExt = ".csv";

            //allow multiselect
            fileDialog.Multiselect = true;

            //display the dialog by calling the showdialog() method
            Nullable<bool> result = fileDialog.ShowDialog();

            if (result == true && fileDialog.FileNames.Length > 0)
            {
                //clear the list first
                if (button.Name == "btnSelectDaily")
                    DailyList.Clear();
                else
                    WeeklyList.Clear();

                foreach (var fileName in fileDialog.FileNames)
                {
                    currentList.Add(fileName);
                }
            }

            RefreshUI(null, null);
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            if (DailyList.Count > 0)
                DailyList.Clear();

            if (WeeklyList.Count > 0)
                WeeklyList.Clear();

            RefreshUI(null, null);
        }

        private void RefreshUI(object sender, RoutedEventArgs e)
        {
            btnSubmit.IsEnabled = !((DailyList.Count == 0) || (WeeklyList.Count == 0));
            //progressBar1.SetPercent(0);
            StatusText = string.Empty;

            (MyTabControl.Items[1] as TabItem).IsEnabled = false;
        }

        private void btnSubmit_Click(object sender, RoutedEventArgs e)
        {
            StatusText = "Working...";

            //make a list from ObservableCollection and encapsulate it in the parser
            List<string> dailyFiles = new List<string>(DailyList);
            List<string> weeklyFiles = new List<string>(WeeklyList);
            GoogleTrendsCsvParser dailyParser = new GoogleTrendsCsvParser(dailyFiles);
            GoogleTrendsCsvParser weeklyParser = new GoogleTrendsCsvParser(weeklyFiles);

            string defaultFileName = string.Format("results-{0}-{1}",
                dailyParser.GetSearchTerm(),
                dailyParser.GetTopMostDate());

            string outputDir = null;
            if (IsSameDirBoxChecked)
                outputDir = Path.GetDirectoryName(DailyList[0]);
            else
                outputDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            //form the output filename
            SaveFileDialog saveAs = new SaveFileDialog();
            saveAs.FileName = defaultFileName; //default filename
            saveAs.InitialDirectory = outputDir; //default directory
            saveAs.RestoreDirectory = true;
            saveAs.DefaultExt = ".xlsx";
            saveAs.Filter = "Excel Files (*.xlsx)|*.xlsx";

            string fileName = null;
            Nullable<bool> result = saveAs.ShowDialog();
            if (result == true)
                fileName = saveAs.FileName;
            else
                fileName = String.Format("{0}\\{1}.xlsx", outputDir, defaultFileName);

            //declare the excel package and fill it with data
            using (ExcelPackage package = new ExcelPackage())
            using (FileStream fileStream = new FileStream(fileName, FileMode.Create))
            {
                ExcelClient client = new ExcelClient(dailyParser, weeklyParser, package, fileStream);
                client.ProcessAndSave();
            }

            StatusText = String.Format("File created: {0}", Path.GetFileName(fileName));
        }
    }
}
