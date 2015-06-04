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

        //private List<int> numberOfSearchTerms;

        #region Dependency property

        public static readonly DependencyProperty DailyListProperty = DependencyProperty.Register("DailyList", typeof(ObservableCollection<string>), typeof(MainWindow));
        public static readonly DependencyProperty WeeklyListProperty = DependencyProperty.Register("WeeklyList", typeof(ObservableCollection<string>), typeof(MainWindow));
        public static readonly DependencyProperty PromptWhereToSaveProperty = DependencyProperty.Register("IsSameDirBoxChecked", typeof(bool), typeof(MainWindow), new UIPropertyMetadata(true));
        public static readonly DependencyProperty StatusBarTextProperty = DependencyProperty.Register("StatusText", typeof(string), typeof(MainWindow), new UIPropertyMetadata(String.Empty));
        public static readonly DependencyProperty MakeChartsCheckBoxProperty = DependencyProperty.Register("IsMakeChartsChecked", typeof(bool), typeof(MainWindow), new UIPropertyMetadata(true));

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

        public bool AskWhereToSave
        {
            get { return (bool)GetValue(PromptWhereToSaveProperty); }
            set { SetValue(PromptWhereToSaveProperty, value); }
        }

        public bool IsMakeChartsChecked
        {
            get { return (bool)GetValue(MakeChartsCheckBoxProperty); }
            set { SetValue(MakeChartsCheckBoxProperty, value); }
        }

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

            //numberOfSearchTerms = new List<int>();
            //numberOfSearchTerms.Add(1);
            //numberOfSearchTerms.Add(2);
            //numberOfSearchTerms.Add(3);
            //numberOfSearchTerms.Add(4);
            //numberOfSearchTerms.Add(5);

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

            string defaultFileName = String.Format("results-{0}-{1}",
                dailyParser.GetSearchTerm(),
                dailyParser.GetTopMostDate());

            //string outputDir = null;
            //if (AskWhereToSave)
            //    outputDir = Path.GetDirectoryName(DailyList[0]);
            //else
            //    outputDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string outputDir = Path.GetDirectoryName(DailyList[0]);
            string fileName = null;

            if (AskWhereToSave)
            {
                //form the output filename
                SaveFileDialog saveAs = new SaveFileDialog();
                saveAs.FileName = defaultFileName; //default filename
                saveAs.InitialDirectory = outputDir; //default directory
                saveAs.RestoreDirectory = true;
                saveAs.DefaultExt = ".xlsx";
                saveAs.Filter = "Excel Files (*.xlsx)|*.xlsx";

                Nullable<bool> result = saveAs.ShowDialog();
                if (result == true)
                    fileName = saveAs.FileName;
                else
                {
                    StatusText = "You pushed canceled, so I didn't do anything. :)";

                    return;
                }
            }
            else
            {
                fileName = String.Format("{0}\\{1}.xlsx", outputDir, defaultFileName);
            }

            try
            {
                //declare the excel package and fill it with data
                using (ExcelPackage package = new ExcelPackage())
                using (FileStream fileStream = new FileStream(fileName, FileMode.Create))
                {
                    ExcelClient client = new ExcelClient(dailyParser, weeklyParser, package, fileStream);

                    client.Process(); //normalize and recalculate all indices

                    if (IsMakeChartsChecked)
                    {
                        client.AddCharts();
                        client.SetSheetAsActive(2); //set the chart sheet as active
                    }

                    client.Save(); //save the sheet
                    client.Dispose(); //close and dispose the package
                }

                StatusText = String.Format("File created: {0}", Path.GetFileName(fileName));
            }
            catch (IOException ioExc)
            {
                StatusText = "File in use error: " + ioExc.Message;
            }
            catch (Exception exc)
            {
                StatusText = "Error: " + exc.Message;
            }
        }
    }
}
