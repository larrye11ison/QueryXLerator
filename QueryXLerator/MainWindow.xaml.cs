using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml;

namespace QueryXLerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            RunningTasksX = new ObservableCollection<FileGenerationTaskViewModel>();
            CompletedTasksX = new ObservableCollection<FileGenerationTaskViewModel>();
            OutputPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        public ObservableCollection<FileGenerationTaskViewModel> CompletedTasksX { get; set; }

        public string OutputPath { get; set; }

        public ObservableCollection<FileGenerationTaskViewModel> RunningTasksX { get; set; }

        public IEnumerable<ExcelTableStyle> TableStyleNames
        {
            get
            {
                return DataTape.TableStyleNames()
                    .Select(n => new ExcelTableStyle
                    {
                        Name = n,
                        ImageSource = ExcelTableStyleSampleImages.GetImageForStyle(n)
                    })
                    .OrderBy(ts =>
                        {
                            if (ts.Name == "None")
                            {
                                return "00000000000";
                            }
                            else
                            {
                                return ts.Name;
                            }
                        });
            }
        }

        private void FormatQueryButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                queryText.Text = PoorMansTSqlFormatterLib.SqlFormattingManager.DefaultFormat(queryText.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("FAIL: " + ex.ToString(), "Error Occurred");
            }
        }

        private void LoadQueryText(object sender, RoutedEventArgs e)
        {
            var tb = sender as Button;
            var fgt = tb.DataContext as FileGenerationTaskViewModel;
            this.queryText.Text = fgt.SqlQuery;
            this.outputFileNameTextBox.Text = fgt.FileName;
            if (fgt.IncludeBlankResults)
            {
                NoExcludeEmptyResultsFromExcelFile.IsChecked = true;
            }
            else
            {
                YesIncludeEmptyResultsInExcelFile.IsChecked = true;
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.DataContext = this;
            var mainWindowType = typeof(MainWindow);
            var fullResourceName = mainWindowType.Namespace + ".tsql.xshd";

            IHighlightingDefinition hl;
            using (var stream = mainWindowType.Assembly.GetManifestResourceStream(fullResourceName))
            {
                using (var reader = new XmlTextReader(stream))
                {
                    hl = HighlightingLoader.Load(reader, HighlightingManager.Instance);
                }
            }

            RunningTasks.ItemsSource = RunningTasksX;
            CompletedTasks.ItemsSource = CompletedTasksX;

            this.queryText.SyntaxHighlighting = hl;
            queryText.Text = @"-- Just to demonstrate async exec of queries...
WAITFOR DELAY '00:00:03';

SELECT 'WorksheetTabName' AS __tabname__ -- sets the tab name in Excel
	,database_id AS [Total of database id/sum] -- will insert “sum” function in summary row
    -- Any excel function can be used here, just use a fwd slash followed by the function
	-- name. The rest of the content will be the column header.
    ,NAME -- just a regular column
FROM master.sys.databases;
";
        }

        private void RemoveCompletedTask(object sender, RoutedEventArgs e)
        {
            var tb = sender as Button;
            var fgt = tb.DataContext as FileGenerationTaskViewModel;
            CompletedTasksX.Remove(fgt);
        }
        private void CancelRunningTask(object sender, RoutedEventArgs e)
        {
            var tb = sender as Button;
            var fgt = tb.DataContext as FileGenerationTaskViewModel;
            fgt.Cancel();
        }

        private async void RunQueryButton_Click(object sender, RoutedEventArgs e)
        {
            string outputFileName = outputFileNameTextBox.Text;
            if (string.IsNullOrWhiteSpace(outputFileName))
            {
                throw new ArgumentException("Output File Name must have a value.");
            }
            char[] invalidPathAndFileCharacters = System.IO.Path.GetInvalidPathChars().Union(System.IO.Path.GetInvalidFileNameChars()).ToArray();
            if (outputFileName.IndexOfAny(invalidPathAndFileCharacters) > -1)
            {
                throw new ArgumentException("Output File Name contains one or more invalid characters.");
            }
            if (outputFileName.EndsWith(".xlsx", StringComparison.CurrentCultureIgnoreCase) == false)
            {
                outputFileName += ".xlsx";
            }
            var myDocsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string finalOutputPath = System.IO.Path.Combine(myDocsPath, outputFileName);

            if (System.IO.File.Exists(finalOutputPath))
            {
                var mbr = MessageBox.Show("Output file already exists - overwrite?", "Overwrite file?", MessageBoxButton.YesNo);
                if (mbr == MessageBoxResult.No)
                {
                    return;
                }
            }
            var t = new FileGenerationTaskViewModel();
            t.Description = outputFileName;
            RunningTasksX.Insert(0, t);
            var includeEmptyResults = NoExcludeEmptyResultsFromExcelFile.IsChecked;

            var selectedTableStyleSelectedValue = this.SelectedTableStyle.SelectedValue as ExcelTableStyle;
            var tableStyle = selectedTableStyleSelectedValue == null ? "" : selectedTableStyleSelectedValue.Name;

            await t.Run(finalOutputPath, queryText.Text, connectionStringTextBox.Text, outputFileNameTextBox.Text, includeEmptyResults.Value, tableStyle);
            RunningTasksX.Remove(t);
            CompletedTasksX.Insert(0, t);
        }

        public class ExcelTableStyle
        {
            public ImageSource ImageSource { get; set; }

            public string Name { get; set; }
        }

        private void OpenDocumentsFolder(object sender, RoutedEventArgs e)
        {
            var butt = sender as Button;
            System.Diagnostics.Process.Start(butt.Content.ToString());
        }
    }
}