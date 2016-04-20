using Fluent;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Xml;
using cunt = System.Windows.Controls;

namespace QueryXLerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
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

        private void CancelRunningTask(object sender, RoutedEventArgs e)
        {
            var tb = sender as cunt.Button;
            var fgt = tb.DataContext as FileGenerationTaskViewModel;
            fgt.Cancel();
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
            var tb = sender as cunt.Button;
            var fgt = tb.DataContext as FileGenerationTaskViewModel;
            this.queryText.Text = fgt.SqlQuery;
            this.outputFileNameTextBox.Text = fgt.FileName;
            if (fgt.IncludeBlankResults)
            {
                includeEmptyResultsetsInExcelOutputFile.IsChecked = true;
                //NoExcludeEmptyResultsFromExcelFile.IsChecked = true;
            }
            else
            {
                includeEmptyResultsetsInExcelOutputFile.IsChecked = false;
                //YesIncludeEmptyResultsInExcelFile.IsChecked = true;
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

            //RunningTasks.ItemsSource = RunningTasksX;
            //CompletedTasks.ItemsSource = CompletedTasksX;

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

        private void OpenDocumentsFolder(object sender, RoutedEventArgs e)
        {
            var butt = sender as cunt.Button;
            System.Diagnostics.Process.Start(butt.Content.ToString());
        }

        private void RemoveCompletedTask(object sender, RoutedEventArgs e)
        {
            var tb = sender as cunt.Button;
            var fgt = tb.DataContext as FileGenerationTaskViewModel;
            CompletedTasksX.Remove(fgt);
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
                string messageText = $"The specified output file already exists:{Environment.NewLine}{finalOutputPath};";
                var result = await this.ShowMessageAsync("Overwrite file?", messageText, MessageDialogStyle.AffirmativeAndNegative, new MetroDialogSettings
                {
                    AffirmativeButtonText = "Continue and Overwrite the file",
                    NegativeButtonText = "Cancel entire operation"
                });

                if (result == MessageDialogResult.Negative)
                {
                    return;
                }
            }
            var t = new FileGenerationTaskViewModel();
            t.Description = outputFileName;
            RunningTasksX.Insert(0, t);
            var includeEmptyResults = includeEmptyResultsetsInExcelOutputFile.IsChecked;
            //NoExcludeEmptyResultsFromExcelFile.IsChecked;

            var selectedTableStyleSelectedValue = SelectedTableStyleGallery.SelectedValue as ExcelTableStyle;
                //this.SelectedTableStyle.SelectedValue as ExcelTableStyle;
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
    }
}