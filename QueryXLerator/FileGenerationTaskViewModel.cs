using System;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

namespace QueryXLerator
{
    public class FileGenerationTaskViewModel : ViewModelBase
    {
        private CancellationTokenSource cancelTokenSource = null;

        public bool CanCancel { get; private set; }

        public string Description { get; set; }

        public TimeSpan Duration { get; set; }

        public string DurationString
        {
            get
            {
                return Duration.ToString(@"hh\:mm\:ss");
            }
        }

        public string FileName { get; set; }

        public bool IncludeBlankResults { get; set; }

        public bool IsInErrorState { get; set; }

        public bool IsTaskComplete { get; set; }

        public string SqlQuery { get; set; }

        public DateTime Started { get; set; }

        public string Status { get; set; }

        public void Cancel()
        {
            cancelTokenSource.Cancel();
        }

        public async Task Run(string finalOutputPath, string queryText, string cnString, string fileName, bool includeEmptyResultsInExcelFile, string tableStyleName)
        {
            IsTaskComplete = false;
            SqlQuery = queryText;
            FileName = fileName;
            IncludeBlankResults = includeEmptyResultsInExcelFile;

            System.Timers.Timer t = new System.Timers.Timer(1000);
            try
            {
                Status = "Running...";
                Started = DateTime.Now;
                t.Elapsed += t_Elapsed;
                t.Start();
                cancelTokenSource = new CancellationTokenSource();
                CanCancel = true;
                await Task.Run(() =>
                {
                    DataTape.WriteOutputFile(finalOutputPath, queryText, cnString, includeEmptyResultsInExcelFile, tableStyleName, cancelTokenSource.Token);
                });
                Status = "Complete.";
            }
            catch (Exception ex)
            {
                if (cancelTokenSource.IsCancellationRequested)
                {
                    Status = "Cancelled by user.";
                }
                else
                {
                    IsInErrorState = true;
                    Status = String.Format("FAIL: {0}", ex);
                }
            }
            finally
            {
                CanCancel = false;
                t.Dispose();
                IsTaskComplete = true;
            }
        }

        private void t_Elapsed(object sender, ElapsedEventArgs e)
        {
            Duration = DateTime.Now.Subtract(Started);
            RaisePropChanged("DurationString");
        }
    }
}