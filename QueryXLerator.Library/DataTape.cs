using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;

namespace QueryXLerator
{
    public class DataTape
    {
        private const string magicTabNameFieldHeaderColumnNameString = "__tabname__";

        private static readonly Func<string, bool> _IsColumnNameSpecialAndToBeIgnored = columnName =>
        {
            return columnName.ToLower().Contains(magicTabNameFieldHeaderColumnNameString.ToLower());
        };

        private static readonly Regex columnFormatRegex = new Regex(@"\/(?<p>[%\$a-z]+)");
        private static readonly Regex invalidTabNameRegex = new Regex(@"[\[\]\*\/\\\?\:]");

        private static Dictionary<string, RowFunctions> ExcelFuncNames = new Dictionary<string, RowFunctions>();

        public static void AddDataToWorksheet(string xlFilePath, SqlCommand cmd, string worksheetName, string tableName, bool skipEmptyResults = false)
        {
            if (cmd == null)
            {
                throw new ArgumentNullException("cmd");
            }
            ValidateAddDataParameters(xlFilePath, worksheetName, tableName);
            InjectSqlCommandIntoExcelPackage(xlFilePath, cmd, worksheetName, tableName, skipEmptyResults);
        }

        public static void AddDataToWorksheet(string xlFilePath, string commandText, string connectionString, string worksheetName, string tableName, bool skipEmptyResults = false)
        {
            if (string.IsNullOrEmpty(commandText) || commandText.Trim().Length == 0)
            {
                throw new Exception("No commandText was specified.");
            }
            if (string.IsNullOrEmpty(connectionString) || connectionString.Trim().Length == 0)
            {
                throw new Exception("No connectionString was specified.");
            }
            ValidateAddDataParameters(xlFilePath, worksheetName, tableName);

            using (var cn = new SqlConnection(connectionString))
            {
                cn.Open();
                using (var cmd = cn.CreateCommand())
                {
                    cmd.CommandTimeout = 0;
                    cmd.CommandText = commandText;
                    cmd.CommandType = System.Data.CommandType.Text;
                    InjectSqlCommandIntoExcelPackage(xlFilePath, cmd, worksheetName, tableName, skipEmptyResults);
                }
            }

            //using (var pkg = new ExcelPackage(new System.IO.FileInfo(xlFilePath)))
            //{
            //    var worksheets = pkg.Workbook.Worksheets;
            //    var wksheet = worksheets.Where(ws => string.Compare(ws.Name, worksheetName, true) == 0).FirstOrDefault();
            //    if (wksheet != null)
            //    {
            //        worksheets.Delete(wksheet);
            //    }

            //    using (var cn = new SqlConnection(connectionString))
            //    {
            //        cn.Open();
            //        using (var cmd = cn.CreateCommand())
            //        {
            //            cmd.CommandTimeout = 0;
            //            cmd.CommandText = commandText;
            //            cmd.CommandType = System.Data.CommandType.Text;
            //            using (var rdr = cmd.ExecuteReader())
            //            {
            //                WriteWorksheet(pkg.Workbook.Worksheets, worksheetName, _IsColumnNameSpecialAndToBeIgnored, rdr, tableName);
            //            }
            //        }
            //    }
            //    pkg.Save();
            //}
        }

        public static IEnumerable<string> TableStyleNames()
        {
            return Enum.GetNames(typeof(TableStyles))
                .Where(n => n.IndexOf("Custom") == -1)
                .OrderBy(t => t);
        }

        public static void WriteOutputFile(string outputPath, string commandText, string connectionString, bool skipEmptyResults = false, string tableStyleName = "")
        {
            using (var cn = new SqlConnection())
            {
                cn.ConnectionString = connectionString;
                cn.Open();
                using (var cmd = cn.CreateCommand())
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = commandText;
                    cmd.CommandTimeout = 16000;
                    WriteOutputFile(outputPath, cmd, cn, skipEmptyResults, tableStyleName);
                }
            }
        }

        public static void WriteOutputFile(string outputPath, SqlCommand cmd, SqlConnection cn, bool skipEmptyResults = false, string tableStyleName = "")
        {
            if (System.IO.File.Exists(outputPath))
            {
                System.IO.File.Delete(outputPath);
            }

            //_IsColumnNameSpecialAndToBeIgnored =

            using (var pkg = new ExcelPackage(new System.IO.FileInfo(outputPath)))
            {
                var tabNumber = 0;
                {
                    ExcelFuncNames = Enum.GetValues(typeof(RowFunctions))
                        .Cast<RowFunctions>()
                        .ToDictionary(x => Enum.GetName(typeof(RowFunctions), x), x => (RowFunctions)x);

                    using (var rdr = cmd.ExecuteReader())
                    {
                        do
                        {
                            var proposedWorksheetName = String.Format("Result_{0}", tabNumber++);
                            WriteWorksheet(pkg.Workbook.Worksheets, proposedWorksheetName, _IsColumnNameSpecialAndToBeIgnored, rdr, skipEmptyResults, null, tableStyleName);
                        } while (rdr.NextResult());

                        pkg.Save();
                    }
                }
            }
        }

        private static ColumnMetadata GetColumnMetadata(string columnName)
        {
            if (string.IsNullOrEmpty(columnName) || columnName.Trim().Length == 0)
            {
                return new ColumnMetadata { ExcelFormatString = "", Name = " ", RowFunction = RowFunctions.None };
            }
            var proposedColumnName = columnName.Trim();

            var formatMatches = columnFormatRegex.Matches(columnName).Cast<Match>()
                .Where(m => m.Groups.Count == 2)
                .Select(m => m.Groups[1].Value.Trim());
            var theFormatString = "";
            RowFunctions rowFunction = RowFunctions.None;

            // If duplicates exist, last one with win ie.. /sum/average would be avg
            foreach (var match in formatMatches)
            {
                // Special Cases
                switch (match)
                {
                    case "$":
                        theFormatString = ColumnFormats.CurrencyFormat;
                        break;

                    case "%":
                        theFormatString = ColumnFormats.PercentFormat;
                        break;

                    default:
                        break;
                }

                var foo = ExcelFuncNames.Where(x => string.Compare(x.Key, match, true) == 0);
                if (foo.Count() > 0)
                {
                    rowFunction = foo.FirstOrDefault().Value;
                }

                proposedColumnName = columnFormatRegex.Replace(columnName, "");
            }

            // do a bit of cleanup - in case there are some special char's in the rest of the field name
            proposedColumnName = proposedColumnName
                .Replace("#", "Num");

            return new ColumnMetadata
            {
                ExcelFormatString = theFormatString,
                Name = proposedColumnName,
                RowFunction = rowFunction
            };
        }

        private static void InjectSqlCommandIntoExcelPackage(string xlFilePath, SqlCommand cmd, string worksheetName, string tableName, bool skipEmptyResults)
        {
            using (var pkg = new ExcelPackage(new System.IO.FileInfo(xlFilePath)))
            {
                var nmdRange = pkg.Workbook.Names.Where(n => string.Compare(n.Name, tableName, true) == 0).FirstOrDefault();
                if (nmdRange != null)
                {
                    pkg.Workbook.Names.Remove(nmdRange.Name);
                }
                var worksheets = pkg.Workbook.Worksheets;
                var wksheet = worksheets.Where(ws => string.Compare(ws.Name, worksheetName, true) == 0).FirstOrDefault();
                if (wksheet != null)
                {
                    worksheets.Delete(wksheet);
                }
                using (var rdr = cmd.ExecuteReader())
                {
                    WriteWorksheet(pkg.Workbook.Worksheets, worksheetName, _IsColumnNameSpecialAndToBeIgnored, rdr, skipEmptyResults, tableName);
                }
                pkg.Save();
            }
        }

        private static string Uniqueify(IEnumerable<string> previousValues, string proposedValue)
        {
            var proposedReturnValue = proposedValue;
            int uniqueifier = 0;
            while (previousValues.Any(v => string.Compare(v, proposedReturnValue, true) == 0))
            {
                proposedReturnValue = String.Format("{0}_{1}", proposedValue, ++uniqueifier);
            }
            //previousValues.Add(proposedReturnValue.ToLower());
            return proposedReturnValue;
        }

        private static void ValidateAddDataParameters(string xlFilePath, string worksheetName, string tableName)
        {
            if (string.IsNullOrEmpty(worksheetName) || worksheetName.Trim().Length == 0)
            {
                throw new Exception("No worksheetName was specified.");
            }
            if (string.IsNullOrEmpty(tableName) || tableName.Trim().Length == 0)
            {
                throw new Exception("No tableName was specified.");
            }
            if (System.IO.File.Exists(xlFilePath) == false)
            {
                throw new System.IO.FileNotFoundException(String.Format("No Excel file found at '{0}'.", xlFilePath));
            }
        }

        private static void WriteWorksheet(ExcelWorksheets worksheets,
            string proposedWorksheetName,
            Func<string, bool> IsColumnNameSpecialAndToBeIgnored,
            SqlDataReader rdr, bool skipEmptyResults, string tableName = null, string tableStyleName = null)
        {
            if (skipEmptyResults)
            {
                if (rdr.HasRows == false)
                {
                    return;
                }
            }

            var matchingTableStyleName = TableStyleNames()
                .Where(ts => string.Compare(ts, tableStyleName, true) == 0)
                .FirstOrDefault();
            if (matchingTableStyleName == null)
            {
                matchingTableStyleName = "None";
            }
            TableStyles theTableStyle = (TableStyles)System.Enum.Parse(typeof(TableStyles), matchingTableStyleName);

            IEnumerable<string> previousWorksheetNames = worksheets.Select(ws => ws.Name).ToArray();

            // set up the column formats
            Dictionary<int, ColumnHandler> columnHandlers = new Dictionary<int, ColumnHandler>();

            var sheet = worksheets.Add(String.Format("_{0:N}", Guid.NewGuid()));

            var columnMetadata = Enumerable.Range(0, rdr.FieldCount)
                .Select(cc => new
                {
                    ReaderIndex = cc,
                    ColumnMetaData = GetColumnMetadata(rdr.GetName(cc)),
                    ProviderType = rdr.GetProviderSpecificFieldType(cc),
                    Type = rdr.GetFieldType(cc)
                }).ToArray();

            var excelColumnIndex = 1;
            var realColumns = columnMetadata
                .Where(c => IsColumnNameSpecialAndToBeIgnored(c.ColumnMetaData.Name) == false)
                .Select(c => new { Column = c, ExcelIndex = excelColumnIndex++ }) // everyone loves having side-effects in a Linq query!!
                .ToArray();

            var specialColumnForTabName = columnMetadata
                .Where(c => IsColumnNameSpecialAndToBeIgnored(c.ColumnMetaData.Name))
                .FirstOrDefault();

            // keep track of the actual header we used so we don't use it again going forward...
            var columnHeaders = new List<string>();

            ////////////////////////////////////////////////
            // Write column headers
            foreach (var c in realColumns)
            {
                var excelIndex = c.ExcelIndex;

                // make sure the column header is unique, else the Excel table will blow chunks
                var proposedColumnName = Uniqueify(columnHeaders, c.Column.ColumnMetaData.Name);

                // Set the column header
                sheet.Cells[1, excelIndex].Value = proposedColumnName;

                // Set the format for the entire column
                var columnFormat = ColumnFormats.MapTypeToColumnHandler(c.Column.Type, c.Column.ProviderType);
                columnHandlers[excelIndex] = columnFormat;
                ExcelColumn column = sheet.Column(excelIndex);
                column.Style.Numberformat.Format = columnMetadata[c.Column.ReaderIndex].ColumnMetaData.ExcelFormatString == ""
                    ? columnFormat.ExcelFormatName()
                    : columnMetadata[c.Column.ReaderIndex].ColumnMetaData.ExcelFormatString;
                column.Width = 20;
            }

            int excelRowNumber = 2;

            var tabNameSet = false;
            if (specialColumnForTabName != null && specialColumnForTabName.ColumnMetaData.Name.Replace(magicTabNameFieldHeaderColumnNameString, String.Empty) != String.Empty)
            {
                proposedWorksheetName = specialColumnForTabName.ColumnMetaData.Name.Replace(magicTabNameFieldHeaderColumnNameString, String.Empty);
                tabNameSet = true;
            }

            // write out the actual rows of data
            while (rdr.Read())
            {
                // if the query specified the magic column name, use it as the TAB name.
                if (!tabNameSet && specialColumnForTabName != null)
                {
                    proposedWorksheetName = rdr.GetString(specialColumnForTabName.ReaderIndex);
                    tabNameSet = true;
                }

                foreach (var c in realColumns)
                {
                    var columnIndex = c.Column.ReaderIndex;
                    if ((rdr.IsDBNull(columnIndex) == false)
                        && (string.IsNullOrEmpty(rdr[columnIndex].ToString().Trim()) == false))
                    {
                        sheet.Cells[excelRowNumber, c.ExcelIndex].Value =
                            columnHandlers[c.ExcelIndex].Formatter(rdr.GetValue(columnIndex));
                    }
                }
                excelRowNumber++;
            }

            // make sure the tab name is unique and does not contain any of:
            //          [ ] * / \ ? :
            var tempTabName = invalidTabNameRegex.Replace(proposedWorksheetName, "");
            tempTabName = Uniqueify(previousWorksheetNames, tempTabName);
            sheet.Name = tempTabName;

            // set up the "table" in excel - this gives our data some formatting, automatically enables
            // sorting and filtering, plus allows us to easily put summary/totals row at the bottom.
            var tableAddress = new ExcelAddressBase(1, 1, excelRowNumber - 1, realColumns.Count());

            // set up a default tablename
            string newTableName = String.Format("Table_{0}", tempTabName);
            if (string.IsNullOrEmpty(tableName) == false && tableName.Length > 0)
            {
                // but if one was passed in, use that instead.
                newTableName = tableName;
            }
            var tbl = sheet.Tables.Add(tableAddress, newTableName);

            tbl.TableStyle = theTableStyle;

            tbl.ShowTotal = false;

            foreach (var c in tbl.Columns)
            {
                var colMeta = columnMetadata.Where(f => f.ColumnMetaData.Name == c.Name);
                if (colMeta.Count() > 0 && colMeta.FirstOrDefault().ColumnMetaData.RowFunction != RowFunctions.None)
                {
                    c.TotalsRowFunction = colMeta.FirstOrDefault().ColumnMetaData.RowFunction;
                    tbl.ShowTotal = true;
                }

                // TODO: grouping is untested, possibly broken, unsure if it even works. Need to finish impl and test.
                //var groupingMatch = new Regex(@"~(?<groupingLevel>\d)~").Match(c.Name);
                //if (groupingMatch.Groups.Count == 2)
                //{
                //    var groupingLevelString = groupingMatch.Groups[1].Value;
                //    var groupingLevel = int.Parse(groupingLevelString);
                //    sheet.Column(c.Position + 1).OutlineLevel = groupingLevel;
                //}
            }
            sheet.Cells[tableAddress.Address].AutoFitColumns();
        }

        private class ColumnMetadata
        {
            public string ExcelFormatString { get; set; }

            public string Name { get; set; }

            public RowFunctions RowFunction { get; set; }
        }
    }
}