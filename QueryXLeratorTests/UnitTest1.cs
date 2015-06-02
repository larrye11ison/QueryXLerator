using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Linq;
using System.IO;

namespace QueryXLeratorTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void HackerCentral()
        {
            //var templateLoc = @"C:\Users\pwilkins\Documents\SRP-Collections-Pivot-TEMPLATE.xlsx";
            var dest = @"\\eur-sql-001\e$\jobs\GSEReports\Magnetar\WeeklyAdvanceLineData_2014-06-15.xlsx";

            if (System.IO.File.Exists(dest))
            {
            	System.IO.File.Delete(dest);
            }

            QueryXLerator.DataTape.WriteOutputFile(
                dest,
                "exec miser..WeelyAdvanceLineReport 'previous'",
                "server=sql-001; integrated security=sspi;"
                );
            QueryXLerator.DataTape.AddDataToWorksheet(
                dest,
                "exec miser..WeelyAdvanceLineReport 'current'",
                "server=sql-001; integrated security=sspi;",
                "kostas", 
                "curtis");
            //using (var cn = new SqlConnection())
            //{
            //    cn.ConnectionString = "server=eur-sql-stg; integrated security=sspi;";
            //    cn.Open();
            //    using (var cmd = cn.CreateCommand())
            //    {
            //        cmd.CommandText = sql;
            //        cmd.CommandType = System.Data.CommandType.Text;
            //        var p1 = cmd.Parameters.AddWithValue("@backto", new System.DateTime(2013, 05, 01));
            //        var allRecords = cmd.Parameters.AddWithValue("@includeAllRecords", 1);
            //        allRecords.Value = 0;

            //        QueryXLerator.DataTape.AddDataToWorksheet(dest, cmd, "PostLiqTransactions", "postliqdata");
            //        allRecords.Value = 1;
            //        QueryXLerator.DataTape.AddDataToWorksheet(dest, cmd, "Data", "PivotData_1");
            //    }
            //}
        }

        //[TestMethod]
        //public void FindPivotTables()
        //{
        //    using (var pkg = new ExcelPackage(new FileInfo(@"C:\Users\pwilkins\Documents\SRP-Collections-Pivot-2013-07-29.xlsx")))
        //    {
        //        var pivots = pkg.Workbook.Worksheets.SelectMany(ws => ws.PivotTables).ToArray();
        //        //pivots[0].CacheDefinition.SourceRange
        //        System.Diagnostics.Debug.WriteLine(pivots.Length);
        //    }
        //}

        private const string sql =

            @"

--DECLARE @backTo AS DATE = '$backToDate';

WITH srpbase
AS (
	SELECT cast(cast(sr.TransactionDate AS DATE) AS CHAR(7)) AS [Month]
		,datepart(day, sr.TransactionDate) AS [Day]
		,sr.TransactionAmount * - 1 AS TransactionAmount
		,sr.InvestorLoanNumber
		,sr.TransactionDescription
		,sr.TransactionDate
		,sr.ServicerLoanNumber
		,sr.EffectiveDate
		,sr.UPBPostTransaction
		,sr.HistoryCounter
		,ad.LienPosition
		,ad.InvestorID
	FROM miser..VW_ShelvingRock sr
	INNER JOIN miser..asset_detail ad ON ad.loanid = sr.servicerloannumber
	WHERE sr.TransactionDate >= @backTo
	and sr.TransactionDescription != 'New Loan'
	)
	,liqs
AS (
	SELECT b.ServicerLoanNumber
		,b.EffectiveDate
		,b.UPBPostTransaction
		--
		-- This row number allows us to find the earliest ILS transaction where
		-- the UPB was basically zero (LTE to 1). The effective date of that xact
		-- will be used as the liq date.
		--
		,ROW_NUMBER() OVER (
			PARTITION BY b.ServicerLoanNumber ORDER BY b.effectivedate
				,b.historycounter
			) AS FirstLiqRowNum
		,
		--
		-- This row number allows us to find the LAST ILS transaction within the period.
		--
		ROW_NUMBER() OVER (
			PARTITION BY b.ServicerLoanNumber ORDER BY CASE
					-- The SAP transactions all have NULL in this field. None of the ILS
					-- transactions are null. Use this to push the SAP transactions
					-- to the bottom sort-wise (in effect make sure they're ignored).
					WHEN b.UPBPostTransaction IS NULL
						THEN 1
					ELSE 0
					END
				,b.effectivedate DESC
				,b.historycounter DESC
			) AS RowNumLastIlsTrans
	FROM srpbase b
	WHERE b.UPBPostTransaction <= 1
	)
	,dayduh
AS (
	SELECT b.ServicerLoanNumber
		,isnull(cast(cast(l.EffectiveDate AS DATE) AS CHAR(10)), '') AS LiqEffectiveDate
		,b.Day
		,b.Month
		,CASE
			WHEN b.EffectiveDate >= l.EffectiveDate
				THEN 'yes'
			ELSE ''
			END AS IsAfterLiquidation
		,CASE
			WHEN ll.UPBPostTransaction <= 1
				THEN 'yes'
			ELSE ''
			END AS IsInLiqAtEndOfPeriod
		,b.TransactionAmount
		,b.TransactionDescription
		,b.UPBPostTransaction
		,b.HistoryCounter
		,b.LienPosition
		,b.InvestorID
		,b.EffectiveDate
	FROM srpbase b
	LEFT JOIN liqs l ON l.ServicerLoanNumber = b.ServicerLoanNumber
		AND l.FirstLiqRowNum = 1
	-- the loan was still in liq status (based solely on UPB) at the end
	-- of the period.
	LEFT JOIN liqs ll ON ll.ServicerLoanNumber = b.ServicerLoanNumber
		AND ll.RowNumLastIlsTrans = 1
	)
SELECT *
FROM dayduh b
--
-- This script/query is used twice - once to get all records for the pivot table
-- and again to get ONLY the post-liq transactions. The @includeAllRecords
-- part of the WHERE clause allows external callers to specify the mode in which
-- it should operate.
--
WHERE b.IsAfterLiquidation = 'yes' or @includeAllRecords = 1
ORDER BY b.ServicerLoanNumber
	,b.EffectiveDate DESC
	,b.HistoryCounter DESC

";
    }
}