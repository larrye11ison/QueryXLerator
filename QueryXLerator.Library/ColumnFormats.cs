using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QueryXLerator
{
    internal class ColumnFormats
    {
        public readonly static string CurrencyFormat = "$#,##0.00_);($#,##0.00)";
        public readonly static string GeneralNumericFormat = "#,##0.000_);(#,##0.000)";
        public readonly static string PercentFormat = "0.00%";

        private static readonly Func<object, object> byteArrayFormatter =
            (i) =>
            {
                var bytez = i as byte[];
                var stud = bytez
                    .Take(1900)
                    .Aggregate(new StringBuilder("0x"), (x, y) =>
                    {
                        return x.AppendFormat("{0:X2}", y);
                    });

                // truncate it if it's too long
                if (bytez.Length > 2000)
                {
                    stud.Length = 1900;
                    stud.Append("... !!! truncated !!!");
                }
                return stud.ToString();
            };

        private static readonly Dictionary<Type, string> formatMappings = new Dictionary<Type, string>();

        static ColumnFormats()
        {
            formatMappings.Add(typeof(System.Data.SqlTypes.SqlDateTime), "m/d/yyyy");
            formatMappings.Add(typeof(System.DateTime), "m/d/yyyy");
            formatMappings.Add(typeof(System.Data.SqlTypes.SqlDouble), GeneralNumericFormat);
            formatMappings.Add(typeof(System.Data.SqlTypes.SqlDecimal), GeneralNumericFormat);
            formatMappings.Add(typeof(System.Data.SqlTypes.SqlMoney), CurrencyFormat);
        }

        public static ColumnHandler MapTypeToColumnHandler(Type type, Type providerType)
        {
            ColumnHandler rv = new ColumnHandler();

            if (formatMappings.ContainsKey(providerType))
            {
                var formatString = formatMappings[providerType];
                rv.ExcelFormatName = () => formatString;
            }
            if (providerType == typeof(System.Data.SqlTypes.SqlMoney) ||
                providerType == typeof(System.Data.SqlTypes.SqlDecimal) ||
                providerType == typeof(System.Data.SqlTypes.SqlSingle) ||
                providerType == typeof(System.Data.SqlTypes.SqlInt16) ||
                providerType == typeof(System.Data.SqlTypes.SqlInt32) ||
                providerType == typeof(System.Data.SqlTypes.SqlInt64) ||
                providerType == typeof(System.Data.SqlTypes.SqlDouble))
            {
                rv.RowFunction = () => OfficeOpenXml.Table.RowFunctions.Sum;
            }
            if (type == typeof(byte[]))
            {
                rv.Formatter = byteArrayFormatter;
            }
            return rv;
        }
    }
}