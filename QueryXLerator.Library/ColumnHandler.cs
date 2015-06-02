using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace QueryXLerator
{
    /// <summary>
    ///
    /// </summary>
    public class ColumnHandler
    {
        /// <summary>
        /// The name (actually, the literal excel format string) used to
        /// format the data in the destination excel document.
        /// </summary>
        public Func<string> ExcelFormatName = () => "General";

        /// <summary>
        /// By default, the Identity function (returns exactly what was input).
        /// But can be overridden to perform any special processing that may be necessary.
        /// </summary>
        public Func<object, object> Formatter = (i) => { return i; };

        public Func<RowFunctions> RowFunction = () => RowFunctions.None;
    }
}