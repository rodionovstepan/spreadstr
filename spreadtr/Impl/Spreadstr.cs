namespace spreadstr.Impl
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using OfficeOpenXml;

    public class Spreadstr : ISpreadstr
    {
        public Spreadsheet Generate<T>(IEnumerable<T> items, ISpreadstrConverter<T> converter)
        {
            var sheetname = string.Format("Spreadsheet at {0:dd-MM-yyyy}", DateTime.Now);

            using (var excel = new ExcelPackage())
            {
                var worksheet = excel.Workbook.Worksheets.Add(sheetname);
                var rowIndex = 1;

                foreach (var item in items)
                {
                    var row = converter.Convert(item);
                    var columnIndex = 1;

                    foreach (var column in row.Columns)
                    {
                        worksheet.Cells[rowIndex, columnIndex++].Value = column;
                        worksheet.Column(columnIndex).AutoFit();
                    }

                    ++rowIndex;
                }

                using (var stream = new MemoryStream())
                {
                    excel.SaveAs(stream);
                    return new Spreadsheet(stream.ToArray());
                }
            }
        }
    }
}