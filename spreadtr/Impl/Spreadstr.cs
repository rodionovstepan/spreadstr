namespace spreadstr.Impl
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using OfficeOpenXml;

    public class Spreadstr : ISpreadstr
    {
        private int _rowIndex = 1;

        public Spreadsheet Generate<T>(IEnumerable<T> items, ISpreadstrConverter<T> converter, string title)
        {
            var sheetname = string.Format("Spreadsheet at {0:dd-MM-yyyy}", DateTime.Now);

            using (var excel = new ExcelPackage())
            {
                var worksheet = excel.Workbook.Worksheets.Add(sheetname);

                WriteTitle(title, worksheet);
                WriteItems(items, converter, worksheet);

                using (var stream = new MemoryStream())
                {
                    excel.SaveAs(stream);
                    return new Spreadsheet(stream.ToArray());
                }
            }
        }

        private void WriteTitle(string title, ExcelWorksheet worksheet)
        {
            worksheet.Cells[_rowIndex, 1].Style.Font.Bold = true;
            worksheet.Cells[_rowIndex++, 1].Value = title;
        }

        private void WriteItems<T>(IEnumerable<T> items, ISpreadstrConverter<T> converter, ExcelWorksheet worksheet)
        {
            foreach (var item in items)
            {
                var row = converter.Convert(item);
                var columnIndex = 1;

                foreach (var column in row.Columns)
                {
                    worksheet.Cells[_rowIndex, columnIndex++].Value = column;
                    worksheet.Column(columnIndex).AutoFit();
                }

                ++_rowIndex;
            }
        }
    }
}