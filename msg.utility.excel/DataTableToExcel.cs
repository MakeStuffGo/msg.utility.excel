using OfficeOpenXml;
using System.Data;
using System.IO;

namespace msg.utility.excel
{

    public static class DataTableToExcel
    {
        public static MemoryStream ToExcel(this DataTable table, string worksheetName)
        {
            var ms = new MemoryStream();

            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(worksheetName);
                ws.Cells["A1"].LoadFromDataTable(table, true);
                pck.Save();
                pck.SaveAs(ms);
            }

            return ms;
        }

        public static MemoryStream ToExcel(this DataTable[] tables, string[] worksheetNames)
        {
            var ms = new MemoryStream();

            using (ExcelPackage pck = new ExcelPackage())
            {
                for (int i = 0; i < tables.Length; i++)
                {
                    var worksheetName = worksheetNames[i];
                    var table = tables[i];

                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add(worksheetName);
                    ws.Cells["A1"].LoadFromDataTable(table, true);
                }

                pck.Save();
                pck.SaveAs(ms);
            }

            return ms;
        }
    }
}
