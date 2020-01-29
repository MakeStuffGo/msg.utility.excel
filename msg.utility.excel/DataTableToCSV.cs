using System.Data;
using System.IO;
using System.Text;

namespace msg.utility.excel
{
    public static class DataTableToCSV
    {
        public static MemoryStream ToCSV(this DataTable table)
        {
            var result = new StringBuilder();
            for (int i = 0; i < table.Columns.Count; i++)
            {
                result.Append(table.Columns[i].ColumnName);
                result.Append(i == table.Columns.Count - 1 ? "\n" : ",");
            }

            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    result.Append(row[i]);
                    result.Append(i == table.Columns.Count - 1 ? "\n" : ",");
                }
            }

            return GenerateStreamFromString(result.ToString());
        }

        private static MemoryStream GenerateStreamFromString(string s)
        {
            return new MemoryStream(Encoding.UTF8.GetBytes(s ?? ""));
        }
    }
}
