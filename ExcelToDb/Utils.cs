using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelToDb
{
    public static class Utils
    {
        /// <summary>
        /// Convert any class into DataTable type
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="self"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IEnumerable<T> self)
        {
            var properties = typeof(T).GetProperties();

            var dataTable = new DataTable();
            foreach (var info in properties)
                dataTable.Columns.Add(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType);

            foreach (var entity in self)
                dataTable.Rows.Add(properties.Select(p => p.GetValue(entity)).ToArray());

            return dataTable;
        }

        public static string DataRowToCsvString(this DataRow dataRow)
        {
            var result = new StringBuilder();
            var rowLength = dataRow.ItemArray.Length;
            for (int i = 0; i < dataRow.ItemArray.Length; i++)
            {
                result.Append(dataRow.ItemArray[i]);
                if (i == rowLength - 1) break;
                result.Append(",");
            }

            return result.ToString();
        }

        public enum Month
        {
            Jan = 1,
            Feb = 2,
            Mar = 3,
            Apr = 4,
            May = 5,
            Jun = 6,
            Jul = 7,
            Aug = 8,
            Sep = 9,
            Oct = 10,
            Nov = 11,
            Dec = 12
        }
    }
}