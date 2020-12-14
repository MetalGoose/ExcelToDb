using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelToDb
{
    internal class Program
    {
        public static void Main()
        {
            var fileName = @"P-ABL-10-F-4A-0001-ULIP2020-AYXX-2102";
            var dateString = "2020_Dec_04_1156";
            var result = GetPersonInfo(fileName, dateString);
            Console.WriteLine(result.ToString());
        }

        /// <summary>
        /// Get person info from special string
        /// </summary>
        /// <param name="personRawData">format like - "P-ABL-10-F-4A-0001-ULIP2020-AYXX-2102"</param>
        /// <param name="rawDateTimeString">format like - 2020_Dec_04_1156, where last four nums - HHMM</param>
        /// <returns></returns>
        private static Person GetPersonInfo(string personRawData, string rawDateTimeString)
        {
            var personData = personRawData.Split('-');

            var person = new Person()
            {
                Id = Guid.NewGuid(),
                PersonInitials = personData[1],
                Age = personData[2],
                Gender = personData[3],
                Class = personData[4],
                SchoolNum = personData[5],
                City = personData[6].Substring(0, 2),
                OnlineOrOffline = personData[6].Substring(2, 2),
                FirstTesterInitials = personData[7].Substring(0, 2),
                SecondTesterInitials = personData[7].Substring(2, 2),
                PersonCode = personData[8],
                TestDate = ParseDateTime(rawDateTimeString)
            };

            return person;
        }

        /// <summary>
        /// Get DateTime from special date string
        /// </summary>
        /// <param name="rawDateTimeString">format like - 2020_Dec_04_1156, where last four nums - HHMM</param>
        /// <returns></returns>
        private static DateTime ParseDateTime(string rawDateTimeString)
        {
            string[] separateDateTimeParts = rawDateTimeString.Split('_');
            var year = int.Parse(separateDateTimeParts[0]);
            var numberOfMonth = (int)((Month)Enum.Parse(typeof(Month), separateDateTimeParts[1]));
            var day = int.Parse(separateDateTimeParts[2]);
            var resultDate = new DateTime(year, numberOfMonth, day);
            return resultDate;
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

        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            var reader = new StreamReader(strFilePath);
            string[] headers = reader.ReadLine().Split(',');
            var dataTable = new DataTable();
            foreach (string header in headers)
            {
                dataTable.Columns.Add(header);
            }
            while (!reader.EndOfStream)
            {
                string[] rows = Regex.Split(reader.ReadLine(), ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                var dataRow = dataTable.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dataRow[i] = rows[i];
                }
                dataTable.Rows.Add(dataRow);
            }
            return dataTable;
        }
    }
}