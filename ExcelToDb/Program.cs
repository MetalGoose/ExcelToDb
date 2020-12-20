using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace ExcelToDb
{
    internal class Program
    {
        public static void Main()
        {
            var dataSource = "Data Source=TrainingResults.db;Cache=Shared";
            var fileName = @"P-ABL-10-F-4A-0001-ULIP2020-AYXX-2102";
            var fullDataRow = @"ball_block_1.xlsx,1,301balloon.JPG,any,f,2,1,0,0,0,0,0,0,0,0,63.71207350003533,None,None,0,63.71207350003533,None,,P-ABL-10-F-4A-0001-ULIP2020-AYXX-2102,BLAGORODNOVA,001,2020_Dec_04_1156,CMT_balloons_psychopy,2020.2.6,59.98023051910389,";
            var testFilePath = @"E:\Projects\ExcelToDb\ExcelToDb\TestData\P-ABL-10-F-4A-0001-ULIP2020-AYXX-2102_CMT_balloons_psychopy_2020_Dec_04_1156.csv";
            var dateString = "2020_Dec_04_1156";

            //TESTS (kek)

            //var result = GetPersonInfo(fileName);
            //Console.WriteLine(result.ToString());

            var resultTable = ConvertCsvToDataTable(testFilePath);
            //PrintDataTable(resultTable);
            resultTable = AddKeyColumnsToDataTable(resultTable);
            //PrintDataTable(resultTable);

            var dbHandler = new DbHandler(dataSource);

            var headers = new List<string>();
            foreach (DataColumn column in resultTable.Columns)
            {
                headers.Add(column.ColumnName);
            }
            dbHandler.CreateTable(headers, "MainTest");
            dbHandler.InsertData("MainTest", resultTable);
            dbHandler.PrintTable("MainTest");
            Console.ReadKey(true);
            //var persons = GetPersons(resultTable);
            //foreach (var personsValue in persons.Values)
            //{
            //    Console.WriteLine(personsValue.ToString());
            //}
        }

        /// <summary>
        /// Get person info from special string
        /// </summary>
        /// <param name="personRawData">format like - "P-ABL-10-F-4A-0001-ULIP2020-AYXX-2102"</param>
        /// <returns></returns>
        private static Person GetPersonInfo(string personRawData)
        {
            var pattern = @"[A-Z]-[A-Z]{3}-\d*-[A-Z]-.{2,3}-\d{4}-\w{1,}-[A-Z]{4}-\d{1,}";
            var rawDataList = personRawData.Split(',').ToList();
            string personData = "";

            foreach (var value in rawDataList)
            {
                if (Regex.IsMatch(value, pattern))
                {
                    personData = value;
                }
            }

            if (personData is null)
                throw new NullReferenceException("Person data was not found");

            var personDataList = personData.Split('-');

            var person = new Person()
            {
                Id = Guid.NewGuid(),
                PersonInitials = personDataList[1],
                Age = personDataList[2],
                Gender = personDataList[3],
                Class = personDataList[4],
                SchoolNum = personDataList[5],
                City = personDataList[6].Substring(0, 2),
                PersonCode = personDataList[8],
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
            var numberOfMonth = (int)((Utils.Month)Enum.Parse(typeof(Utils.Month), separateDateTimeParts[1]));
            var day = int.Parse(separateDateTimeParts[2]);
            var resultDate = new DateTime(year, numberOfMonth, day);
            return resultDate;
        }

        public static DataTable ConvertCsvToDataTable(string strFilePath)
        {
            using var reader = new StreamReader(strFilePath);
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

        public static void InsertData(string connectionString)
        {
            using (var connection = new SqliteConnection($"Data Source={connectionString}"))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText =
                @"
                    SELECT name
                    FROM user
                    WHERE id = $id
                ";
                command.Parameters.AddWithValue("$id", 3);

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var name = reader.GetString(0);

                        Console.WriteLine($"Hello, {name}!");
                    }
                }
            }
        }

        /// <summary>
        /// Extract person data from raw string
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        public static Dictionary<string, Person> GetPersons(DataTable dataTable)
        {
            var resultDictionary = new Dictionary<string, Person>();

            foreach (DataRow row in dataTable.Rows)
            {
                var person = GetPersonInfo(row.DataRowToCsvString());
                if (!resultDictionary.ContainsKey(person.PersonCode))
                {
                    resultDictionary.Add(person.PersonCode, person);
                }
            }

            return resultDictionary;
        }

        public static DataTable AddKeyColumnsToDataTable(DataTable table)
        {
            if (table.Columns.Contains("id"))
            {
                table.Columns["id"].ColumnName = "person_data_key";
            }
            table.Columns.Add("Id", typeof(Guid));
            table.Columns.Add("PersonId", typeof(Guid));
            return table;
        }

        /// <summary>
        /// Use ref just cuz i can. Nothing more, nothing less)
        /// </summary>
        /// <param name="dataRow"></param>
        public static void AddPersonInfo(ref DataRow dataRow)
        {
            if (dataRow["Id"] != null && dataRow["PersonId"] != null)
            {
                var person = GetPersonInfo(dataRow.DataRowToCsvString());
                dataRow["id"] = Guid.NewGuid();
                dataRow["PersonId"] = person.Id;
            }
            else
            {
                throw new NullReferenceException("DataTable does not contain key columns");
            }
        }

        /// <summary>
        /// Just print rows from DataTable. Nothing more, nothing less.
        /// </summary>
        /// <param name="dataTable"></param>
        public static void PrintDataTable(DataTable dataTable)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    Console.Write(item + "____");
                }

                return;
            }
        }
    }
}