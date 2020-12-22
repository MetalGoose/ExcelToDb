using System;
using System.Collections.Generic;
using System.Data;
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
            var dbName = "TrainingResults.db";
            var dataSource = $"Data Source={dbName};Cache=Shared";
            var folderName = "files";

            #region Test data

            //var fileName = @"P-ABL-10-F-4A-0001-ULIP2020-AYXX-2102";
            //var fullDataRow = @"ball_block_1.xlsx,1,301balloon.JPG,any,f,2,1,0,0,0,0,0,0,0,0,63.71207350003533,None,None,0,63.71207350003533,None,,P-ABL-10-F-4A-0001-ULIP2020-AYXX-2102,BLAGORODNOVA,001,2020_Dec_04_1156,CMT_balloons_psychopy,2020.2.6,59.98023051910389,";
            //var testFilePath = @"E:\Projects\ExcelToDb\ExcelToDb\TestData\P-ABL-10-F-4A-0001-ULIP2020-AYXX-2102_CMT_balloons_psychopy_2020_Dec_04_1156.csv";
            //var dateString = "2020_Dec_04_1156";

            #endregion Test data

            Directory.CreateDirectory(folderName);
            Console.WriteLine("Поместите все файлы с данными в папку files, которая находится в корневой папке программы и нажмите любую клавишу.");
            Console.ReadKey();
            MainProcess(folderName, dataSource, dbName);
        }

        public static void MainProcess(string folder, string connectionString, string dbName)
        {
            var consoleSpinner = new ConsoleSpinner();
            var dbHandler = new DbHandler(connectionString);

            #region Step 1: find all files

            var pathsToFiles = GetFiles(folder);
            Console.WriteLine($"Найдено файлов: {pathsToFiles.Count}");
            if (pathsToFiles.Count == 0)
            {
                Console.WriteLine($"Файлы не найдены. Поместите файлы данных с расширением .csv в папку {folder}, которая находится в корне программы.");
                return;
            }

            #endregion Step 1: find all files

            #region Step 2: Parse all files into datatables

            Console.Write("Подготовка данных...");
            var tables = new List<DataTable>();
            foreach (var filePath in pathsToFiles)
            {
                consoleSpinner.Turn();
                var preparedNewTable = AddKeyColumnsToDataTable(ConvertCsvToDataTable(filePath));
                tables.Add(preparedNewTable);
            }

            Console.WriteLine();

            #endregion Step 2: Parse all files into datatables

            #region Step 3: Get all persons

            Console.WriteLine("Получение информации о тестируемых...");
            Console.WriteLine();
            var personDict = new Dictionary<int, Person>();
            foreach (var dataTable in tables)
            {
                var persons = GetPersons(dataTable);
                foreach (var person in persons)
                {
                    if (!personDict.ContainsKey(person.Id))
                    {
                        personDict.Add(person.Id, person);
                    }
                }
            }
            Console.WriteLine();
            Console.WriteLine($"Найдено тестируемых: {personDict.Count} ");

            #endregion Step 3: Get all persons

            #region Step 4: Add person id into tables

            Console.WriteLine("Связывание тестируемых и информации о прохождении теста...");
            Console.WriteLine();
            foreach (var dataTable in tables)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    consoleSpinner.Turn();
                    AddPersonInfo(dataTable.Rows[i]);
                }
            }

            Console.WriteLine();
            Console.WriteLine("Связи построены.");

            #endregion Step 4: Add person id into tables

            #region Step 5: Create tables

            Console.WriteLine("Создание таблиц (если не существуют)");

            var personDataTable = personDict.Values.ToDataTable();
            var trainingDataTables = new List<DataTable>();
            var challengeDataTables = new List<DataTable>();

            //Sort challenge types
            foreach (DataTable table in tables)
            {
                if (GetChallengeTableName(table) == "Training")
                {
                    trainingDataTables.Add(table);
                }
                else
                {
                    challengeDataTables.Add(table);
                }
            }

            var personHeaders = new List<string>();
            foreach (DataColumn column in personDataTable.Columns)
            {
                personHeaders.Add(column.ColumnName);
                consoleSpinner.Turn();
            }

            var trainingHeaders = new List<string>();
            foreach (DataColumn column in trainingDataTables[0].Columns)
            {
                trainingHeaders.Add(column.ColumnName);
                consoleSpinner.Turn();
            }

            var challengeHeaders = new List<string>();
            foreach (DataColumn column in challengeDataTables[0].Columns)
            {
                challengeHeaders.Add(column.ColumnName);
                consoleSpinner.Turn();
            }

            dbHandler.CreateTable(personHeaders, "Persons");
            dbHandler.CreateTable(trainingHeaders, "Training");
            dbHandler.CreateTable(challengeHeaders, "Challenge");

            Console.WriteLine("Информация о таблицах обновлена.");

            #endregion Step 5: Create tables

            #region Step 6: Insert data

            Console.WriteLine($"Запись данных в {dbName}");
            Console.WriteLine();
            try
            {
                dbHandler.InsertData("Persons", personDataTable);
            }
            catch (Exception e)
            {
                PrintError(e);
            }
            
            consoleSpinner.Turn();
            foreach (DataTable trainingDataTable in trainingDataTables)
            {
                try
                {
                    dbHandler.InsertData("Training", trainingDataTable);
                    consoleSpinner.Turn();
                }
                catch (Exception e)
                {
                    PrintError(e);
                }
            }
            foreach (DataTable challengeDataTable in challengeDataTables)
            {
                try
                {
                    dbHandler.InsertData("Challenge", challengeDataTable);
                    consoleSpinner.Turn();
                }
                catch (Exception e)
                {
                    PrintError(e);
                }
            }

            Console.WriteLine();
            Console.WriteLine("Данные успешно записаны.");

            #endregion Step 6: Insert data
        }

        /// <summary>
        /// Returns all files with .csv extension.
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        public static List<string> GetFiles(string folder)
        {
            return Directory.GetFiles(folder, "*.csv").ToList();
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
            string personData = null;

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
                Id = int.Parse(personDataList[8]),
                PersonInitials = personDataList[1],
                Age = personDataList[2],
                Gender = personDataList[3],
                Class = personDataList[4],
                SchoolNum = personDataList[5],
                City = personDataList[6].Substring(0, 2)
            };

            return person;
        }

        private static string GetChallengeTableName(DataTable table)
        {
            var tableName = "";
            foreach (DataRow tableRow in table.Rows)
            {
                if (string.IsNullOrEmpty(tableRow["expName"].ToString()))
                {
                    continue;
                }

                tableName = tableRow["expName"].ToString().Contains("TR_") ? "Training" : "Challenge";
                break;
            }

            return tableName;
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
                command.CommandText = "SELECT name FROM userWHERE id = $id";
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
        /// <returns>A list with persons</returns>
        public static List<Person> GetPersons(DataTable dataTable)
        {
            var resultList = new List<Person>();

            foreach (DataRow row in dataTable.Rows)
            {
                try
                {
                    var person = GetPersonInfo(row.DataRowToCsvString());
                    resultList.Add(person);
                }
                catch (Exception e)
                {
                    PrintError(e);
                }
            }

            return resultList;
        }

        private static void PrintError(Exception e)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine($"WARNING: {e.Message}");
            Console.ForegroundColor = ConsoleColor.White;
        }

        public static DataTable AddKeyColumnsToDataTable(DataTable table)
        {
            if (table.Columns.Contains("id"))
            {
                table.Columns["id"].ColumnName = "person_data_key";
            }
            table.Columns.Add("Id", typeof(Guid));
            table.Columns.Add("PersonId", typeof(int));
            return table;
        }

        /// <summary>
        /// Use ref just cuz i can. Nothing more, nothing less
        /// </summary>
        /// <param name="dataRow"></param>
        public static void AddPersonInfo(DataRow dataRow)
        {
            if (dataRow["Id"] != null && dataRow["PersonId"] != null)
            {
                try
                {
                    var person = GetPersonInfo(dataRow.DataRowToCsvString());
                    dataRow["id"] = Guid.NewGuid();
                    dataRow["PersonId"] = person.Id;
                }
                catch (Exception e)
                {
                    PrintError(e);
                }
            }
            else
            {
                throw new NullReferenceException("Row does not contain key columns");
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