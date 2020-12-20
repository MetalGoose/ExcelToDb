using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Microsoft.Data.Sqlite;

namespace ExcelToDb
{
    public class DbHandler
    {
        public string ConnectionString { get; }

        public DbHandler(string baseConnectionString)
        {
            ConnectionString = new SqliteConnectionStringBuilder(baseConnectionString) { Mode = SqliteOpenMode.ReadWriteCreate }.ToString();
        }

        public void CreateTable(List<string> columnNames, string tableName)
        {
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();

            var columns = new StringBuilder();

            columnNames.ForEach(columnName =>
            {
                columns.Append(columnName == "Id" ? "Id TEXT PRIMARY KEY," : $@"{columnName.Replace('.', '_')} TEXT,");
            });

            columns.Remove(columns.Length - 1, 1);
            var resultColumnNames = columns.ToString();

            var commandString = $@"CREATE TABLE IF NOT EXISTS {tableName}({resultColumnNames})";

            using var initCommand = new SqliteCommand(commandString, connection);

            initCommand.ExecuteNonQuery();
        }

        public void PrintTable(string tableName)
        {
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            using var initCommand = new SqliteCommand($"SELECT * FROM {tableName}", connection);
            using var reader = initCommand.ExecuteReader();

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    var builder = new StringBuilder();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        builder.Append(reader.GetValue(i) + "___");
                    }

                    Console.WriteLine(builder);
                }
            }
            else
            {
                Console.WriteLine("No rows found.");
            }
        }

        public void InsertData(string tableName, DataTable tableData)
        {
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            var builder = new StringBuilder();
            foreach (DataRow row in tableData.Rows)
            {
                row["Id"] = Guid.NewGuid();
                row["PersonId"] = Guid.NewGuid();
                for (int i = 0; i < row.ItemArray.Length; i++)
                {
                    builder.Append(i == row.ItemArray.Length - 1 ? $"'{row.ItemArray[i]}'" : $"'{row.ItemArray[i]}',");
                }
                var commandString = $@"INSERT INTO {tableName} VALUES ({builder})";
                using var initCommand = new SqliteCommand(commandString, connection);

                initCommand.ExecuteNonQuery();
                builder.Clear();
            }
        }
    }
}