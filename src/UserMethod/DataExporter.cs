﻿using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace WiniumTests.src {
    /// <summary>
    /// Exports Parsed Data from the .txt file the tests output on cleanup and sends that to sql
    /// </summary>
    class DataExporter {
        private readonly SqlConnection connection;
        public DataExporter(string connectionString) {
            connection = new SqlConnection(connectionString);
        }
        private void AddToTestRunTable(TestData data) {

            //running the command
            SqlCommand command = new SqlCommand();
            connection.Open();
            command.CommandTimeout = 60;
            command.Connection = connection;
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "spAddTestRunData";

            //adding the values to the parameter
            command.Parameters.Add("testerName", SqlDbType.NVarChar, 50).Value = data.CreatedBy;
            command.Parameters.Add("applicationName", SqlDbType.NVarChar, 50).Value = data.ApplicationName;
            command.Parameters.Add("applicationVersion", SqlDbType.NVarChar, 50).Value = data.ApplicationVersion;
            command.Parameters.Add("testsFailed", SqlDbType.SmallInt).Value = data.TestsFailed;
            command.Parameters.Add("testsPassed", SqlDbType.SmallInt).Value = data.TestsPassed;
            command.Parameters.Add("createdDate", SqlDbType.DateTime).Value = data.CreatedDate;
            command.Parameters.Add("createdBy", SqlDbType.NVarChar, 50).Value = data.CreatedBy;
            command.Parameters.Add("modifiedBy", SqlDbType.NVarChar, 50).Value = data.CreatedBy;
            command.Parameters.Add("modifiedDate", SqlDbType.DateTime, 50).Value = data.CreatedDate;

            command.ExecuteNonQuery();
            data.TestRunId = int.Parse(command.ExecuteScalar().ToString());
            connection.Close();
        }
        private void AddToTestCaseTable(TestData data) {
            //running the command
            SqlCommand command = new SqlCommand();
            connection.Open();
            command.CommandTimeout = 60;
            command.Connection = connection;
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "spAddTestCaseData";

            //adding the values to the parameter
            command.Parameters.Add("testName", SqlDbType.NVarChar, 50).Value = data.TestName;
            command.Parameters.Add("testStatusId", SqlDbType.BigInt).Value = data.TestStatus;
            command.Parameters.Add("testDate", SqlDbType.DateTime, 50).Value = data.CreatedDate;
            command.Parameters.Add("imagePath", SqlDbType.NVarChar, 256).Value = data.ImagePath;
            command.Parameters.Add("createdDate", SqlDbType.DateTime).Value = data.CreatedDate;
            command.Parameters.Add("createdBy", SqlDbType.NVarChar, 50).Value = data.CreatedBy;
            command.Parameters.Add("modifiedBy", SqlDbType.NVarChar, 50).Value = data.CreatedBy;
            command.Parameters.Add("modifiedDate", SqlDbType.DateTime, 50).Value = data.CreatedDate;
            command.Parameters.Add("testRunId", SqlDbType.BigInt, 50).Value = data.TestRunId - 1;

            command.ExecuteNonQuery();
            connection.Close();
        }
        public void ParseFile(TestData data) {
            foreach (string file in Directory.EnumerateFiles(ConfigurationManager.AppSettings.Get("FileLocation"))) {
                string[] lines = File.ReadAllLines(file);


                data.CreatedDate = DateTime.Parse(lines[1].Substring(0, lines[1].Length - 2));
                data.TestsPassed = int.Parse(lines[2].Substring(14));
                data.TestsFailed = int.Parse(lines[3].Substring(14));
                data.CreatedBy = lines[4].Substring(8);
                data.ApplicationVersion = lines[5].Substring(13);
                data.ApplicationName = lines[6].Substring(10);

                AddToTestRunTable(data);

                int i = 7;
                while (i < lines.Length) {
                    data.TestName = lines[i].Substring(8);

                    if (lines[i].Substring(0, 8).Equals("passed| ")) {
                        data.TestStatus = 1;
                        data.ImagePath = null;
                        i++;
                    } else {
                        data.TestStatus = 0;
                        data.ImagePath = lines[i + 1];
                        i += 2;
                    }
                    AddToTestCaseTable(data);
                }
                File.Move(file, ConfigurationManager.AppSettings.Get("ReadFileLocation") + "Test " + (data.TestRunId - 1).ToString() + ".txt");
            }
        }
    }
    class TestData {
        public int TestsFailed { get; set; }
        public int TestsPassed { get; set; }
        public string CreatedBy { get; set; }
        public DateTime CreatedDate { get; set; }
        public string ImagePath { get; set; }
        public string ApplicationName { get; set; }
        public string ApplicationVersion { get; set; }
        public string TestName { get; set; }
        public int TestStatus { get; set; }
        public int TestRunId { get; set; }
    }
}
