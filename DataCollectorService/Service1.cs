using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.ServiceProcess;
using System.Timers;
using System.Linq;


namespace DataCollectorService
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer;
        private string inputDirectory = @"C:\Users\offic\OneDrive\Bureau\C#(Stage)\InputDirectory";
        private string processedDirectory = @"C:\Users\offic\OneDrive\Bureau\C#(Stage)\ProcessedDirectory";
        private string errorDirectory = @"C:\Users\offic\OneDrive\Bureau\C#(Stage)\ErrorDirectory";
        private string logsDirectory = @"C:\Users\offic\OneDrive\Bureau\C#(Stage)\LogsDirectory";

        public Service1()
        {
            InitializeComponent();
            timer = new Timer();
            timer.Interval = 60000; // 1 minute
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
        }

        protected override void OnStart(string[] args)
        {
            timer.Start();
        }

        protected override void OnStop()
        {
            timer.Stop();
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            ProcessFiles();
        }

        private void ProcessFiles()
        {
            try
            {
                string[] files = Directory.GetFiles(inputDirectory);

                foreach (var file in files)
                {
                    try
                    {
                        DataParser parser = new DataParser(this, errorDirectory);
                        var fileType = parser.DetectFileType(file);

                        List<object> results;
                        switch (fileType)
                        {
                            case DataFileType.Type1:
                                results = parser.ParseDataFile1(file).Cast<object>().ToList();
                                InsertIntoDatabase(results, "AnalysisResultsType1");
                                break;
                            case DataFileType.Type2:
                                results = parser.ParseDataFile2(file).Cast<object>().ToList();
                                InsertIntoDatabase(results, "AnalysisResultsType2");
                                break;
                            case DataFileType.Type3:
                                results = parser.ParseDataFile3(file).Cast<object>().ToList();
                                InsertIntoDatabase(results, "AnalysisResultsType3");
                                break;
                            case DataFileType.Type4:
                                results = parser.ParseDataFile4(file).Cast<object>().ToList();
                                InsertIntoDatabase(results, "AnalysisResultsType4");
                                break;
                            case DataFileType.Type5:
                                results = parser.ParseDataFile5(file).Cast<object>().ToList();
                                InsertIntoDatabase(results, "AnalysisResultsType5");
                                break;
                            case DataFileType.Type6:
                                results = parser.ParseDataFile6(file).Cast<object>().ToList();
                                InsertIntoDatabase(results, "AnalysisResults");
                                break;
                            default:
                                throw new Exception("Unknown file type.");
                        }

                        // Move file to processed directory only if no errors occurred
                        File.Move(file, Path.Combine(processedDirectory, Path.GetFileName(file)));

                        // Log success
                        LogToFile($"File {Path.GetFileName(file)} processing is complete.");
                    }
                    catch (Exception ex)
                    {
                        // Move file to error directory if an error occurred
                        File.Move(file, Path.Combine(errorDirectory, Path.GetFileName(file)));

                        // Log error
                        LogToFile($"Error processing file {Path.GetFileName(file)}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                LogToFile($"Unexpected error: {ex.Message}");
            }
        }






        private void InsertIntoDatabase(List<object> results, string tableName)
        {
            string connectionString = "Data Source=LAPTOP-N1447M01;Initial Catalog=DataCollectionDB;Integrated Security=True";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                foreach (var result in results)
                {
                    string query = "";
                    SqlCommand command = new SqlCommand();

                    if (result is DataEntry1 entry1)
                    {
                        query = $"INSERT INTO {tableName} (Number, SampleName, DateTime, Sum, ResultType, SiO2, Al2O3, Fe2O3, MnO, MgO, Na2O, K2O, TiO2, P2O5) " +
                                "VALUES (@Number, @SampleName, @DateTime, @Sum, @ResultType, @SiO2, @Al2O3, @Fe2O3, @MnO, @MgO, @Na2O, @K2O, @TiO2 ,@P2O5)";

                        command = new SqlCommand(query, connection);
                        command.Parameters.AddWithValue("@Number", entry1.Number);
                        command.Parameters.AddWithValue("@SampleName", entry1.SampleName);
                        command.Parameters.AddWithValue("@DateTime", entry1.DateTime);
                        command.Parameters.AddWithValue("@Sum", entry1.Sum);
                        command.Parameters.AddWithValue("@ResultType", entry1.ResultType);
                        command.Parameters.AddWithValue("@SiO2", entry1.SiO2);
                        command.Parameters.AddWithValue("@Al2O3", entry1.Al2O3);
                        command.Parameters.AddWithValue("@Fe2O3", entry1.Fe2O3);
                        command.Parameters.AddWithValue("@MnO", entry1.MnO);
                        command.Parameters.AddWithValue("@MgO", entry1.MgO);
                        command.Parameters.AddWithValue("@Na2O", entry1.Na2O);
                        command.Parameters.AddWithValue("@K2O", entry1.K2O);
                        command.Parameters.AddWithValue("@TiO2", entry1.TiO2);
                        command.Parameters.AddWithValue("@P2O5", entry1.P2O5);

                    }
                    else if (result is DataEntry2 entry2)
                    {
                        query = $"INSERT INTO {tableName} (Date, Time, Element, SampleID, Concentration, Unit) " +
                                "VALUES (@Date, @Time, @Element, @SampleID, @Concentration, @Unit)";

                        command = new SqlCommand(query, connection);
                        command.Parameters.AddWithValue("@Date", entry2.Date);
                        command.Parameters.AddWithValue("@Time", entry2.Time);
                        command.Parameters.AddWithValue("@Element", entry2.Element);
                        command.Parameters.AddWithValue("@SampleID", entry2.SampleID);
                        command.Parameters.AddWithValue("@Concentration", entry2.Concentration);
                        command.Parameters.AddWithValue("@Unit", entry2.Unit);
                    }

                    else if (result is DataEntry3 entry3)
                    {
                        query = $"INSERT INTO {tableName} (Date, Time, Element, SampleID, Concentration, Unit) " +
                                "VALUES (@Date, @Time, @Element, @SampleID, @Concentration, @Unit)";

                        command = new SqlCommand(query, connection);
                        command.Parameters.AddWithValue("@Date", entry3.Date);
                        command.Parameters.AddWithValue("@Time", entry3.Time);
                        command.Parameters.AddWithValue("@Element", entry3.Element);
                        command.Parameters.AddWithValue("@SampleID", entry3.SampleID);
                        command.Parameters.AddWithValue("@Concentration", entry3.Concentration);
                        command.Parameters.AddWithValue("@Unit", entry3.Unit);
                    }

                    else if (result is DataEntry4 entry4)
                    {
                        query = $"INSERT INTO {tableName} (Date, Name, AverageCarbon, AverageSulfur) " +
                                "VALUES (@Date, @Name, @AverageCarbon, @AverageSulfur)";

                        command = new SqlCommand(query, connection);
                        command.Parameters.AddWithValue("@Date", entry4.Date);
                        command.Parameters.AddWithValue("@Name", entry4.Name);
                        command.Parameters.AddWithValue("@AverageCarbon", entry4.AverageCarbon);
                        command.Parameters.AddWithValue("@AverageSulfur", entry4.AverageSulfur);
                    }

                    else if (result is DataEntry5 entry5)
                    {
                        query = $"INSERT INTO {tableName} (SampleId, Mn, Fe, Mg, Nb, Ca) " +
                                "VALUES (@SampleId, @Mn, @Fe, @Mg ,@Nb, @Ca)";

                        command = new SqlCommand(query, connection);
                        command.Parameters.AddWithValue("@SampleId", entry5.SampleId);
                        command.Parameters.AddWithValue("@Mn", entry5.Mn);
                        command.Parameters.AddWithValue("@Fe", entry5.Fe);
                        command.Parameters.AddWithValue("@Mg", entry5.Mg);
                        command.Parameters.AddWithValue("@Nb", entry5.Nb);
                        command.Parameters.AddWithValue("@Ca", entry5.Ca);
                    }

                    else if (result is DataEntry6 entry6)
                    {
                        query = $"INSERT INTO {tableName} (SampleID, ResultName, Date, Time, Element, Value) " +
                               "VALUES (@SampleID, @ResultName, @Date, @Time, @Element, @Value)";

                        command = new SqlCommand(query, connection);
                        command.Parameters.AddWithValue("@SampleID", entry6.SampleID);
                        command.Parameters.AddWithValue("@ResultName", entry6.ResultName);
                        command.Parameters.AddWithValue("@Date", entry6.Date);
                        command.Parameters.AddWithValue("@Time", entry6.Time);
                        command.Parameters.AddWithValue("@Element", entry6.Element);
                        command.Parameters.AddWithValue("@Value", entry6.Value);
                    }


                    command.ExecuteNonQuery();
                }
            }
        }



        public void LogToFile(string message)
        {
            string logFile = Path.Combine(logsDirectory, "service_log.txt");
            File.AppendAllText(logFile, $"{DateTime.Now} - {message}\n");
        }
    }
}
