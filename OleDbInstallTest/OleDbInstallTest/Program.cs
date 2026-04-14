using System;
using System.Data.OleDb;
using System.IO;
using System.Net.Http;
using System.Web.Script.Serialization;

namespace OleDbInstallTest
{
    class Program
    {
        private static string _logFilePath;

        static int Main(string[] args)
        {
            try
            {
                // Parse command line arguments
                if (!ParseArguments(args))
                {
                    Console.WriteLine("Usage: OleDbInstallTest.exe --logfile <path> --token <token> [--apiurl <url>]");
                    Console.WriteLine("  --logfile    Required. Path to the log file.");
                    Console.WriteLine("  --token      Required. Token to retrieve connection string from API.");
                    Console.WriteLine("  --apiurl     Optional. API base URL (default: http://localhost:5000)");
                    return 1;
                }

                // Get token and API URL from arguments
                string token = GetArgumentValue(args, "--token");
                if (string.IsNullOrEmpty(token))
                {
                    Console.WriteLine("Error: --token is required.");
                    return 1;
                }

                string apiUrl = GetArgumentValue(args, "--apiurl") ?? "http://localhost:5000";

                LogMessage($"Starting OLE DB Install Test");
                LogMessage($"Using API URL: {apiUrl}");
                LogMessage($"Using token: {token}");

                // Execute the insert
                ExecuteInsert(token, apiUrl);

                LogMessage("Insert completed successfully.");
                Console.WriteLine("Insert completed successfully.");
                return 0;
            }
            catch (Exception ex)
            {
                string errorMessage = $"Error: {ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMessage += $" | Inner: {ex.InnerException.Message}";
                }

                LogMessage(errorMessage);
                Console.WriteLine(errorMessage);
                return 1;
            }
        }

        private static bool ParseArguments(string[] args)
        {
            _logFilePath = GetArgumentValue(args, "--logfile");
            return !string.IsNullOrEmpty(_logFilePath);
        }

        private static string GetArgumentValue(string[] args, string argumentName)
        {
            for (int i = 0; i < args.Length - 1; i++)
            {
                if (args[i].Equals(argumentName, StringComparison.OrdinalIgnoreCase))
                {
                    return args[i + 1];
                }
            }
            return null;
        }

        private static void ExecuteInsert(string token, string apiUrl)
        {
            // Get connection string from web API
            string connectionString = GetConnectionStringFromApi(token, apiUrl);
            LogMessage($"Retrieved connection string from API");

            string sql = @"INSERT INTO log_table 
                           SELECT 'XX', bus_id, 'user1' AS user_id, NULL AS ecode, 
                                  'Install test' AS msg, GETDATE() AS enter_dttm";

            LogMessage($"Executing SQL: {sql.Replace(Environment.NewLine, " ").Replace("  ", " ")}");

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                LogMessage("Opening connection...");
                connection.Open();
                LogMessage("Connection opened successfully.");

                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    int rowsAffected = command.ExecuteNonQuery();
                    LogMessage($"Rows affected: {rowsAffected}");
                }
            }
        }

        private static string GetConnectionStringFromApi(string token, string apiUrl)
        {
            string requestUrl = $"{apiUrl.TrimEnd('/')}/api/connectionstring?token={Uri.EscapeDataString(token)}";
            LogMessage($"Requesting connection string from: {requestUrl}");

            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(30);
                
                var response = client.GetAsync(requestUrl).Result;
                string responseBody = response.Content.ReadAsStringAsync().Result;

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"API returned {(int)response.StatusCode}: {responseBody}");
                }

                // Parse JSON response using JavaScriptSerializer
                var serializer = new JavaScriptSerializer();
                var result = serializer.Deserialize<System.Collections.Generic.Dictionary<string, object>>(responseBody);

                if (result != null && result.ContainsKey("connectionString"))
                {
                    return result["connectionString"].ToString();
                }

                throw new Exception("API response did not contain connectionString");
            }
        }

        private static void LogMessage(string message)
        {
            if (string.IsNullOrEmpty(_logFilePath))
            {
                return;
            }

            try
            {
                // Ensure directory exists
                string directory = Path.GetDirectoryName(_logFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Create or append to log file
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
                
                using (StreamWriter writer = new StreamWriter(_logFilePath, append: true))
                {
                    writer.WriteLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                // If logging fails, write to console but don't throw
                Console.WriteLine($"Warning: Could not write to log file: {ex.Message}");
            }
        }
    }
}
