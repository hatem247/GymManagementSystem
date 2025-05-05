using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Globalization;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows;

namespace GymManagementSystem
{
    public static class ExcelHelper
    {
        private static string excelPath = "Sheet.xlsx";

        static ExcelHelper()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static Client search(string phone_search)
        {
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var sheet = package.Workbook.Worksheets[0];
                for (int i = 2; i <= sheet.Dimension.End.Row; i++)
                {
                    string phone = sheet.Cells[i, 2].Text;
                    if (phone == phone_search)
                    {
                        Client client = new Client
                        {
                            FullName = sheet.Cells[i, 1].Text,
                            PhoneNumber = sheet.Cells[i, 2].Text,
                            Weight = double.Parse(sheet.Cells[i, 3].Text),
                            SubscriptionType = sheet.Cells[i, 4].Text,
                            SubscriptionStart = DateTime.Parse(sheet.Cells[i, 5].Text),
                            SubscriptionEnd = DateTime.Parse(sheet.Cells[i, 6].Text),
                            IsFrozen = sheet.Cells[i, 7].Text.ToLower() == "true"
                        };
                        return client;
                    }
                }
                MessageBox.Show("Client not found.");
                return null;
            }
        }
        public static bool AddClient(string fullName, string weight, string bundle, string subsciptionType, string phoneNumber)
        {
            try
            {
                DateTime startDate = DateTime.Today;
                DateTime endDate;

                switch (bundle)
                {
                    case "15 Days":
                        endDate = startDate.AddDays(15);
                        break;
                    case "1 Month":
                        endDate = startDate.AddMonths(1);
                        break;
                    case "3 Months":
                        endDate = startDate.AddMonths(3);
                        break;
                    case "6 Months":
                        endDate = startDate.AddMonths(6);
                        break;
                    default:
                        endDate = startDate;
                        break;
                }

                FileInfo fileInfo = new FileInfo(excelPath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"] ?? package.Workbook.Worksheets.Add("Clients");

                    int row = worksheet.Dimension?.Rows + 1 ?? 2;

                    if (row == 2)
                    {
                        worksheet.Cells[1, 1].Value = "Full Name";
                        worksheet.Cells[1, 2].Value = "Phone Number";
                        worksheet.Cells[1, 3].Value = "Weight";
                        worksheet.Cells[1, 4].Value = "Subscription";
                        worksheet.Cells[1, 5].Value = "Start Date";
                        worksheet.Cells[1, 6].Value = "End Date";
                        worksheet.Cells[1, 7].Value = "Frozen";
                    }

                    worksheet.Cells[row, 1].Value = fullName;
                    worksheet.Cells[row, 2].Value = phoneNumber;
                    worksheet.Cells[row, 3].Value = weight;
                    worksheet.Cells[row, 4].Value = bundle + " " + subsciptionType;
                    worksheet.Cells[row, 5].Value = startDate.ToShortDateString();
                    worksheet.Cells[row, 6].Value = endDate.ToShortDateString();
                    worksheet.Cells[row, 7].Value = "No";
                    package.Save();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding client: {ex.Message}");
                return false;
            }
        }

        public static List<Client> LoadAllClients()
        {
            List<Client> clients = new List<Client>();

            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists)
            {
                Console.WriteLine($"File not found: {excelPath}");
                return clients;
            }

            try
            {
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Worksheet 'Clients' not found.");
                        return clients;
                    }

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        Client client = new Client();

                        client.FullName = worksheet.Cells[row, 1].Value?.ToString() ?? "";

                        client.PhoneNumber = worksheet.Cells[row, 2].Value?.ToString() ?? "";

                        string weightString = worksheet.Cells[row, 3].Value?.ToString();
                        if (!string.IsNullOrEmpty(weightString) && double.TryParse(weightString, NumberStyles.Any, CultureInfo.InvariantCulture, out double weight))
                        {
                            client.Weight = weight;
                        }
                        else
                        {
                            client.Weight = 0;
                            Console.WriteLine($"Invalid weight format at row {row}. Setting weight to 0.");
                        }

                        client.SubscriptionType = worksheet.Cells[row, 4].Value?.ToString() ?? "";

                        string startDateString = worksheet.Cells[row, 5].Value?.ToString();
                        if (!string.IsNullOrEmpty(startDateString) && DateTime.TryParse(startDateString, out DateTime start))
                        {
                            client.SubscriptionStart = start;
                        }
                        else
                        {
                            client.SubscriptionStart = DateTime.MinValue;
                            Console.WriteLine($"Invalid start date format at row {row}. Setting start date to MinValue.");
                        }

                        string endDateString = worksheet.Cells[row, 6].Value?.ToString();
                        if (!string.IsNullOrEmpty(endDateString) && DateTime.TryParse(endDateString, out DateTime end))
                        {
                            client.SubscriptionEnd = end;
                        }
                        else
                        {
                            client.SubscriptionEnd = DateTime.MinValue;
                            Console.WriteLine($"Invalid end date format at row {row}. Setting end date to MinValue.");
                        }

                        client.IsFrozen = worksheet.Cells[row, 7].Value?.ToString() == "Yes";

                        clients.Add(client);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading clients: {ex.Message}");
            }

            return clients;
        }


        public static bool EditClient(Client client)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!fileInfo.Exists)
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Worksheet 'Clients' not found.");
                        return false;
                    }

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        if (worksheet.Cells[row, 2].Value?.ToString() == client.PhoneNumber)
                        {
                            worksheet.Cells[row, 1].Value = client.FullName;
                            worksheet.Cells[row, 3].Value = client.Weight;
                            package.Save();
                            return true;
                        }
                    }
                    Console.WriteLine($"Client with phone number {client.PhoneNumber} not found for editing.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error editing client: {ex.Message}");
                return false;
            }
        }


        public static bool DeleteClient(string phoneNumber)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!fileInfo.Exists)
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Worksheet 'Clients' not found.");
                        return false;
                    }

                    int totalRows = worksheet.Dimension.Rows;
                    for (int row = 2; row <= totalRows; row++)
                    {
                        if (worksheet.Cells[row, 2].Value?.ToString() == phoneNumber)
                        {
                            worksheet.DeleteRow(row);
                            package.Save();
                            return true;
                        }
                    }
                    Console.WriteLine($"Client with phone number {phoneNumber} not found for deletion.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error deleting client: {ex.Message}");
                return false;
            }
        }


        public static bool FreezeClient(string phoneNumber, int freezeDays)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!fileInfo.Exists)
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Worksheet 'Clients' not found.");
                        return false;
                    }

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var cellValue = worksheet.Cells[row, 2].Value?.ToString();
                        if (cellValue == phoneNumber)
                        {
                            if (DateTime.TryParse(worksheet.Cells[row, 6].Value?.ToString(), out DateTime endDate))
                            {
                                endDate = endDate.AddDays(freezeDays);
                                worksheet.Cells[row, 6].Value = endDate.ToShortDateString();
                            }
                            else
                            {
                                worksheet.Cells[row, 6].Value = DateTime.Today.AddDays(freezeDays).ToShortDateString();
                            }

                            worksheet.Cells[row, 7].Value = "Yes";
                            package.Save();
                            return true;
                        }
                    }
                    Console.WriteLine($"Client with phone number {phoneNumber} not found for freezing.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error freezing client: {ex.Message}");
                return false;
            }
        }


        public static bool UnfreezeClient(string phoneNumber)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!fileInfo.Exists)
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Worksheet 'Clients' not found.");
                        return false;
                    }

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var cellValue = worksheet.Cells[row, 2].Value?.ToString();
                        if (cellValue == phoneNumber)
                        {
                            worksheet.Cells[row, 7].Value = "No";
                            package.Save();
                            return true;
                        }
                    }
                    Console.WriteLine($"Client with phone number {phoneNumber} not found for unfreezing.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error unfreezing client: {ex.Message}");
                return false;
            }
        }


        public static bool AutoUnfreezeClients()
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!fileInfo.Exists)
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Worksheet 'Clients' not found.");
                        return false;
                    }

                    bool changed = false;

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var frozenStatus = worksheet.Cells[row, 7].Value?.ToString();
                        var endDateStr = worksheet.Cells[row, 6].Value?.ToString();

                        if (frozenStatus == "Yes" && DateTime.TryParse(endDateStr, out DateTime endDate))
                        {
                            if (DateTime.Today >= endDate)
                            {
                                worksheet.Cells[row, 7].Value = "No";
                                changed = true;
                            }
                        }
                    }

                    if (changed)
                    {
                        package.Save();
                        return true;
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error auto-unfreezing clients: {ex.Message}");
                return false;
            }
        }


        public static bool RenewClientSubscription(string phoneNumber, string bundle)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!fileInfo.Exists)
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Worksheet 'Clients' not found.");
                        return false;
                    }

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        if (worksheet.Cells[row, 2].Value?.ToString() == phoneNumber)
                        {
                            DateTime newStartDate = DateTime.Today;
                            DateTime newEndDate;

                            switch (bundle)
                            {
                                case "15 Days":
                                    newEndDate = newStartDate.AddDays(15);
                                    break;
                                case "1 Month":
                                    newEndDate = newStartDate.AddMonths(1);
                                    break;
                                case "3 Months":
                                    newEndDate = newStartDate.AddMonths(3);
                                    break;
                                default:
                                    newEndDate = newStartDate;
                                    break;
                            }

                            worksheet.Cells[row, 5].Value = newStartDate.ToShortDateString();
                            worksheet.Cells[row, 6].Value = newEndDate.ToShortDateString();
                            worksheet.Cells[row, 7].Value = "No";
                            package.Save();
                            return true;
                        }
                    }
                    Console.WriteLine($"Client with phone number {phoneNumber} not found for renewal.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error renewing client subscription: {ex.Message}");
                return false;
            }
        }


        public static bool AddLogEntry(string name, string phone)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Logs"] ?? package.Workbook.Worksheets.Add("Logs");

                    int row = worksheet.Dimension?.Rows + 1 ?? 2;

                    if (row == 2)
                    {
                        worksheet.Cells[1, 1].Value = "Name";
                        worksheet.Cells[1, 2].Value = "Phone";
                        worksheet.Cells[1, 3].Value = "Date";
                        worksheet.Cells[1, 4].Value = "Hour";
                        worksheet.Cells[1, 5].Value = "Minute";
                    }

                    DateTime now = DateTime.Now;
                    worksheet.Cells[row, 1].Value = name;
                    worksheet.Cells[row, 2].Value = phone;
                    worksheet.Cells[row, 3].Value = now.ToShortDateString();
                    worksheet.Cells[row, 4].Value = now.Hour;
                    worksheet.Cells[row, 5].Value = now.Minute;

                    package.Save();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding log entry: {ex.Message}");
                return false;
            }
        }


        public static bool AddIncomeEntry(string name, string phone, string bundle)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Income"] ?? package.Workbook.Worksheets.Add("Income");

                    int row = worksheet.Dimension?.Rows + 1 ?? 2;

                    if (row == 2)
                    {
                        worksheet.Cells[1, 1].Value = "Name";
                        worksheet.Cells[1, 2].Value = "Phone";
                        worksheet.Cells[1, 3].Value = "Date";
                        worksheet.Cells[1, 4].Value = "Hour";
                        worksheet.Cells[1, 5].Value = "Minute";
                        worksheet.Cells[1, 6].Value = "Amount";
                    }

                    DateTime now = DateTime.Now;
                    int amount = GetAmount(bundle);

                    worksheet.Cells[row, 1].Value = name;
                    worksheet.Cells[row, 2].Value = phone;
                    worksheet.Cells[row, 3].Value = now.ToShortDateString();
                    worksheet.Cells[row, 4].Value = now.Hour;
                    worksheet.Cells[row, 5].Value = now.Minute;
                    worksheet.Cells[row, 6].Value = amount;

                    package.Save();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding income entry: {ex.Message}");
                return false;
            }
        }


        private static int GetAmount(string bundle)
        {
            switch (bundle)
            {
                case "15 Days Gym only": return 200;
                case "15 Days Gym and cardio": return 250;
                case "1 Month Gym only": return 300;
                case "1 Month Gym and cardio": return 350;
                case "3 Months Gym only": return 750;
                case "3 Months Gym and cardio": return 900;
                case "6 Months Gym": return 1350;
                case "6 Months Gym and cardio": return 1600;
                default: return 0;
            }
        }


        public static List<LogEntry> GetLogs()
        {
            List<LogEntry> logs = new List<LogEntry>();
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists)
            {
                Console.WriteLine($"Log file not found: {excelPath}");
                return logs;
            }

            try
            {
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Logs"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Worksheet 'Logs' not found.");
                        return logs;
                    }

                    int rows = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rows; row++)
                    {
                        logs.Add(new LogEntry
                        {
                            Name = worksheet.Cells[row, 1].Text,
                            Phone = worksheet.Cells[row, 2].Text,
                            Date = worksheet.Cells[row, 3].Text,
                            Hour = worksheet.Cells[row, 4].Text,
                            Minute = worksheet.Cells[row, 5].Text
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting logs: {ex.Message}");
            }
            return logs;
        }


        public static List<IncomeEntry> GetIncome()
        {
            List<IncomeEntry> incomes = new List<IncomeEntry>();
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists)
            {
                Console.WriteLine($"Income file not found: {excelPath}");
                return incomes;
            }

            try
            {
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Income"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Worksheet 'Income' not found.");
                        return incomes;
                    }

                    int rows = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rows; row++)
                    {
                        incomes.Add(new IncomeEntry
                        {
                            Name = worksheet.Cells[row, 1].Text,
                            Phone = worksheet.Cells[row, 2].Text,
                            Date = worksheet.Cells[row, 3].Text,
                            Hour = worksheet.Cells[row, 4].Text,
                            Minute = worksheet.Cells[row, 5].Text,
                            Amount = worksheet.Cells[row, 6].Text
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting income: {ex.Message}");
            }
            return incomes;
        }
    }
}