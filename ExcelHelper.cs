using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Globalization;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows;
using System.Linq;
using System.Text.RegularExpressions;

namespace GymManagementSystem
{
    public static class ExcelHelper
    {
        public static string excelPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sheet.xlsx");

        static ExcelHelper()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static Client Search(string phone_search)
        {
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var sheet = package.Workbook.Worksheets["Clients"];
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
                return null;
            }
        }

        public static bool AddClient(string fullName, string weight, string bundle, string subsciptionType, string phoneNumber)
        {
            try
            {
                var regex = new Regex(@"^01[0125][0-9]{8}$");
                if (!regex.IsMatch(phoneNumber))
                {
                    MessageBox.Show("Invalid Phone Number");
                    return false;
                }
                Client client = Search(phoneNumber);
                if (client != null)
                {
                    MessageBox.Show("This Phone number registered before");
                    return false;
                }
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
                MessageBox.Show($"Error adding client: {ex.Message}");
                return false;
            }
        }

        public static List<Client> LoadAllClients()
        {
            List<Client> clients = new List<Client>();

            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists)
            {
                MessageBox.Show($"File not found: {excelPath}");
                return clients;
            }

            try
            {
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Clients' not found.");
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
                            MessageBox.Show($"Invalid weight format at row {row}. Setting weight to 0.");
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
                            MessageBox.Show($"Invalid start date format at row {row}. Setting start date to MinValue.");
                        }

                        string endDateString = worksheet.Cells[row, 6].Value?.ToString();
                        if (!string.IsNullOrEmpty(endDateString) && DateTime.TryParse(endDateString, out DateTime end))
                        {
                            client.SubscriptionEnd = end;
                        }
                        else
                        {
                            client.SubscriptionEnd = DateTime.MinValue;
                            MessageBox.Show($"Invalid end date format at row {row}. Setting end date to MinValue.");
                        }

                        client.IsFrozen = worksheet.Cells[row, 7].Value?.ToString() == "Yes";

                        clients.Add(client);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading clients: {ex.Message}");
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
                    MessageBox.Show($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Clients' not found.");
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
                    MessageBox.Show($"Client with phone number {client.PhoneNumber} not found for editing.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error editing client: {ex.Message}");
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
                    MessageBox.Show($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Clients' not found.");
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
                    MessageBox.Show($"Client with phone number {phoneNumber} not found for deletion.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting client: {ex.Message}");
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
                    MessageBox.Show($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Clients' not found.");
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
                    MessageBox.Show($"Client with phone number {phoneNumber} not found for freezing.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error freezing client: {ex.Message}");
                return false;
            }
        }

        public static bool AddExtraDays(string phoneNumber, int Days)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!fileInfo.Exists)
                {
                    MessageBox.Show($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Clients' not found.");
                        return false;
                    }

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var cellValue = worksheet.Cells[row, 2].Value?.ToString();
                        if (cellValue == phoneNumber)
                        {
                            if (DateTime.TryParse(worksheet.Cells[row, 6].Value?.ToString(), out DateTime endDate))
                            {
                                endDate = endDate.AddDays(Days);
                                worksheet.Cells[row, 6].Value = endDate.ToShortDateString();
                            }
                            else
                            {
                                worksheet.Cells[row, 6].Value = DateTime.Today.AddDays(Days).ToShortDateString();
                            }
                            package.Save();
                            return true;
                        }
                    }
                    MessageBox.Show($"Client with phone number {phoneNumber} not found for freezing.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error freezing client: {ex.Message}");
                return false;
            }
        }

        public static bool UnfreezeClient(string phoneNumber)
        {
            try
            {
                Client client = Search(phoneNumber);
                int daysFreezed = 0;
                DateTime oldEndDate = DateTime.MinValue;
                if (client != null)
                {
                    if (client.SubscriptionType == "3 Months") oldEndDate = client.SubscriptionStart.AddMonths(3);
                    else oldEndDate = client.SubscriptionStart.AddMonths(6);
                    daysFreezed = (client.SubscriptionEnd - oldEndDate).Days;
                    var logs = GetLogs("");
                    var filtered = logs.Where(l => l.Phone == phoneNumber).ToList();
                    DateTime lastlog = DateTime.Parse(filtered[filtered.Count - 1].Date);
                    int actualFreezed = (DateTime.Today - lastlog).Days;
                    client.SubscriptionEnd.AddDays(actualFreezed - daysFreezed);
                }
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!fileInfo.Exists)
                {
                    MessageBox.Show($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Clients' not found.");
                        return false;
                    }

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var cellValue = worksheet.Cells[row, 2].Value?.ToString();
                        if (cellValue == phoneNumber)
                        {
                            worksheet.Cells[row, 6].Value = client.SubscriptionEnd.ToShortDateString();
                            worksheet.Cells[row, 7].Value = "No";
                            package.Save();
                            return true;
                        }
                    }
                    MessageBox.Show($"Client with phone number {phoneNumber} not found for unfreezing.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error unfreezing client: {ex.Message}");
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
                    MessageBox.Show($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Clients' not found.");
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
                MessageBox.Show($"Error auto-unfreezing clients: {ex.Message}");
                return false;
            }
        }

        public static bool RenewClientSubscription(string phoneNumber, string bundleDurition, string subscriptionType)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(excelPath);
                if (!fileInfo.Exists)
                {
                    MessageBox.Show($"File not found: {excelPath}");
                    return false;
                }

                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Clients"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Clients' not found.");
                        return false;
                    }

                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        if (worksheet.Cells[row, 2].Value?.ToString() == phoneNumber)
                        {
                            DateTime newStartDate = DateTime.Today;
                            DateTime newEndDate;

                            switch (bundleDurition)
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
                                case "6 Months":
                                    newEndDate = newStartDate.AddMonths(6);
                                    break;
                                default:
                                    newEndDate = newStartDate;
                                    break;
                            }

                            worksheet.Cells[row, 4].Value = bundleDurition + " " + subscriptionType;
                            worksheet.Cells[row, 5].Value = newStartDate.ToShortDateString();
                            worksheet.Cells[row, 6].Value = newEndDate.ToShortDateString();
                            worksheet.Cells[row, 7].Value = "No";
                            AddIncomeEntry(worksheet.Cells[row, 1].Value.ToString(), phoneNumber, bundleDurition + " " + subscriptionType);
                            package.Save();
                            return true;
                        }
                    }
                    MessageBox.Show($"Client with phone number {phoneNumber} not found for renewal.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error renewing client subscription: {ex.Message}");
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
                        worksheet.Cells[1, 4].Value = "Time";
                    }

                    DateTime now = DateTime.Now;
                    worksheet.Cells[row, 1].Value = name;
                    worksheet.Cells[row, 2].Value = phone;
                    worksheet.Cells[row, 3].Value = now.ToShortDateString();
                    worksheet.Cells[row, 4].Value = now.ToString("hh:mm:ss tt");

                    package.Save();
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding log entry: {ex.Message}");
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
                        worksheet.Cells[1, 4].Value = "Time";
                        worksheet.Cells[1, 5].Value = "Amount";
                    }

                    DateTime now = DateTime.Now;
                    int amount = GetAmount(bundle);

                    worksheet.Cells[row, 1].Value = name;
                    worksheet.Cells[row, 2].Value = phone;
                    worksheet.Cells[row, 3].Value = now.ToShortDateString();
                    worksheet.Cells[row, 4].Value = now.ToString("hh:mm:ss tt");
                    worksheet.Cells[row, 5].Value = amount;

                    package.Save();
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding income entry: {ex.Message}");
                return false;
            }
        }

        public static (DateTime start, DateTime end) GetDateRange(string filter)
        {
            var today = DateTime.Today;

            switch (filter)
            {
                case "Today":
                    return (today, today.AddDays(1).AddTicks(-1));

                case "Yesterday":
                    var yesterday = today.AddDays(-1);
                    return (yesterday, today.AddTicks(-1));

                case "Last Week":
                    var firstDayOfWeek = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;
                    var daysToSubtract = (int)today.DayOfWeek - (int)firstDayOfWeek - 7;
                    var lastWeekStart = today.AddDays(daysToSubtract);
                    var lastWeekEnd = lastWeekStart.AddDays(7).AddTicks(-1);
                    return (lastWeekStart, lastWeekEnd);

                case "Last Month":
                    var lastMonthStart = new DateTime(today.Year, today.Month, 1).AddMonths(-1);
                    var lastMonthEnd = new DateTime(today.Year, today.Month, 1).AddTicks(-1);
                    return (lastMonthStart, lastMonthEnd);

                default:
                    throw new ArgumentException("Invalid filter. Supported values are: Today, Yesterday, Last Week, Last Month.");
            }
        }

        public static int GetAmount(string bundle)
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

        public static List<LogEntry> GetLogs(string filter)
        {
            List<LogEntry> logs = new List<LogEntry>();
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists)
            {
                MessageBox.Show($"Log file not found: {excelPath}");
                return logs;
            }

            try
            {
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Logs"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Logs' not found.");
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
                            Time = worksheet.Cells[row, 4].Text
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting logs: {ex.Message}");
            }
            if(filter == "") return logs;
            else
            {
                var (start, end) = GetDateRange(filter);
                return logs.Where(i =>
                {
                    if (DateTime.TryParse(i.Date, out DateTime parsedDate))
                    {
                        return parsedDate >= start && parsedDate <= end;
                    }
                    return false;
                }).ToList();
            }
        }

        public static List<IncomeEntry> GetIncome(string filter)
        {
            List<IncomeEntry> incomes = new List<IncomeEntry>();
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists)
            {
                MessageBox.Show($"Excel file not found: {excelPath}");
                return incomes;
            }

            try
            {
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Income"];
                    if (worksheet == null)
                    {
                        MessageBox.Show("Worksheet 'Income' not found.");
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
                            Time = worksheet.Cells[row, 4].Text,
                            Amount = worksheet.Cells[row, 5].Text
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting income: {ex.Message}");
            }

            if(filter == "")
            {
                return incomes;
            }

            else
            {
                var (start, end) = GetDateRange(filter);
                return incomes.Where(i =>
                {
                    if (DateTime.TryParse(i.Date, out DateTime parsedDate))
                    {
                        return parsedDate >= start && parsedDate <= end;
                    }
                    return false;
                }).ToList();
            }
        }
    }
}