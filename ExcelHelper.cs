using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace GymManagementSystem
{
    public static class ExcelHelper
    {
        private static string excelPath = "Clients.xlsx";

        static ExcelHelper()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static void AddClient(string fullName, string age, string weight, string height, string bundle, string phoneNumber)
        {
            DateTime startDate = DateTime.Today;
            DateTime endDate;

            if (bundle == "15 Days")
                endDate = startDate.AddDays(15);
            else if (bundle == "1 Month")
                endDate = startDate.AddMonths(1);
            else if (bundle == "3 Months")
                endDate = startDate.AddMonths(3);
            else
                endDate = startDate;

            FileInfo fileInfo = new FileInfo(excelPath);
            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets.Count == 0
                    ? package.Workbook.Worksheets.Add("Clients")
                    : package.Workbook.Worksheets["Clients"];

                int row = worksheet.Dimension?.Rows + 1 ?? 2;

                if (row == 2)
                {
                    worksheet.Cells[1, 1].Value = "Full Name";
                    worksheet.Cells[1, 2].Value = "Age";
                    worksheet.Cells[1, 3].Value = "Phone Number";
                    worksheet.Cells[1, 4].Value = "Weight";
                    worksheet.Cells[1, 5].Value = "Height";
                    worksheet.Cells[1, 6].Value = "Subscription";
                    worksheet.Cells[1, 7].Value = "Start Date";
                    worksheet.Cells[1, 8].Value = "End Date";
                    worksheet.Cells[1, 9].Value = "Frozen";
                }

                worksheet.Cells[row, 1].Value = fullName;
                worksheet.Cells[row, 2].Value = age;
                worksheet.Cells[row, 3].Value = phoneNumber;
                worksheet.Cells[row, 4].Value = weight;
                worksheet.Cells[row, 5].Value = height;
                worksheet.Cells[row, 6].Value = bundle;
                worksheet.Cells[row, 7].Value = startDate.ToShortDateString();
                worksheet.Cells[row, 8].Value = endDate.ToShortDateString();
                worksheet.Cells[row, 9].Value = "No";
                package.Save();
            }
        }

        public static List<Client> LoadAllClients()
        {
            List<Client> clients = new List<Client>();

            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists) return clients;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Clients"];
                if (worksheet == null) return clients;

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    Client client = new Client
                    {
                        FullName = worksheet.Cells[row, 1].Value?.ToString() ?? "",
                        Age = int.TryParse(worksheet.Cells[row, 2].Value?.ToString(), out int age) ? age : 0,
                        PhoneNumber = worksheet.Cells[row, 3].Value?.ToString() ?? "",
                        Weight = double.TryParse(worksheet.Cells[row, 4].Value?.ToString(), out double weight) ? weight : 0,
                        Height = double.TryParse(worksheet.Cells[row, 5].Value?.ToString(), out double height) ? height : 0,
                        SubscriptionType = worksheet.Cells[row, 6].Value?.ToString() ?? "",
                        SubscriptionStart = DateTime.TryParse(worksheet.Cells[row, 7].Value?.ToString(), out DateTime start) ? start : DateTime.MinValue,
                        SubscriptionEnd = DateTime.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out DateTime end) ? end : DateTime.MinValue,
                        IsFrozen = worksheet.Cells[row, 9].Value?.ToString() == "Yes"
                    };

                    clients.Add(client);
                }
            }

            return clients;
        }

        public static void EditClient(Client client)
        {
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists) return;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Clients"];
                if (worksheet == null) return;

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    if (worksheet.Cells[row, 4].Value?.ToString() == client.PhoneNumber)
                    {
                        worksheet.Cells[row, 1].Value = client.FullName;
                        worksheet.Cells[row, 2].Value = client.Age;
                        worksheet.Cells[row, 5].Value = client.Weight;
                        worksheet.Cells[row, 6].Value = client.Height;
                        break;
                    }
                }
                package.Save();
            }
        }

        public static void DeleteClient(string phoneNumber)
        {
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists) return;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Clients"];
                if (worksheet == null) return;

                int totalRows = worksheet.Dimension.Rows;
                for (int row = 2; row <= totalRows; row++)
                {
                    if (worksheet.Cells[row, 4].Value?.ToString() == phoneNumber)
                    {
                        worksheet.DeleteRow(row);
                        break;
                    }
                }
                package.Save();
            }
        }


        public static void FreezeClient(string phoneNumber, int freezeDays)
        {
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists) return;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Clients"];
                if (worksheet == null) return;

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var cellValue = worksheet.Cells[row, 3].Value?.ToString();
                    if (cellValue == phoneNumber)
                    {
                        if (DateTime.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out DateTime endDate))
                        {
                            endDate = endDate.AddDays(freezeDays);
                            worksheet.Cells[row, 8].Value = endDate.ToShortDateString();
                        }
                        else
                        {
                            worksheet.Cells[row, 8].Value = DateTime.Today.AddDays(freezeDays).ToShortDateString();
                        }

                        worksheet.Cells[row, 9].Value = "Yes";
                        break;
                    }
                }
                package.Save();
            }
        }

        public static void UnfreezeClient(string phoneNumber)
        {
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists) return;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Clients"];
                if (worksheet == null) return;

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var cellValue = worksheet.Cells[row, 3].Value?.ToString();
                    if (cellValue == phoneNumber)
                    {
                        worksheet.Cells[row, 9].Value = "No";
                        break;
                    }
                }
                package.Save();
            }
        }

        public static void AutoUnfreezeClients()
        {
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists) return;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Clients"];
                if (worksheet == null) return;

                bool changed = false;

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var frozenStatus = worksheet.Cells[row, 9].Value?.ToString();
                    var endDateStr = worksheet.Cells[row, 8].Value?.ToString();

                    if (frozenStatus == "Yes" && DateTime.TryParse(endDateStr, out DateTime endDate))
                    {
                        if (DateTime.Today >= endDate)
                        {
                            worksheet.Cells[row, 9].Value = "No"; // unfreeze
                            changed = true;
                        }
                    }
                }

                if (changed)
                {
                    package.Save();
                }
            }
        }


        public static void RenewClientSubscription(string phoneNumber, string bundle)
        {
            FileInfo fileInfo = new FileInfo(excelPath);
            if (!fileInfo.Exists) return;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Clients"];
                if (worksheet == null) return;

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    if (worksheet.Cells[row, 4].Value?.ToString() == phoneNumber)
                    {
                        DateTime newStartDate = DateTime.Today;
                        DateTime newEndDate;

                        if (bundle == "15 Days")
                            newEndDate = newStartDate.AddDays(15);
                        else if (bundle == "1 Month")
                            newEndDate = newStartDate.AddMonths(1);
                        else if (bundle == "3 Months")
                            newEndDate = newStartDate.AddMonths(3);
                        else
                            newEndDate = newStartDate;

                        worksheet.Cells[row, 8].Value = newStartDate.ToShortDateString();
                        worksheet.Cells[row, 9].Value = newEndDate.ToShortDateString();
                        worksheet.Cells[row, 10].Value = "No";
                        break;
                    }
                }
                package.Save();
            }
        }

        public static void AddLogEntry(string name, string phone)
        {
            string logPath = "Logs.xlsx";
            FileInfo fileInfo = new FileInfo(logPath);
            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets.Count == 0
                    ? package.Workbook.Worksheets.Add("Logs")
                    : package.Workbook.Worksheets["Logs"];

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
            }
        }

        public static void AddIncomeEntry(string name, string phone, string bundle)
        {
            string incomePath = "Income.xlsx";
            FileInfo fileInfo = new FileInfo(incomePath);
            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets.Count == 0
                    ? package.Workbook.Worksheets.Add("Income")
                    : package.Workbook.Worksheets["Income"];

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
            }
        }

        private static int GetAmount(string bundle)
        {
            if (bundle == "15 Days Gym only") return 200;
            if (bundle == "15 Days Gym and cardio") return 250;
            if (bundle == "1 Month Gym only") return 300;
            if (bundle == "1 Month Gym and cardio") return 350;
            if (bundle == "3 Months Gym only") return 750;
            if (bundle == "3 Months Gym and cardio") return 900;
            return 0;
        }

        public static List<LogEntry> GetLogs()
        {
            List<LogEntry> logs = new List<LogEntry>();
            string logPath = "Logs.xlsx";
            FileInfo fileInfo = new FileInfo(logPath);
            if (!fileInfo.Exists) return logs;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Logs"];
                if (worksheet == null) return logs;

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
            return logs;
        }

        public static List<IncomeEntry> GetIncome()
        {
            List<IncomeEntry> incomes = new List<IncomeEntry>();
            string incomePath = "Income.xlsx";
            FileInfo fileInfo = new FileInfo(incomePath);
            if (!fileInfo.Exists) return incomes;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets["Income"];
                if (worksheet == null) return incomes;

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
            return incomes;
        }


    }
}
