using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net.Mail;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace VendorExpiryEmailService
{
    public partial class Service1 : ServiceBase
    {
        private CancellationTokenSource cancellationTokenSource;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            cancellationTokenSource = new CancellationTokenSource();
            Task.Run(() => RunDailyTaskLoopAsync(cancellationTokenSource.Token));
            Log("Service started.");
        }

        protected override void OnStop()
        {
            cancellationTokenSource?.Cancel();
            Log("Service stopped.");
        }

        private async Task RunDailyTaskLoopAsync(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                DateTime now = DateTime.Now;
                DateTime nextRun = now.Date.AddHours(12); // Today at 9 AM

                if (now > nextRun)
                    nextRun = nextRun.AddDays(1); // Schedule for tomorrow 9 AM

                TimeSpan delay = nextRun - now;
                Log($"Next email scheduled at {nextRun}");

                try
                {
                    await Task.Delay(delay, token); // Wait until 9 AM
                }
                catch (TaskCanceledException)
                {
                    return; // Service stopped
                }

                try
                {
                    await SendExpiryEmailsAsync();
                    await GenerateAndSendCSVReportAsync();
                    Log("Daily tasks completed.");
                }
                catch (Exception ex)
                {
                    Log("Error during daily task: " + ex.Message);
                }
            }
        }

        private async Task SendExpiryEmailsAsync()
        {
            string connStr = ConfigurationManager.ConnectionStrings["emailsendermain"]?.ConnectionString;
            if (string.IsNullOrEmpty(connStr))
            {
                Log("Error: emailsendermain connection string is missing or empty.");
                return;
            }

            using (MySqlConnection conn = new MySqlConnection(connStr))
            {
                await conn.OpenAsync();

                MySqlCommand cmd = new MySqlCommand(
                    @"SELECT vendor_name, document_number, expiry_date, email 
                      FROM vendor 
                      WHERE expiry_date IS NOT NULL 
                      AND DATEDIFF(expiry_date, CURDATE()) IN (15, -15, -60)", conn);

                var expiryGroups = new Dictionary<DateTime, List<(string, string, string)>>();

                using (var dr = await cmd.ExecuteReaderAsync())
                {
                    while (await dr.ReadAsync())
                    {
                        string email = dr["email"]?.ToString()?.Trim() ?? "";
                        string name = dr["vendor_name"]?.ToString()?.Trim() ?? "";
                        string doc = dr["document_number"]?.ToString()?.Trim() ?? "";
                        DateTime expiry = Convert.ToDateTime(dr["expiry_date"]);

                        if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(doc))
                            continue;

                        if (!expiryGroups.ContainsKey(expiry))
                            expiryGroups[expiry] = new List<(string, string, string)>();

                        expiryGroups[expiry].Add((name, doc, email));
                    }
                }

                foreach (var kvp in expiryGroups)
                {
                    var expiryDate = kvp.Key;
                    var vendors = kvp.Value;
                    string to = string.Join(",", GetDistinctEmails(vendors));

                    string body = $"Dear Vendor(s),\n\nThe following documents are set to expire on {expiryDate:yyyy-MM-dd}:\n\n";
                    foreach (var v in vendors)
                        body += $"- {v.Item1}: Document {v.Item2}\n";

                    body += "\nPlease take the necessary actions.\n\nRegards,\nVendor Management Team";

                    string subject = $"Document Expiry Alert - {expiryDate:yyyy-MM-dd}";
                    await SendEmailAsync(to, subject, body);
                    Log($"Sent expiry alert for {expiryDate:yyyy-MM-dd} to {to}");
                }
            }
        }

        private IEnumerable<string> GetDistinctEmails(List<(string VendorName, string DocumentNumber, string Email)> vendors)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var v in vendors)
            {
                if (!string.IsNullOrWhiteSpace(v.Email) && seen.Add(v.Email))
                    yield return v.Email;
            }
        }

        private async Task GenerateAndSendCSVReportAsync()
        {
            string connStr = ConfigurationManager.ConnectionStrings["emailsendermain"]?.ConnectionString;
            if (string.IsNullOrEmpty(connStr))
            {
                Log("Error: emailsendermain connection string is missing or empty.");
                return;
            }

            string folderPath = ConfigurationManager.AppSettings["CSVFolder"];
            if (string.IsNullOrEmpty(folderPath))
            {
                Log("Error: CSVFolder app setting is missing.");
                return;
            }

            Directory.CreateDirectory(folderPath);
            string fileName = $"VendorReport_{DateTime.Now:yyyyMMddHHmmss}.csv";
            string fullPath = Path.Combine(folderPath, fileName);

            List<(string VendorName, string DocumentNumber, DateTime ExpiryDate, string Email)> report = new List<(string VendorName, string DocumentNumber, DateTime ExpiryDate, string Email)>();

            using (MySqlConnection conn = new MySqlConnection(connStr))
            {
                await conn.OpenAsync();
                MySqlCommand cmd = new MySqlCommand(
                    "SELECT vendor_name, document_number, expiry_date, email FROM vendor WHERE DATEDIFF(expiry_date, CURDATE()) > 20", conn);

                using (var dr = await cmd.ExecuteReaderAsync())
                {
                    while (await dr.ReadAsync())
                    {
                        string name = dr["vendor_name"]?.ToString() ?? "";
                        string doc = dr["document_number"]?.ToString() ?? "";
                        DateTime expiry = dr["expiry_date"] != DBNull.Value ? Convert.ToDateTime(dr["expiry_date"]) : DateTime.MinValue;
                        string email = dr["email"]?.ToString() ?? "";

                        report.Add((name, doc, expiry, email));
                    }
                }
            }

            using (StreamWriter sw = new StreamWriter(fullPath))
            {
                await sw.WriteLineAsync("Vendor Name,Document Number,Expiry Date,Email");
                foreach (var r in report)
                {
                    string line = $"{r.VendorName},{r.DocumentNumber},{r.ExpiryDate:yyyy-MM-dd},{r.Email}";
                    await sw.WriteLineAsync(line);
                }
            }

            var emails = new HashSet<string>();
            foreach (var v in report)
                if (!string.IsNullOrWhiteSpace(v.Email))
                    emails.Add(v.Email);

            await SendEmailWithAttachmentAsync(string.Join(",", emails), "Vendor Expiry Report", "Attached is the report for vendors with expiry > 20 days.", fullPath);
            Log("CSV report sent.");
        }

        private async Task SendEmailAsync(string to, string subject, string body)
        {
            using (SmtpClient smtp = new SmtpClient())
            {
                MailMessage mail = new MailMessage("jammyfaron@gmail.com", to, subject, body);
                await smtp.SendMailAsync(mail);
            }
        }

        private async Task SendEmailWithAttachmentAsync(string to, string subject, string body, string path)
        {
            using (SmtpClient smtp = new SmtpClient())
            {
                MailMessage mail = new MailMessage("jammyfaron@gmail.com", to, subject, body);
                mail.Attachments.Add(new Attachment(path));
                await smtp.SendMailAsync(mail);
            }
        }

        private void Log(string message)
        {
            string logPath = AppDomain.CurrentDomain.BaseDirectory + "log.txt";
            using (StreamWriter sw = new StreamWriter(logPath, true))
                sw.WriteLine($"{DateTime.Now}: {message}");
        }
    }
}
