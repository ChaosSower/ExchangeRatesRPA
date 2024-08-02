using System.Text;
using System.Net.Mail;
using System.Net;

using ExchangeRatesRPA.Classes;

using HtmlAgilityPack;

namespace ExchangeRatesRPA
{
    public class Program
    {
        private static async Task Main()
        {
            try
            {
                Data Data = new();
                decimal CurrencyRate = await ParseSiteForCurrencyRate(Data.TargetCurrency ?? string.Empty); // получение курса заданной валюты
                string FilePath = CreateCsvFile(CurrencyRate, Data.TargetCurrency ?? string.Empty); // создание CSV-файла
                SendEmail(Data.SenderEmail ?? string.Empty, Data.SenderName ?? string.Empty, Data.RecipientEmail ?? string.Empty, Data.RecipientName ?? string.Empty, Data.TargetCurrency ?? string.Empty, FilePath); // отправка файла на почту
            }

            catch (Exception Exception)
            {
                Console.WriteLine($"Произошло исключение: {Exception.Message}");
            }
        }

        /// <summary>
        /// Асинхронная задача получения данных о валюте
        /// </summary>
        /// <param name="TargetCurrency">Целевая валюта</param>
        /// <returns>Значение целевой валюты</returns>
        /// <exception cref="Exception"></exception>
        private static async Task<decimal> ParseSiteForCurrencyRate(string TargetCurrency)
        {
            using HttpClient HttpClient = new();
            string Response = await HttpClient.GetStringAsync("https://www.cbr.ru/currency_base/daily/");

            HtmlDocument HtmlDocument = new();
            HtmlDocument.LoadHtml(Response);

            HtmlNodeCollection HtmlNodeCollection = HtmlDocument.DocumentNode.SelectNodes($"//tr[td[contains(text(), '{TargetCurrency}')]]/td"); // поиск курса валюты в таблице

            if (HtmlNodeCollection != null && HtmlNodeCollection.Count >= 2)
            {
                if (decimal.TryParse(HtmlNodeCollection[1].InnerText, out decimal CurrencyRate))
                {
                    return CurrencyRate;
                }
            }

            throw new Exception($"Не удалось получить курс валюты [{TargetCurrency}].");
        }

        /// <summary>
        /// Асинхронная задача создания файла .CSV с получеными данными для отправки
        /// </summary>
        /// <param name="CurrencyRate">Значение валюты</param>
        /// <param name="CurrencyName">Название валюты</param>
        /// <returns></returns>
        private static string CreateCsvFile(decimal CurrencyRate, string CurrencyName)
        {
            string FilePath = Path.Combine(Path.GetTempPath(), $"[{CurrencyName}].csv");

            StringBuilder StringBuilder = new();
            StringBuilder.AppendLine("Валюта,Дата,Курс");
            StringBuilder.AppendLine($"{CurrencyName},{DateTime.Now:yyyy-MM-dd},{CurrencyRate}");

            File.WriteAllText(FilePath, StringBuilder.ToString(), Encoding.UTF8);

            return FilePath;
        }

        /// <summary>
        /// Асинхронная задача отправки файла .CSV на почту
        /// </summary>
        /// <param name="SenderEmail">Почта отправки</param>
        /// <param name="SenderName">Инициалы отправителя</param>
        /// <param name="RecipientEmail">Почта назначения</param>
        /// <param name="RecipientName">Инициалы получателя</param>
        /// <param name="CurrencyName">Название валюты</param>
        /// <param name="AttachmentPath">Путь к файлу .CSV</param>
        private static void SendEmail(string SenderEmail, string SenderName, string RecipientEmail, string RecipientName, string CurrencyName, string AttachmentPath)
        {
            MailAddress FromMailAddress = new($"{SenderEmail}", $"{SenderName}");
            MailAddress ToMailAddress = new($"{RecipientEmail}", $"{RecipientName}");
            string MailSubject = $"Курс валюты [{CurrencyName}] на текущую дату";
            string MailBody = $"Прикреплён файл с курсом валюты [{CurrencyName}] на текущую дату.\n" +
                $"Сообщение отправлено автоматически, отвечать на него не нужно.";

            SmtpClient SmtpClient = new()
            {
                Host = "smtp.example.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential("your_email@example.com", "your_password")
            };

            using MailMessage Message = new(FromMailAddress, ToMailAddress)
            {
                Subject = MailSubject,
                Body = MailBody,
                IsBodyHtml = true
            };

            Message.Attachments.Add(new(AttachmentPath));
            SmtpClient.Send(Message);
        }
    }
}