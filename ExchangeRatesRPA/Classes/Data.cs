using System.Collections.Specialized;
using System.Configuration;

namespace ExchangeRatesRPA.Classes
{
    /// <summary>
    /// Класс, содержащий нужные данные из файла App.Config
    /// </summary>
    internal class Data
    {
        /// <summary>
        /// Поле файла конфигурации
        /// </summary>
        private static readonly NameValueCollection AppSettings = ConfigurationManager.AppSettings;
        
        /// <summary>
        /// Целевая валюта
        /// </summary>
        public static string? TargetCurrency { get; private set; }

        /// <summary>
        /// Почтовый адрес отправителя
        /// </summary>
        public static string? SenderEmail { get; private set; }

        /// <summary>
        /// Инициалы отправителя
        /// </summary>
        public static string? SenderName { get; private set; }

        /// <summary>
        /// Почтовый адрес получателя
        /// </summary>
        public static string? RecipientEmail { get; private set; }

        /// <summary>
        /// Инициалы получателя
        /// </summary>
        public static string? RecipientName { get; private set; }

        /// <summary>
        /// Объект класса с необходимыми данными
        /// </summary>
        public Data()
        {
            foreach (string? Key in AppSettings)
            {
                switch (Key)
                {
                    case "TargetCurrency":

                        TargetCurrency = AppSettings[Key];

                        break;

                    case "SenderEmail":

                        SenderEmail = AppSettings[Key];

                        break;

                    case "SenderName":

                        SenderName = AppSettings[Key];

                        break;

                    case "RecipientEmail":

                        RecipientEmail = AppSettings[Key];

                        break;

                    case "RecipientName":

                        RecipientName = AppSettings[Key];

                        break;
                }
            }
        }
    }
}