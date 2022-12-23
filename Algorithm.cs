using Fizzler.Systems.HtmlAgilityPack;
using Google.Apis.Sheets.v4;
using HtmlAgilityPack;
using System.Diagnostics;
using Telegram.Bot;
using Telegram.Bot.Types;

namespace LeboncoinMonitor
{
    internal class Algorithm
    {
        static string _environmentPath = "";
        static string _sheetId = "";
        static IniFile _configFile = null;
        static SheetsService _gService = null;
        static TelegramBotClient _telegramBot = null;
        static async Task Main(string[] args)
        {
            #region Init environment
            Console.WriteLine($"Starting LeboncoinMonitor.");
            _environmentPath = Path.GetFullPath(Path.Combine(@$"{Environment.CurrentDirectory}", @"..\..\..\"));
            _sheetId = "1-fiRqyCj4kxkFfKYqM80GatCwhh54fh0WLvgfyGnvIo";
            _configFile = new IniFile($@"{_environmentPath}\Misc\configFile.ini");
            _gService = ImportantMethods.InitializeGoogleSheet();
            _telegramBot = new TelegramBotClient("5800406984:AAHhcmepW8D8TGaV9dLIrhXPe4gwyrdgLL8");
            #endregion

            #region Algorithm
            var scrapeConfig = ImportantMethods.GetConfig(_gService, _sheetId);
            for (int x = 0; true; x++, Console.WriteLine($"Loop: {x}"))
            {
                try
                {
                    for (int i = 0; i < scrapeConfig.URLsToMonitor.Count; i++, Console.WriteLine($"scrapeConfig.URLsToMonitor: {i}"))
                    {
                        var response = ImportantMethods.GetResponse(environmentPath: _environmentPath, linkToVisit: scrapeConfig.URLsToMonitor[i], headerType: "leboncoin", isPostRequest: false, body: "", useProxy: false);
                        var document = ImportantMethods.GetHtmlNode(response);
                        var allRecords = ImportantMethods.GetVehicles(document);
                        var alreadyScrapedURLs = System.IO.File.ReadAllLines($@"{_environmentPath}\Misc\Files For Scraper\alreadyScrapedURLs.txt").ToList();
                        for (int b = 0; b < allRecords.Count; b++)
                        {
                            bool isNewRecordBool = ImportantMethods.isNewRecord(allRecords[b], alreadyScrapedURLs, _environmentPath);

                            if (alreadyScrapedURLs.Count > 0 && isNewRecordBool) await ImportantMethods.CommunicateNewRecord(allRecords[b], scrapeConfig, _telegramBot);
                        }
                    }
                    Thread.Sleep(TimeSpan.FromMinutes(1));
                }
                catch (Exception ex) { ImportantMethods.Logging(_environmentPath, ex); }
            }
            #endregion

            #region Dispose stuff
            #endregion
        }
    }
    class ResultObj
    {
        public string url { get; set; }
        public string Name { get; set; }
    }
    class ScrapeConfig
    {
        public List<string> URLsToMonitor { get; set; }
        public List<string> TelegramChannelIds { get; set; }
    }
}