using Fizzler.Systems.HtmlAgilityPack;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using HtmlAgilityPack;
using Newtonsoft.Json;
using RestSharp;
using RestSharp.Extensions;
using System.Data;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
using Telegram.Bot;
using Telegram.Bot.Types;

namespace LeboncoinMonitor
{
    [DebuggerStepThrough]
    class ImportantMethods
    {
        public static bool isNewRecord(ResultObj resultObj, List<string> alreadyScrapedURLs, string environmentPath)
        {
            var isNewRecord = true;
            for (int i = 0; i < alreadyScrapedURLs.Count; i++)
            {
                if (alreadyScrapedURLs[i] == resultObj.url) { isNewRecord = false; break; }
            }
            if (isNewRecord) System.IO.File.AppendAllLines($@"{environmentPath}\Misc\Files For Scraper\alreadyScrapedURLs.txt", new string[] { resultObj.url });
            return isNewRecord;
        }

        public static async Task CommunicateNewRecord(ResultObj resultObj, ScrapeConfig scrapeConfig, TelegramBotClient telegramBot)
        {
            var jobTitle = GetJobTitle(resultObj.Name, resultObj.url);

            for (int i = 0; i < scrapeConfig.TelegramChannelIds.Count; i++)
            {
                await telegramBot.SendTextMessageAsync(chatId: scrapeConfig.TelegramChannelIds[i], text: $"{jobTitle}\n\n" +
                $"<strong>📙 Ad Name:</strong>\t\t{resultObj.Name}\n", parseMode: Telegram.Bot.Types.Enums.ParseMode.Html);

                Thread.Sleep(TimeSpan.FromSeconds(5));
            }
        }
        public static string GetJobTitle(string title, string url)
        {
            return title.Length >= 30 ? $"<strong><a href=\"{url}\">{title.Substring(0, 30)}...</a> is a new ad ❗❗</strong>" : $"<strong><a href=\"{url}\">{title}</a> is a new ad ❗❗</strong>";
        }
        public static List<ResultObj> GetVehicles(HtmlNode document)
        {
            var resultObjs = new List<ResultObj>();
            var vehicleElements = document.SelectNodes("//div[contains(@class,'styles_adCard')]");
            foreach (var vehicleElement in vehicleElements)
            {
                var tempObj = new ResultObj();

                try { tempObj.url = "https://www.leboncoin.fr"+ vehicleElement.QuerySelector("a[data-test-id=\"ad\"]").GetAttributeValue("href", ""); } catch (Exception) { }
                try { tempObj.Name = vehicleElement.QuerySelector("p[data-qa-id=\"aditem_title\"]").InnerText; } catch (Exception) { }

                resultObjs.Add(tempObj);
            }
            return resultObjs;
        }

        public static ScrapeConfig GetConfig(SheetsService gService, string sheetId)
        {
            var gSheetConfigRawData = ImportantMethods.ReadSpreadsheetEntries("Config!A2:B9000", gService, sheetId).ToList();
            var scrapeConfig = new ScrapeConfig();
            var urlsToMonitor = new List<string>();
            var telegramChannelIds = new List<string>();

            for (int i = 0; i < gSheetConfigRawData.Count; i++)
            {
                try
                {
                    string urlToMonitor = "", telegramChannelId = "";

                    try { urlToMonitor = gSheetConfigRawData[i][0].ToString(); } catch (Exception) { }
                    try { telegramChannelId = gSheetConfigRawData[i][1].ToString(); } catch (Exception) { }

                    if (urlToMonitor.Length > 0) urlsToMonitor.Add(urlToMonitor);
                    if (telegramChannelId.Length > 0) telegramChannelIds.Add(telegramChannelId);
                }
                catch (Exception) { }
            }
            scrapeConfig.URLsToMonitor = urlsToMonitor;
            scrapeConfig.TelegramChannelIds= telegramChannelIds;
            return scrapeConfig;
        }
        public static void Logging(string environmentPath, Exception ex)
        {
            try { System.IO.File.Create($@"{environmentPath}\Misc\Log.txt").Dispose(); } catch (Exception) { }
            System.IO.File.AppendAllLines($@"{environmentPath}\Misc\Log.txt", new string[] { ex.ToString() });
        }
        public static HtmlNode GetHtmlNode(string responseStr)
        {
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(responseStr);
            var document = htmlDoc.DocumentNode;
            return document;
        }
        public static string GetResponse(string environmentPath, string linkToVisit, string headerType = "", bool isPostRequest = false, string body = "", bool useProxy = false)
        {
            var responseContent = "";
            for (int i = 0; i < 50 && responseContent.Length == 0 || responseContent.Contains("HTTP Error 400"); i++)
            {
                var client = new RestClient(linkToVisit)
                {
                    UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36",

                };
                if (useProxy)
                {
                    
                }
                var request = GetAppropriateHeader(headerType, isPostRequest, body, environmentPath);

                var response = client.ExecuteAsync(request);

                if (!response.Wait(30000)) { continue; };
                responseContent = response.Result.Content;
                if (responseContent.Length == 0 || responseContent.Contains("HTTP Error 400") ) { continue; }
            }

            return responseContent;
        }
        private static RestRequest GetAppropriateHeader(string headerType, bool isPostRequest, string body, string environmentPath = "")
        {
            var request = new RestRequest();
            IEnumerable<FileInfo> headerFile = null;

            if (isPostRequest) request = new RestRequest(Method.POST);

            headerFile = new DirectoryInfo($@"{environmentPath}\Misc\ApiCookies").GetFiles().Where(x => x.Name == $"{headerType}.txt");
            var fileContent = System.IO.File.ReadAllLines(headerFile.First().FullName);
            foreach (var line in fileContent)
            {
                try
                {
                    var headerName = Regex.Match(line, @".+?(?=\:)").Value.Trim();
                    var headerValue = Regex.Match(line, @"(?<=:).*").Value.Trim();

                    //if (headerName.ToLower().Contains("accept-encoding")) { headerValue = "application/json"; }

                    if (headerName.ToLower().Contains("user-agent")) { }
                    if (headerName.ToLower().Contains("content-length")) { }
                    else if (headerName.ToLower().Contains("host")) { }
                    else if (headerName.Contains("{") || headerValue.Contains("{")) { }
                    else request.AddHeader(headerName, headerValue);
                }
                catch (Exception) { }
            }
            if (body.Length > 0)
            {
                if (headerType == "services2")
                    request.AddParameter("application/x-www-form-urlencoded; charset=UTF-8", body, ParameterType.RequestBody);
                else
                    request.AddParameter("multipart/form-data; boundary=----WebKitFormBoundarywAjxGVrKcQAsXeO6", body, ParameterType.RequestBody);
            }

            return request;
        }
        public static void ClearSheet(string range, SheetsService sheetService, string googleSheet_sheetID)
        {
            // TODO: Assign values to desired properties of `requestBody`:
            Google.Apis.Sheets.v4.Data.ClearValuesRequest requestBody = new Google.Apis.Sheets.v4.Data.ClearValuesRequest();

            SpreadsheetsResource.ValuesResource.ClearRequest request = sheetService.Spreadsheets.Values.Clear(requestBody, googleSheet_sheetID, range);

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Google.Apis.Sheets.v4.Data.ClearValuesResponse response = request.Execute();

            // TODO: Change code below to process the `response` object:
            Console.WriteLine(JsonConvert.SerializeObject(response));
        }
        public static void SendEmail(string subject, string body, string destinationEmail, List<string> attachements)
        {
            string username = "suheylupworkbot@gmail.com", password = "stumynehzzollpzo";
            using (System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587))
            {
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(username, password);
                MailMessage msgObj = new MailMessage();
                msgObj.To.Add(destinationEmail);
                msgObj.From = new System.Net.Mail.MailAddress(username);
                msgObj.Subject = subject;
                msgObj.Body = body;
                msgObj.IsBodyHtml = true;
                for (int i = 0; i < attachements.Count; i++)
                    msgObj.Attachments.Add(new Attachment(attachements[i]));
                client.Send(msgObj);
            }
        }
        public static void SendEmail(string subject, string body, string destinationEmail)
        {
            string username = "suheylupworkbot@gmail.com", password = "stumynehzzollpzo";
            using (System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587))
            {
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(username, password);
                MailMessage msgObj = new MailMessage();
                msgObj.To.Add(destinationEmail);
                msgObj.From = new System.Net.Mail.MailAddress(username);
                msgObj.Subject = subject;
                msgObj.Body = body;
                msgObj.IsBodyHtml = true;
                client.Send(msgObj);
            }
        }
        public static void AppentIntoTopSpreadsheet2(List<string> objects, SheetsService sheetService, string googleSheet_sheetID, int sheetId)
        {
            InsertDimensionRequest insertRow = new InsertDimensionRequest();
            insertRow.Range = new DimensionRange()
            {
                SheetId = sheetId,
                Dimension = "ROWS",
                StartIndex = 1,
                EndIndex = 2
            };

            PasteDataRequest data = new PasteDataRequest
            {
                Data = string.Join(";%_", objects),
                Delimiter = ";%_",
                Coordinate = new GridCoordinate
                {
                    ColumnIndex = 0,
                    RowIndex = 1,
                    SheetId = sheetId
                },
            };

            BatchUpdateSpreadsheetRequest r = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>
                {
                    new Request{ InsertDimension = insertRow },
                    new Request{ PasteData = data }
                }
            };

            BatchUpdateSpreadsheetResponse response1 = sheetService.Spreadsheets.BatchUpdate(r, googleSheet_sheetID).Execute();

        }
        public static void AppentIntoSpreadsheet(string range, List<object> objects, SheetsService sheetService, string googleSheet_sheetID)
        {
            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { objects };
            var appendRequest = sheetService.Spreadsheets.Values.Append(valueRange, googleSheet_sheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }
        public static void AppentIntoTopSpreadsheet(List<string> objects, SheetsService sheetService, string googleSheet_sheetID, int sheetId)
        {
            InsertDimensionRequest insertRow = new InsertDimensionRequest();
            insertRow.Range = new DimensionRange()
            {
                SheetId = sheetId,
                Dimension = "ROWS",
                StartIndex = 1,
                EndIndex = 2
            };

            PasteDataRequest data = new PasteDataRequest
            {
                Data = string.Join(";%_", objects),
                Delimiter = ";%_",
                Coordinate = new GridCoordinate
                {
                    ColumnIndex = 0,
                    RowIndex = 1,
                    SheetId = sheetId
                },
            };

            BatchUpdateSpreadsheetRequest r = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>
                {
                    new Request{ InsertDimension = insertRow },
                    new Request{ PasteData = data }
                }
            };

            BatchUpdateSpreadsheetResponse response1 = sheetService.Spreadsheets.BatchUpdate(r, googleSheet_sheetID).Execute();

        }
        public static IList<IList<object>> ReadSpreadsheetEntries(string range, SheetsService sheetService, string googleSheet_sheetID)
        {
            var request = sheetService.Spreadsheets.Values.Get(googleSheet_sheetID, range);
            var response = request.Execute();
            var values = response.Values;
            return values;
        }
        public static SheetsService InitializeGoogleSheet()
        {
            GoogleCredential credential;
            string[] googleSheet_scopes = { SheetsService.Scope.Spreadsheets };
            using (var stream = new FileStream("projectaffordany-1587302173157-d20f1c21561a.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(googleSheet_scopes);
            }

            SheetsService sheetService;
            sheetService = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = "Task", });
            return sheetService;
        }
        public static List<string> removeDuplicatesFromList(List<string> input)
        {
            try { input = input.GroupBy(x => x).Select(x => x.First()).ToList(); } catch (Exception) { }
            return input;
        }
    }
}
