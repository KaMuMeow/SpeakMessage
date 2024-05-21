using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;

namespace SpeakMessage
{
    public class Program
    {
        static string apiKey = "";
        static string spreadsheetId = "your-spreadsheet-id";
        static string range = "Sheet1!A1:D10"; // 調整範圍以符合你的需求
        async static void Main(string[] args)
        {
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                ApiKey = apiKey,
                ApplicationName = "Google Sheets API .NET Quickstart",
            });

            var previousValues = await GetSheetValuesAsync(service, spreadsheetId, range);

            while (true)
            {
                var currentValues = await GetSheetValuesAsync(service, spreadsheetId, range);

                if (!AreValuesEqual(previousValues, currentValues))
                {
                    Console.WriteLine("Sheet has changed!");
                    previousValues = currentValues;
                }

                // 每隔一段時間檢查一次，這裡設置為 10 秒
                Thread.Sleep(10000);
            }
        }
        static async Task<IList<IList<object>>> GetSheetValuesAsync(SheetsService service, string spreadsheetId, string range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = await request.ExecuteAsync();
            return response.Values;
        }
        static bool AreValuesEqual(IList<IList<object>> values1, IList<IList<object>> values2)
        {
            if (values1 == null && values2 == null) return true;
            if (values1 == null || values2 == null) return false;
            if (values1.Count != values2.Count) return false;

            for (int i = 0; i < values1.Count; i++)
            {
                if (values1[i].Count != values2[i].Count) return false;
                for (int j = 0; j < values1[i].Count; j++)
                {
                    if (!values1[i][j].Equals(values2[i][j])) return false;
                }
            }
            return true;
        }
    }
}