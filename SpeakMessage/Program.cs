using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Speech.Synthesis;

namespace SpeakMessage
{
    public class Program
    {
        static string apiKey = "";
        static string spreadsheetId = "1FKo3E8T39vlhOmPzGJIMhWcjlHj3lpPyF2zeJzp-0BM";
        static string range = "工作表1!A:B"; // 調整範圍以符合你的需求
        static async Task Main(string[] args)
        {
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                ApiKey = apiKey,
                ApplicationName = "Google Sheets API .NET Quickstart",
            });

            var previousValues = await GetSheetValuesAsync(service, spreadsheetId, range);
            OutputFiled(service, previousValues);
            while (true)
            {
                var currentValues = await GetSheetValuesAsync(service, spreadsheetId, range);

                if (!AreValuesEqual(previousValues, currentValues))
                {
                    Console.WriteLine("Sheet has changed!");
                    previousValues = currentValues;
                    OutputFiled(service, previousValues);
                }

                // 每隔一段時間檢查一次，這裡設置為 1 秒
                Thread.Sleep(1000);
            }
        }
        static async Task<IList<IList<object>>> GetSheetValuesAsync(SheetsService service, string spreadsheetId, string range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = await request.ExecuteAsync();
            return response.Values;
        }

        static void UpdateCellValue(SheetsService service, string range, IList<object> newValue)
        {
            // 創建更新請求對象
            var valueRange = new ValueRange
            {
                Range = range,
                Values = new List<IList<object>> { newValue }
            };

            // 執行更新請求
            var request = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();
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
        static void OutputFiled(SheetsService service, IList<IList<object>> input)
        {
            string modifyRange = "工作表1!";
            foreach (var item in input)
            {
                modifyRange = "工作表1!";
                int index = input.IndexOf(item);
                Console.WriteLine($"value1: {item[0]} value2:{item[1]}");
                if (bool.Parse(item[1].ToString()) == true)
                {
                    SpeechText(item[0].ToString());
                    item[1] = "FALSE";
                    modifyRange += $"A{index + 1}:B{index + 1}";
                    UpdateCellValue(service, modifyRange, item);
                }
            }
        }
        static void SpeechText(string text)
        {
            // 創建 SpeechSynthesizer 實例
            using (SpeechSynthesizer synth = new SpeechSynthesizer())
            {
                // 設置語音
                synth.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);

                // 文字轉語音
                synth.Speak(text);
            }
        }
    }
}