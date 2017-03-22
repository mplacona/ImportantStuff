using System;
using System.Collections.Generic;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace ImportantStuff
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Current Legislators";
        static readonly string spreadsheetId = "";
        static readonly string sheet = "";
		static SheetsService service;
        static void Main(string[] args)
        {
            GoogleCredential credential;
			using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
			{
				credential = GoogleCredential.FromStream(stream)
					.CreateScoped(Scopes);
			}

			// Create Google Sheets API service.
			service = new SheetsService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});

            DeleteEntry();
        }

        static void ReadEntries()
		{
			var range = $"{sheet}!A:F";
			SpreadsheetsResource.ValuesResource.GetRequest request =
					service.Spreadsheets.Values.Get(spreadsheetId, range);

			var response = request.Execute();
			IList<IList<object>> values = response.Values;
			if (values != null && values.Count > 0)
			{
				foreach (var row in values)
				{
					// Print columns A to F, which correspond to indices 0 and 4.
					Console.WriteLine("{0} | {1} | {2} | {3} | {4} | {5}", row[0], row[1], row[2], row[3], row[4], row[5]);
				}
			}
			else
			{
				Console.WriteLine("No data found.");
			}
		}

        static void CreateEntry() { 
			var range = $"{sheet}!A:F";
			var valueRange = new ValueRange();

			var oblist = new List<object>() { "Hello!", "This", "was", "insertd", "via", "C#" };
			valueRange.Values = new List<IList<object>> { oblist };

			var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
			appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
			var appendReponse = appendRequest.Execute();
		}

        static void UpdateEntry()
		{
			var range = $"{sheet}!D543";
			var valueRange = new ValueRange();

			var oblist = new List<object>() { "updated" };
			valueRange.Values = new List<IList<object>> { oblist };

			var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
			updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
			var appendReponse = updateRequest.Execute();
		}

        static void DeleteEntry()
		{
			var range = $"{sheet}!A543:F";
			var requestBody = new ClearValuesRequest();

			var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range);
			var deleteReponse = deleteRequest.Execute();
		}
    }
}
