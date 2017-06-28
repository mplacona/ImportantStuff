using System;
using System.Collections.Generic;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace ImportantStuff
{
	internal class Program
    {
	    private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
	    private const string ApplicationName = "Current Legislators";
	    private const string SpreadsheetId = "1P_0tngt7o02xgXr9T-wSaVXz-_OH0ZekYN8RLoWLnA4";
	    private const string Sheet = "legislators-current";
	    private static SheetsService _service;

	    private static void Main()
		{
			GoogleCredential credential;
			using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
			{
				credential = GoogleCredential.FromStream(stream)
					.CreateScoped(Scopes);
			}
		
			// Create Google Sheets API service.
			_service = new SheetsService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});
		
			CreateMultipleEntries(5);
		}

	    private static void ReadEntries()
		{
			var range = $"{Sheet}!A:F";
			SpreadsheetsResource.ValuesResource.GetRequest request =
					_service.Spreadsheets.Values.Get(SpreadsheetId, range);

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

	    private static void CreateEntry() { 
			var range = $"{Sheet}!A:F";
			var valueRange = new ValueRange();

			var oblist = new List<object>() { "Hello!", "This", "was", "insertd", "via", "C#" };
			valueRange.Values = new List<IList<object>> { oblist };

			var appendRequest = _service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
			appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
			appendRequest.Execute();
		}

	    private static void CreateMultipleEntries(int quantity = 1)
	    {
		    var range = $"{Sheet}!A:G";
		    for (var i = 0; i < quantity; i++){
			    var valueRange = new ValueRange();

		    	var oblist = new List<object>() {"Hello!", "This", "was", "insertd", "via", "C#", i};
		    	valueRange.Values = new List<IList<object>> {oblist};

				var appendRequest = _service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
				appendRequest.ValueInputOption =
					SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
				appendRequest.Execute();
	    	}
    	}

	    private static void UpdateEntry()
		{
			var range = $"{Sheet}!D543";
			var valueRange = new ValueRange();

			var oblist = new List<object>() { "updated" };
			valueRange.Values = new List<IList<object>> { oblist };

			var updateRequest = _service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
			updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
			updateRequest.Execute();
		}

	    private static void DeleteEntry()
		{
			var range = $"{Sheet}!A543:F";
			var requestBody = new ClearValuesRequest();

			var deleteRequest = _service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
			deleteRequest.Execute();
		}
    }
}
