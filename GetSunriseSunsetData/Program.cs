using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace GetSunriseSunsetData
{
	class Program
	{
		private static string _uriTemplate = "https://api.sunrise-sunset.org/json?lat={0}&lng={1}&date={2}";
		private static string _lat = "44.95897";
		private static string _lng = "-124.0145";
		private static DateOnly _startDate = new DateOnly(2023, 1, 1);
		private static int _days = 10;
		private static HttpClient _client = new HttpClient();

		static async Task Main()
		{
			DateOnly date;
			TimeOnly sunrise;
			TimeOnly sunset;
			Uri uri;
			string responseJSON;

			var daysTable = new DataTable("Days");
			daysTable.Columns.Add(new DataColumn("date", typeof(DateOnly)));
			daysTable.Columns.Add(new DataColumn("sunrise", typeof(TimeOnly)));
			daysTable.Columns.Add(new DataColumn("sunset", typeof(TimeOnly)));

			for (int i = 0; i < _days; i++)
			{
				//formulate URI
				date = _startDate.AddDays(i);
				uri = new Uri(string.Format(_uriTemplate, _lat, _lng, date.ToString("yyyy-MM-dd")));

				//send request & get response
				try
				{
					Console.Write($"Trying {date.ToString("yyyy-MM-dd")}...  ");
					using (HttpResponseMessage response = await _client.GetAsync(uri))
					{
						response.EnsureSuccessStatusCode();
						responseJSON = await response.Content.ReadAsStringAsync();
					}
					Console.WriteLine("succeeded");
				}
				catch (Exception)
				{
					Console.WriteLine("failed");
					continue;
				}

				//parse response JSON
				using (JsonDocument doc = JsonDocument.Parse(responseJSON))
				{
					JsonElement root = doc.RootElement;
					JsonElement results = root.GetProperty("results");
					sunrise = TimeOnly.FromDateTime(DateTime.Parse(results.GetProperty("sunrise").ToString()).ToLocalTime());
					sunset = TimeOnly.FromDateTime(DateTime.Parse(results.GetProperty("sunset").ToString()).ToLocalTime());
				}

				daysTable.Rows.Add(new Object[] { date, sunrise, sunset });
				Console.WriteLine($"{daysTable.Rows.Count} data points retrieved" + Environment.NewLine);
			}

			//write output to Excel
			var wbk = new XLWorkbook();
			var sht = wbk.AddWorksheet(daysTable, "data");
			sht.Columns(1, 3).AdjustToContents();
			Console.WriteLine(Environment.NewLine + "Data written to Excel" + Environment.NewLine);

			//save Excel file
			string outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
			string timestamp = DateTime.Now.ToString("yyyy-MM-dd-HHmmss");
			string outputFilePath = Path.Join(Path.Join(outputFolder, timestamp + "_SunriseSunsetData.xlsx"));
			Console.WriteLine($"{Environment.NewLine}About to save to {outputFilePath}");
			wbk.SaveAs(outputFilePath);
			wbk.Dispose();
			Console.WriteLine("Excel file saved" + Environment.NewLine);
		}
	}
}