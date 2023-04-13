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
		private static DateTime _startDate = new DateTime(2023, 1, 1);
		private static int _days = 10;
		private static HttpClient _client = new HttpClient();

		static async Task Main()
		{
			DateTime date;
			DateTime sunrise;
			DateTime sunset;
			Uri uri;
			string responseJSON;

			//var dataPoints = new List<DataPoint>();
			var daysTable = new System.Data.DataTable("Days");
			daysTable.Columns.Add(new DataColumn("date", typeof(DateTime))); //TODO - this didn't work out ideally format-wise, tempted to use string, see 2023-04-12-191447_SunriseSunsetData.xlsx
			daysTable.Columns.Add(new DataColumn("sunrise", typeof(DateTime)));
			daysTable.Columns.Add(new DataColumn("sunset", typeof(DateTime)));

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
					sunrise = DateTime.Parse(results.GetProperty("sunrise").ToString()).ToLocalTime();
					sunset = DateTime.Parse(results.GetProperty("sunset").ToString()).ToLocalTime();
				}

				daysTable.Rows.Add(new Object[] { date, sunrise, sunset });
				Console.WriteLine($"{daysTable.Rows.Count} data points retrieved");
			}

			//write output to Excel
			var wbk = new XLWorkbook();
			var sht = wbk.AddWorksheet("data");
			sht.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

			//populate & format header row
			sht.Cell(1, 1).Value = "date";
			sht.Cell(1, 2).Value = "sunrise";
			sht.Cell(1, 3).Value = "sunset";
			sht.Range("1:1").Style.Font.Bold = true;
			sht.Range("1:1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
			sht.SheetView.FreezeRows(1);

			//write data rows to Excel
			int rowIndex = 1;
			foreach (DataRow row in daysTable.Rows)
			{
				rowIndex++;
				sht.Cell(rowIndex, 1).Value = row["date"].ToString(); //TODO - wish I didn't need to call .ToString() here
				sht.Cell(rowIndex, 2).Value = row["sunrise"].ToString();
				sht.Cell(rowIndex, 3).Value = row["sunset"].ToString();
				Console.WriteLine($"Dumped {rowIndex - 1} rows to Excel");
			}

			//set date/time formats
			//TODO - these were basically ignored, see 2023-04-12-191447_SunriseSunsetData.xlsx
			sht.Range(2, 1, rowIndex, 1).Style.NumberFormat.NumberFormatId = 30; //m/d/yy or m-d-yy
			sht.Range(2, 2, rowIndex, 2).Style.NumberFormat.NumberFormatId = 18; //h:mm AM/PM
			sht.Range(2, 3, rowIndex, 3).Style.NumberFormat.NumberFormatId = 18; //h:mm AM/PM

			//save Excel file
			string outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
			string timestamp = DateTime.Now.ToString("yyyy-MM-dd-HHmmss");
			string outputFilePath = Path.Join(Path.Join(outputFolder, timestamp + "_SunriseSunsetData.xlsx"));
			Console.WriteLine($"About to save to {outputFilePath}");
			wbk.SaveAs(outputFilePath);
			wbk.Dispose();
		}
	}
}