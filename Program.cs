using System.Globalization;
using CommandLine;
using CsvHelper;
using CsvHelper.Configuration;
using dotenv.net;
using Flurl.Http;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;

DotEnv.Load();
var config = new ConfigurationBuilder().AddEnvironmentVariables().Build();
var worldCatClient = new WorldCatSearchClient(config["WORLDCAT_CLIENT_ID"], config["WORLDCAT_CLIENT_SECRET"]);
await Parser.Default.ParseArguments<Options>(args).WithParsedAsync(async options =>
{
    var inputFile = new FileInfo(options.InputFile);
    var excel = new FastExcel.FastExcel(inputFile, readOnly: true);
    var sheet = excel.Read(1);
    var csv = new CsvWriter(Console.Out, new CsvConfiguration(CultureInfo.InvariantCulture));

    if (options.From is null)
    {
        var headers =
            sheet.Rows.First().Cells.Select(cell => cell.Value)
            .Concat(new object[] { "OCLC Number", "Count of Libraries", "Count of libraries summary" });
        foreach (var header in headers)
        {
            csv.WriteField(header);
        }
        csv.NextRecord();
    }
    foreach (var row in sheet.Rows.Skip((options.From ?? 2) - 1))
    {
        var cell = row.GetCellByColumnName("I");
        if (cell is null)
        {
            foreach (var field in row.Cells.Select(cell => cell.Value))
            {
                csv.WriteField(field);
            }
            for (int i = 0; i < 3; i++)
            {
                csv.WriteField(string.Empty);
            }
            csv.NextRecord();
            continue;
        }
        var oclcControlNumbers = cell.Value.ToString().Split("; ");
        var mergedNumbers = new HashSet<string>();
        var holdingCounts = new List<(string, int)>();
        foreach (var oclcControlNumber in oclcControlNumbers)
        {
            if (mergedNumbers.Contains(oclcControlNumber)) continue;
            var tryAgain = false;
            var numberOfTries = 0;
            do
            {
                try
                {
                    var holdings = await worldCatClient.GetHoldings(oclcControlNumber);
                    if (holdings.ContainsKey("briefRecords"))
                    {
                        foreach (var record in holdings["briefRecords"].Cast<JObject>())
                        {
                            if (record.ContainsKey("mergedOclcNumbers"))
                            {
                                mergedNumbers.UnionWith(record["mergedOclcNumbers"].Select(x => x.ToString()));
                            }
                            holdingCounts.Add((record["oclcNumber"].ToString(), record["institutionHolding"]?["briefHoldings"]?.Count() ?? 1));
                        }
                    }
                }
                catch (FlurlHttpException e) when (e.StatusCode == 400)
                {
                    Console.Error.WriteLine($"row {row.RowNumber}, oclcNumber {oclcControlNumber} errored:");
                    Console.Error.WriteLine(await e.GetResponseStringAsync());
                }
                catch (FlurlHttpException e) when (e.StatusCode == 500)
                {
                    numberOfTries++;
                    Console.Error.WriteLine($"row {row.RowNumber}, oclcNumber {oclcControlNumber} errored (attempt {numberOfTries}):");
                    if (numberOfTries >= 3) throw;
                    Console.Error.WriteLine(await e.GetResponseStringAsync());
                    tryAgain = true;
                    await Task.Delay(TimeSpan.FromMinutes(1));
                }
                catch (Exception)
                {
                    Console.Error.WriteLine($"row {row.RowNumber}, oclcNumber {oclcControlNumber} errored:");
                    throw;
                }
            } while (tryAgain);
        }
        foreach (var field in row.Cells.Select(cell => cell.Value))
        {
            csv.WriteField(field);
        }
        if (holdingCounts.Any())
        {
            var (oclcNumber, libraryCount) = holdingCounts.MaxBy(x => x.Item2);
            csv.WriteField(oclcNumber);
            csv.WriteField(libraryCount);
            csv.WriteField(GetSummary(libraryCount));
        }
        else
        {
            for (int i = 0; i < 3; i++)
            {
                csv.WriteField(string.Empty);
            }
        }
        csv.NextRecord();
        if (options.To is not null && row.RowNumber == options.To)
        {
            break;
        }
    }
});

static string GetSummary(int libraryCount)
{
    if (libraryCount <= 1) return "unique";
    if (libraryCount <= 5) return "few copies";
    if (libraryCount <= 49) return "some copies";
    return "many copies";
}
