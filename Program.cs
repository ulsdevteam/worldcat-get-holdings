using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using dotenv.net;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;

DotEnv.Load();
var config = new ConfigurationBuilder().AddEnvironmentVariables().Build();
var worldCatClient = new WorldCatSearchClient(config["WORLDCAT_CLIENT_ID"], config["WORLDCAT_CLIENT_SECRET"]);
var inputFile = new FileInfo(args[0]);
var excel = new FastExcel.FastExcel(inputFile, readOnly: true);
var sheet = excel.Read(1);
var csv = new CsvWriter(Console.Out, new CsvConfiguration(CultureInfo.InvariantCulture));

var headers =
    sheet.Rows.First().Cells.Select(cell => cell.Value)
    .Concat(new object[] { "OCLC Number", "Count of Libraries", "Count of libraries summary" });
foreach (var header in headers)
{
    csv.WriteField(header);
}
csv.NextRecord();
foreach (var row in sheet.Rows.Skip(1))
{
    var cell = row.GetCellByColumnName("I");
    if (cell is null) continue;
    var oclcControlNumbers = cell.Value.ToString().Split("; ");
    var mergedNumbers = new HashSet<string>();
    var holdingCounts = new List<(string, int)>();
    foreach (var oclcControlNumber in oclcControlNumbers)
    {
        if (mergedNumbers.Contains(oclcControlNumber)) continue;
        var holdings = await worldCatClient.GetHoldings(oclcControlNumber);
        foreach (var record in holdings["briefRecords"].Cast<JObject>())
        {
            if (record.ContainsKey("mergedOclcNumbers"))
            {
                mergedNumbers.UnionWith(record["mergedOclcNumbers"].Select(x => x.ToString()));
            }
            holdingCounts.Add((record["oclcNumber"].ToString(), record["institutionHolding"]["briefHoldings"].Count()));
        }
    }
    var (oclcNumber, libraryCount) = holdingCounts.MaxBy(x => x.Item2);
    var csvRecord = 
        row.Cells.Select(cell => cell.Value)
        .Concat(new object[] { oclcNumber, libraryCount, GetSummary(libraryCount) });
    foreach (var field in csvRecord)
    {
        csv.WriteField(field);
    }
    csv.NextRecord();
}

string GetSummary(int libraryCount)
{
    if (libraryCount == 1) return "unique";
    if (libraryCount <= 5) return "few copies";
    if (libraryCount <= 49) return "some copies";
    return "many copies";
}
