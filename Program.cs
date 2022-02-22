using CsvHelper;
using CsvHelper.Configuration.Attributes;
using Syroot.Windows.IO;
using System.Globalization;

var downloadFolderPath = new KnownFolder(KnownFolderType.Downloads).Path;
var emailsFile = Path.Combine(downloadFolderPath, "20220221_Emails.CSV");
using var reader = new StreamReader(emailsFile);
using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
var emails = csv.GetRecords<Email>().ToList();
Console.WriteLine($"Found {emails.Count} emails");
var groupedEmails = emails.GroupBy(e => new
{
    e.From_Address,
    From_Name = e.From_Name,
}).Select(g => new { From = g.Key.From_Address, Count = g.Count(), Name = g.Key.From_Name }).Where(g => g.Count > 5);
Console.WriteLine($"Found {groupedEmails.Count()} groups");
foreach (var group in groupedEmails.OrderByDescending(g => g.Count))
{
    Console.WriteLine($"{group.From,48} {group.Name,32} {group.Count,12}");
}

public class Email
{
    public string Subject { get; set; } = string.Empty;

    [Name("From: (Name)")]
    public string From_Name { get; set; } = string.Empty;

    [Name("From: (Address)")]
    public string From_Address { get; set; } = string.Empty;
}