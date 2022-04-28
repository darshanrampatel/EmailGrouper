using CsvHelper;
using CsvHelper.Configuration.Attributes;
using Syroot.Windows.IO;
using System.Globalization;

var downloadFolderPath = new KnownFolder(KnownFolderType.Downloads).Path;
ProcessEmails("20220408_Emails.CSV", false);
ProcessEmails("20220408_Emails_SentMail.CSV", true);

void ProcessEmails(string file, bool sentEmails)
{
    Console.WriteLine($"Processing {file}");
    var emailsFile = Path.Combine(downloadFolderPath, file);
    using var reader = new StreamReader(emailsFile);
    using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
    var emails = csv.GetRecords<Email>().ToList();
    Console.WriteLine($"Found {emails.Count} emails");
    var groupedEmails = emails.GroupBy(e => new
    {
        Address = sentEmails ? e.To_Address : e.From_Address,
        Name = sentEmails ? e.To_Name : e.From_Name,
    }).Select(g => new { g.Key.Address, Count = g.Count(), g.Key.Name }).Where(g => g.Count > 5);
    Console.WriteLine($"Found {groupedEmails.Count()} groups");
    foreach (var group in groupedEmails.OrderByDescending(g => g.Count))
    {
        if (group.Address.Length > 64)
        {
            Console.WriteLine($"{group.Address.Substring(0,61),64} {group.Name.Count(c => c == ';'),64} {group.Count,12}");
        }
        else
        {
            Console.WriteLine($"{group.Address,64} {group.Name,64} {group.Count,12}");
        }
    }
    Console.WriteLine();
}

public class Email
{
    public string Subject { get; set; } = string.Empty;

    [Name("From: (Name)")]
    public string From_Name { get; set; } = string.Empty;

    [Name("From: (Address)")]
    public string From_Address { get; set; } = string.Empty;

    [Name("To: (Name)")]
    public string To_Name { get; set; } = string.Empty;

    [Name("To: (Address)")]
    public string To_Address { get; set; } = string.Empty;
}