using System.Text.Json;

namespace Adressen.cls;
internal class AppSettings
{
    public string PrintDevice { get; set; } = string.Empty;
    public string PrintSource { get; set; } = string.Empty;
    public bool PrintLandscape { get; set; } = false;
    public string PrintFormat { get; set; } = string.Empty;
    public string PrintFont { get; set; } = "Calibri, 12pt"; // Beispiel Standardwert
    public int SenderFontsize { get; set; } = 10;
    public int RecipientFontsize { get; set; } = 12;
    public int SenderIndex { get; set; } = 0;
    public string[] SenderLines1 { get; set; } = [];
    public string[] SenderLines2 { get; set; } = [];
    public string[] SenderLines3 { get; set; } = [];
    public string[] SenderLines4 { get; set; } = [];
    public string[] SenderLines5 { get; set; } = [];
    public string[] SenderLines6 { get; set; } = [];
    public bool PrintSender { get; set; } = true;
    public decimal RecipientOffsetX { get; set; } = 0m;
    public decimal RecipientOffsetY { get; set; } = 0m;
    public decimal SenderOffsetX { get; set; } = 0m;
    public decimal SenderOffsetY { get; set; } = 0m;
    public bool PrintRecipientBold { get; set; } = false;
    public bool PrintSenderBold { get; set; } = false;
    public bool PrintRecipientSalutation { get; set; } = true;
    public bool PrintRecipientCountry { get; set; } = false;
    public bool AskBeforeDelete { get; set; } = true;
    public string ColorScheme { get; set; } = "blue";
    public bool ContactsAutoload { get; set; } = false;
    public bool AskBeforeSaveSQL { get; set; } = false;
    public bool ReloadRecent { get; set; } = true;
    public bool NoAutoload { get; set; } = false;
    public string StandardFile { get; set; } = string.Empty;
    public bool DailyBackup { get; set; } = false;
    public bool WatchFolder { get; set; } = false;
    public bool BackupSuccess { get; set; } = true;
    public decimal SuccessDuration { get; set; } = 2500;
    public string BackupDirectory { get; set; } = string.Empty;
    public string DocumentFolder { get; set; } = string.Empty;
    public string DatabaseFolder { get; set; } = string.Empty;
    public int CopyPatternIndex { get; set; } = 0;
    public string[] CopyPattern1 { get; set; } = ["Anrede", "Präfix_Vorname_Zwischenname_Nachname", "StraßeNr", "PLZ_Ort"];
    public string[] CopyPattern2 { get; set; } = ["Telefon1", "Telefon2", "Mobil", "Fax"];
    public string[] CopyPattern3 { get; set; } = ["Mail1", "Mail2", "Internet"];
    public string[] CopyPattern4 { get; set; } = [];
    public string[] CopyPattern5 { get; set; } = [];
    public string[] CopyPattern6 { get; set; } = [];
    public bool[] HideColumnArr { get; set; } = [true, true, false, false, true, true, true, false, false, false, false, false, true, true, true, false, false, false, false, false, false, false, false, true, true, true, true];
    public int SplitterPosition { get; set; } = 249;
    public bool WindowMaximized { get; set; } = false;
    public int[] ColumnWidths { get; set; } = [100, 100, 200, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100];
    public int BirthdayRemindLimit { get; set; } = 14;
    public int BirthdayRemindAfter { get; set; } = 0;
    public bool BirthdayAddressShow { get; set; } = true;
    public bool BirthdayContactShow { get; set; } = true;
    public List<string> RecentFiles { get; set; } = [];
    public bool? WordProcessorProgram { get; set; } = null;
    public WindowPlacement? WindowPosition
    {
        get; set;
    }
}

public class WindowPlacement
{
    public int X
    {
        get; set;
    }
    public int Y
    {
        get; set;
    }
    public int Width
    {
        get; set;
    }
    public int Height
    {
        get; set;
    }
}

internal static class SettingsManager
{
    private static readonly JsonSerializerOptions _options = new() { WriteIndented = true };

    public static async Task<AppSettings> LoadAsync(string filePath)
    {
        if (!File.Exists(filePath)) { return new AppSettings(); }
        try
        {
            await using var stream = File.OpenRead(filePath);
            var settings = await JsonSerializer.DeserializeAsync<AppSettings>(stream, _options); // Konvertierung String/Zahl/Bool erfolgt automatisch
            return settings ?? new AppSettings(); // gibt die geladenen Einstellungen oder eine neue Instanz bei einem Fehler zurück
        }
        catch { return new AppSettings(); } // Im Falle eines Fehlers (z.B. Datei korrupt), Standardwerte verwenden
    }

    public static void Save(AppSettings settings, string filePath)
    {
        try
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) { Directory.CreateDirectory(directory); }
            using var stream = File.Create(filePath);
            JsonSerializer.SerializeAsync(stream, settings, _options);
        }
        catch { }
    }
}