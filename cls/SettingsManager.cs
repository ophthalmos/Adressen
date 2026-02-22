using System.Text.Json;
using System.Text.Encodings.Web;
using System.Text.Json.Serialization;

namespace Adressen.cls;

public class AppSettings
{
    public const int DatabaseSchemaVersion = 3; // Wird ignoriert (da const)
    public const int MaxRecentFiles = 10;      // kein JsonIgnore erforderlich

    [JsonIgnore]
    public static readonly int[] DefaultColumnWidths =
    [
        100, 100, 200, 100, 100, 100, 100, 100, 100, 100,
        100, 100, 100, 100, 100, 100, 100, 100, 100, 100,
        100, 100, 100, 100, 100, 100, 100
    ];

    [JsonIgnore]
    public static readonly bool[] DefaultHideColumns =
    [
        true, true, false, false, true, true, true, false, false,
        false, false, false, false, false, true, true, true, false,
        false, false, false, false, false, false, false, true, true
    ];

    // --- EIGENSCHAFTEN ---

    // Initialisierung direkt mit Klonen der Standardwerte!
    public bool[] HideColumnArr { get; set; } = (bool[])DefaultHideColumns.Clone();
    public int[] ColumnWidths { get; set; } = (int[])DefaultColumnWidths.Clone();

    // --- Sonstige Einstellungen ---
    public string PrintDevice { get; set; } = string.Empty;
    public string PrintSource { get; set; } = string.Empty;
    public bool PrintLandscape { get; set; } = false;
    public string PrintFormat { get; set; } = string.Empty;
    public string PrintFont { get; set; } = "Calibri, 12pt";
    public int SenderFontsize { get; set; } = 10;
    public int RecipientFontsize { get; set; } = 12;

    public int SenderIndex { get; set; } = 0;
    public string[] SenderLines1 { get; set; } = [];
    public string[] SenderLines2 { get; set; } = [];
    public string[] SenderLines3 { get; set; } = [];
    public string[] SenderLines4 { get; set; } = [];
    public string[] SenderLines5 { get; set; } = [];
    public string[] SenderLines6 { get; set; } = [];

    // Wrapper für DataBinding
    [JsonIgnore]
    public string SenderLines1Joined
    {
        get => string.Join(Environment.NewLine, SenderLines1); set => SenderLines1 = value.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);
    }
    [JsonIgnore]
    public string SenderLines2Joined
    {
        get => string.Join(Environment.NewLine, SenderLines2); set => SenderLines2 = value.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);
    }
    [JsonIgnore]
    public string SenderLines3Joined
    {
        get => string.Join(Environment.NewLine, SenderLines3); set => SenderLines3 = value.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);
    }
    [JsonIgnore]
    public string SenderLines4Joined
    {
        get => string.Join(Environment.NewLine, SenderLines4); set => SenderLines4 = value.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);
    }
    [JsonIgnore]
    public string SenderLines5Joined
    {
        get => string.Join(Environment.NewLine, SenderLines5); set => SenderLines5 = value.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);
    }
    [JsonIgnore]
    public string SenderLines6Joined
    {
        get => string.Join(Environment.NewLine, SenderLines6); set => SenderLines6 = value.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);
    }

    public bool PrintSender { get; set; } = true;
    public decimal RecipientOffsetX { get; set; } = 0m;
    public decimal RecipientOffsetY { get; set; } = 0m;
    public decimal SenderOffsetX { get; set; } = 0m;
    public decimal SenderOffsetY { get; set; } = 0m;

    public bool PrintRecipient { get; set; } = true;
    public bool PrintRecipientBold { get; set; } = false;
    public bool PrintSenderBold { get; set; } = false;
    public bool PrintRecipientSalutation { get; set; } = true;
    public bool RecipientSalutationAbove { get; set; } = true;
    public bool PrintRecipientCountry { get; set; } = false;
    public bool RecipientCountryUpper { get; set; } = false;

    public decimal LineHeightFactor { get; set; } = 1.2m;
    public decimal ZipGapFactor { get; set; } = 0.3m;
    public decimal LandGapFactor { get; set; } = 0.3m;

    public bool AskBeforeDelete { get; set; } = true;
    public string ColorScheme { get; set; } = "blue";
    public bool ContactsAutoload { get; set; } = false;
    public bool AskBeforeSaveSQL { get; set; } = false;
    public bool ReloadRecent { get; set; } = true;
    public bool NoAutoload { get; set; } = false;
    public string StandardFile { get; set; } = string.Empty;

    public bool DailyBackup { get; set; } = false;
    public bool AddZipBackup { get; set; } = false;
    public bool WatchFolder { get; set; } = false;
    public bool BackupSuccess { get; set; } = true;
    public decimal SuccessDuration { get; set; } = 2500;
    public string BackupDirectory { get; set; } = string.Empty;
    public string AddZipDirectory { get; set; } = string.Empty;
    public string DocumentFolder { get; set; } = string.Empty;
    public string DatabaseFolder { get; set; } = string.Empty;

    public int CopyPatternIndex { get; set; } = 0;
    public string[] CopyPattern1 { get; set; } = ["Anrede", "Praefix_Vorname_Zwischenname_Nachname", "Strasse", "PLZ_Ort"];
    public string[] CopyPattern2 { get; set; } = ["Telefon1", "Telefon2", "Mobil", "Fax"];
    public string[] CopyPattern3 { get; set; } = ["Mail1", "Mail2", "Internet"];
    public string[] CopyPattern4 { get; set; } = [];
    public string[] CopyPattern5 { get; set; } = [];
    public string[] CopyPattern6 { get; set; } = [];

    public int SplitterPosition { get; set; } = 249;
    public bool WindowMaximized { get; set; } = false;

    public int BirthdayRemindLimit { get; set; } = 14;
    public int BirthdayRemindAfter { get; set; } = 0;
    public bool BirthdayAddressShow { get; set; } = true;
    public bool BirthdayContactShow { get; set; } = true;

    public List<string> RecentFiles { get; set; } = [];
    public bool? WordProcessorProgram { get; set; } = null;

    public WindowPlacement? PrintWindowPosition
    {
        get; set;
    }
    public WindowPlacement? WindowPosition
    {
        get; set;
    }

    public int UpdateIndex { get; set; } = 3;  // Never

    public DateTime LastUpdateCheck { get; set; } = DateTime.MinValue; // Neues Feld

    public AppSettings DeepClone()
    {
        var json = JsonSerializer.Serialize(this);
        return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
    }

    // Stellt sicher, dass keine leeren Arrays existieren (wichtig nach dem Laden)
    //public void ValidateAndCorrect()
    //{
    //    if (ColumnWidths == null || ColumnWidths.Length == 0) { ColumnWidths = (int[])DefaultColumnWidths.Clone(); }
    //    if (HideColumnArr == null || HideColumnArr.Length == 0) { HideColumnArr = (bool[])DefaultHideColumns.Clone(); }
    //}

    public void ValidateAndCorrect()
    {
        // 1. Arrays absichern
        if (ColumnWidths == null || ColumnWidths.Length == 0) { ColumnWidths = (int[])DefaultColumnWidths.Clone(); }
        if (HideColumnArr == null || HideColumnArr.Length == 0) { HideColumnArr = (bool[])DefaultHideColumns.Clone(); }

        // 2. UNC-Pfade reparieren (einmalig beim Laden aus der JSON)
        StandardFile = Utils.CorrectUNC(StandardFile ?? string.Empty);
        BackupDirectory = Utils.CorrectUNC(BackupDirectory ?? string.Empty);
        AddZipDirectory = Utils.CorrectUNC(AddZipDirectory ?? string.Empty);
        DocumentFolder = Utils.CorrectUNC(DocumentFolder ?? string.Empty);
        DatabaseFolder = Utils.CorrectUNC(DatabaseFolder ?? string.Empty);

        // Auch die Liste der zuletzt geöffneten Dateien reparieren
        if (RecentFiles != null && RecentFiles.Count > 0)
        {
            for (var i = 0; i < RecentFiles.Count; i++) { RecentFiles[i] = Utils.CorrectUNC(RecentFiles[i]); }
        }
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
    private static readonly JsonSerializerOptions _options = new() { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };


    public static AppSettings Load(string filePath)  // Synchrone Methode für den Konstruktor
    {
        if (!File.Exists(filePath)) { return new AppSettings(); }
        try
        {
            using var stream = File.OpenRead(filePath);
            var settings = JsonSerializer.Deserialize<AppSettings>(stream, _options);

            if (settings != null)
            {
                settings.ValidateAndCorrect();
                return settings;
            }
            return new AppSettings();
        }
        catch { return new AppSettings(); }
    }

    //public static async Task<AppSettings> LoadAsync(string filePath)
    //{
    //    if (!File.Exists(filePath)) { return new AppSettings(); }
    //    try
    //    {
    //        await using var stream = File.OpenRead(filePath);
    //        var settings = await JsonSerializer.DeserializeAsync<AppSettings>(stream, _options);

    //        // Validierung: Wenn JSON Arrays leer waren, fülle sie auf
    //        if (settings != null) { settings.ValidateDefaults(); return settings; }

    //        return new AppSettings();
    //    }
    //    catch { return new AppSettings(); }
    //}

    public static void Save(AppSettings settings, string filePath)
    {
        try
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) { Directory.CreateDirectory(directory); }
            var tempPath = filePath + ".tmp";
            using (var stream = File.Create(tempPath)) { JsonSerializer.Serialize(stream, settings, _options); }
            File.Move(tempPath, filePath, overwrite: true);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Fehler beim Speichern: {ex.Message}");
            throw new IOException($"Einstellungen konnten nicht gespeichert werden: {ex.Message}", ex);
        }
    }
}