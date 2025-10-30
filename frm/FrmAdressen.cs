using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Adressen.cls;
using Adressen.frm;
using Adressen.Properties;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Oauth2.v2;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;

namespace Adressen;

public partial class FrmAdressen : Form
{
    public bool[] HideColumnStd => hideColumnStd;

    //public delegate void SQLiteRowUpdatedEventHandler(object sender, SQLiteRowUpdatedEventHandler e);
    private readonly FrmSplashScreen? _splashScreen;
    private static readonly string appPath = Application.ExecutablePath; // EXE-Pfad
    private SQLiteConnection? _connection;
    private SQLiteDataAdapter? _adapter;
    private DataTable? _dataTable;
    //private BindingSource? _bindingSource;
    private string databaseFilePath = string.Empty; // Path.ChangeExtension(appPath, ".adb");
    private bool sAskBeforeSaveSQL = true; // false = Änderungen automatisch speichern
    private AppSettings _settings = new(); // Ein einziges Objekt für alle Einstellungen
    private readonly string _settingsPath;
    private readonly string tokenDir;
    private readonly string secretPath;
    private readonly string boysPath;
    private readonly string girlPath;
    private readonly string cleanRegex = @"[^\+0-9]";
    private readonly string appLong = Application.ProductName ?? "Adressen & Kontakte";
    private readonly string appName = "Adressen";
    private readonly string appCont = "Kontakte";
    private readonly Dictionary<string, string> addBookDict = [];
    private readonly Dictionary<Control, string> dictEditField = [];
    private string pDevice = string.Empty;
    private string pSource = string.Empty;
    private bool pLandscape = true;
    private string pFormat = string.Empty;
    private string pFont = "Calibri";
    private int pSenderSize = 12;
    private int pRecipSize = 14;
    private int pSenderIndex = 0;
    private string[] pSenderLines1 = [];
    private string[] pSenderLines2 = [];
    private string[] pSenderLines3 = [];
    private string[] pSenderLines4 = [];
    private string[] pSenderLines5 = [];
    private string[] pSenderLines6 = [];
    private bool pSenderPrint = false;
    private decimal pRecipX = 0;
    private decimal pRecipY = 0;
    private decimal pSendX = 0;
    private decimal pSendY = 0;
    private bool pRecipBold = false;
    private bool pSendBold = false;
    private bool pSalutation = false;
    private bool pCountry = true;
    private bool sAskBeforeDelete = true;
    private string sColorScheme = "blue";
    private bool sContactsAutoload = false;
    private bool sReloadRecent = false;
    private bool sNoAutoload = false;
    private string sStandardFile = string.Empty;
    private bool sDailyBackup = false;
    private bool sWatchFolder = false;
    private bool sBackupSuccess = true;
    private decimal sSuccessDuration = 2500;
    private string sBackupDirectory = string.Empty;
    private string sLetterDirectory = string.Empty;
    private string sDatabaseFolder = string.Empty;
    private int indexCopyPattern = 0;
    private string[] copyPattern1 = ["Anrede", "Präfix_Vorname_Zwischenname_Nachname", "StraßeNr", "PLZ_Ort"];
    private string[] copyPattern2 = ["Telefon1", "Telefon2", "Mobil", "Fax"];
    private string[] copyPattern3 = ["Mail1", "Mail2", "Internet"];
    private string[] copyPattern4 = [];
    private string[] copyPattern5 = [];
    private string[] copyPattern6 = [];
    private const int latestSchemaVersion = 1; // DB-Ziel-Version: muss bei jeder zukünftigen Änderung an der Datenbankstruktur erhöht werden!!
    private readonly string[] dataFields = ["Anrede", "Präfix", "Nachname", "Vorname", "Zwischenname", "Nickname",
        "Suffix", "Firma", "Straße", "PLZ", "Ort", "Land", "Betreff", "Grußformel", "Schlussformel", "Geburtstag",
        "Mail1", "Mail2", "Telefon1", "Telefon2", "Mobil", "Fax", "Internet", "Notizen", "Gruppen", "Dokumente"]; // Id fehlt absichtlich  
    private bool[] hideColumnArr = new bool[27]; // muss angepasst werden, wenn Felder/Spalten hinzugefügt werden
    private readonly bool[] hideColumnStd = [true, true, false, false, true, true, true, false, false, false, false, false, true, true, true, false, false, false, false, false, false, false, false, true, true, true, true]; // muss angepasst werden, wenn Felder/Spalten hinzugefügt werden
    private int[] columnWidths = [100, 100, 200, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100];
    private int splitterPosition;
    private WindowPlacement? windowPosition;
    private bool windowMaximized = false;
    private readonly bool argsPath = false;
    private Word.Document? wordDoc;
    private dynamic? wordApp; // Word.Application
    private int contactNewRowIndex = -1;
    private readonly Dictionary<string, string> originalAddressData = [];
    private readonly Dictionary<string, string> changedAddressData = [];
    private readonly Dictionary<string, string> originalContactData = [];
    private readonly Dictionary<string, string> changedContactData = [];
    private int prevSelectedAddressRowIndex = -1;
    private int prevSelectedContactRowIndex = -1;
    private bool isSelectionChanging = false;
    private int birthdayRemindLimit = 30;
    private int birthdayRemindAfter = 3;
    private bool birthdayAddressShow = false;
    private bool birthdayContactShow = false;
    private bool ignoreTextChange = false; // ignore when changing text in AddressEditFields/ContactEditFields
    private bool ignoreSearchChange = false;
    private string lastAddressSearch = string.Empty;
    private string lastContactSearch = string.Empty;
    private ToolStripDropDown? calendarDropdown;
    private MonthCalendar? monthCalendar;
    private bool textBoxClicked = false;
    private readonly int maxRecentFiles = 10;
    private List<string> recentFiles = [];
    private bool? sWordProcProg = null; // null=JedesMalfragen, true=MS-Word, false=LibreOffice
    private readonly string[] formats = ["dd.MM.yyyy", "d.MM.yyyy", "dd.M.yyyy", "d.M.yyyy", "dd.M.yy", "d.MM.yy", "d.M.yy"];
    private readonly CultureInfo culture = new("de-DE");
    private string lastSearchText = string.Empty;
    private TabPage? deactivatedPage = null;
    private List<ListViewItem> allDokuLVItems = [];
    private int lastColumn = -1;
    private SortOrder lastOrder = SortOrder.None;
    private string lastTooltipText = string.Empty;
    private bool birthdayShow = true; // false wenn Zugriffstoken für Google-Kontakte fehlt oder abgelaufen ist
    private readonly string[] documentTypes = ["*.doc", "*.dot", "*.docx", "*.doct", "*.docm", "*.odt", "*.ott", "*.fodt", "*.uot", "*.pdf", "*.txt"];
    private List<string> addressCbItems_Anrede = [];
    private List<string> addressCbItems_Präfix = [];
    private List<string> addressCbItems_PLZ = [];
    private List<string> addressCbItems_Ort = [];
    private List<string> addressCbItems_Land = [];
    private List<string> addressCbItems_Schlussformel = [];
    private List<string> contactCbItems_Anrede = [];
    private List<string> contactCbItems_Präfix = [];
    private List<string> contactCbItems_PLZ = [];
    private List<string> contactCbItems_Ort = [];
    private List<string> contactCbItems_Land = [];
    private List<string> contactCbItems_Schlussformel = [];
    private readonly List<string> männlichGrusse =
        [
        "Hallo #vorname",
        "Hallo #nickname",
        "Lieber #vorname",
        "Lieber #nickname",
        "Lieber Herr #nachname",
        "Sehr geehrter Herr #nachname",
        "Sehr geehrter Herr #titel #nachname",
        "Sehr geehrte Kollege #nachname",
        "Sehr geehrte Kollege #titel #nachname",
        "Sehr geehrte Damen und Herren"
        ];
    private readonly List<string> weiblichGrusse =
    [
        "Hallo #vorname",
        "Hallo #nickname",
        "Liebe #vorname",
        "Liebe #nickname",
        "Liebe Frau #nachname",
        "Sehr geehrte Frau #nachname",
        "Sehr geehrte Frau #titel #nachname",
        "Sehr geehrte Kollegin #nachname",
        "Sehr geehrte Kollegin #titel #nachname",
        "Sehr geehrte Damen und Herren"
    ];
    private static readonly Dictionary<string, bool> nameGenderMap = new(StringComparer.OrdinalIgnoreCase);
    private readonly string[] pictureBoxExtensions = [".bmp", ".jpg", ".jpeg", ".png", ".gif"];
    private readonly SortedSet<string> allAddressMemberships = [];
    private readonly SortedSet<string> allContactMemberships = [];
    private SortedSet<string> curAddressMemberships = [];
    private SortedSet<string> curContactMemberships = [];
    private readonly Dictionary<string, string> contactGroupsDict = [];
    private static readonly HashSet<string> excludedGroups = ["myContacts", "all", "blocked", "chatBuddies", "coworkers", "family", "friends"];
    private string userEmail = string.Empty;

    public FrmAdressen(FrmSplashScreen? splashScreen, string[] args)
    {
        if (args.Length >= 1)
        {
            if (File.Exists((string?)args[0]))
            {
                databaseFilePath = (string?)args[0] ?? string.Empty;
                if (!string.IsNullOrEmpty(databaseFilePath)) { argsPath = true; }
            }
        }

        InitializeComponent();
        _splashScreen = splashScreen;  // Splash Screen speichern um ihn beenden zu können (s. Load Event)  
        typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, addressDGV, [true]);
        typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, contactDGV, [true]);

        imageList.Images.Add(Resources.address_book);
        imageList.Images.Add(Resources.address_book_blue);
        imageList.Images.Add(Resources.universal24);
        imageList.Images.Add(Resources.inbox24);
        imageList.Images.Add(Resources.inboxdoc24);
        tabControl.ImageList = imageList; // Bilder zur Laufzeit aus Projekt-Ressourcen laden, vermeidet BinaryFormatter
        tabControl.TabPages[0].ImageIndex = 0;
        tabControl.TabPages[1].ImageIndex = 1;
        tabulation.TabPages[0].ImageIndex = 2;
        tabulation.TabPages[1].ImageIndex = 3;

        if (Utilities.IsInnoSetupValid(Path.GetDirectoryName(appPath)!))
        {
            _settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName, appName + ".json");
            tokenDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName, "token.json");
            secretPath = Path.Combine(Path.GetDirectoryName(appPath) ?? string.Empty, "client_secret.json");
        }
        else
        {
            _settingsPath = Path.ChangeExtension(appPath, ".json");
            tokenDir = Path.Combine(AppContext.BaseDirectory, "token.json"); ;
            secretPath = Path.Combine(AppContext.BaseDirectory, "client_secret.json");
        }
        boysPath = Path.Combine(Path.GetDirectoryName(_settingsPath) ?? string.Empty, "MännlicheVornamen.txt");
        girlPath = Path.Combine(Path.GetDirectoryName(_settingsPath) ?? string.Empty, "WeiblicheVornamen.txt");

        addressDGV.ColumnHeadersDefaultCellStyle.SelectionBackColor = addressDGV.ColumnHeadersDefaultCellStyle.BackColor;
        contactDGV.ColumnHeadersDefaultCellStyle.SelectionBackColor = contactDGV.ColumnHeadersDefaultCellStyle.BackColor;
        searchTSTextBox.TextBox.PlaceholderText = " Suche";
        splitterPosition = splitContainer.SplitterDistance;

        string[] extraFields = ["Präfix_Zwischenname_Nachname", "Vorname_Zwischenname_Nachname", "Präfix_Vorname_Zwischenname_Nachname",
            "Anrede_Präfix_Vorname_Zwischenname_Nachname", "StraßeNr", "PLZ_Ort"];
        foreach (var field in dataFields.SkipLast(2).Concat(extraFields)) { addBookDict[field] = string.Empty; } // Notizen und Dokumente nicht 

        dictEditField.Add(cbAnrede, "Anrede");
        dictEditField.Add(cbPräfix, "Präfix");
        dictEditField.Add(tbNachname, "Nachname");
        dictEditField.Add(tbVorname, "Vorname");
        dictEditField.Add(tbZwischenname, "Zwischenname");
        dictEditField.Add(tbNickname, "Nickname");
        dictEditField.Add(tbSuffix, "Suffix");
        dictEditField.Add(tbFirma, "Firma");
        dictEditField.Add(tbStraße, "Straße");
        dictEditField.Add(cbPLZ, "PLZ");
        dictEditField.Add(cbOrt, "Ort");
        dictEditField.Add(cbLand, "Land");
        dictEditField.Add(tbBetreff, "Betreff");
        dictEditField.Add(cbGrußformel, "Grußformel");
        dictEditField.Add(cbSchlussformel, "Schlussformel");
        dictEditField.Add(tbMail1, "Mail1");
        dictEditField.Add(tbMail2, "Mail2");
        dictEditField.Add(tbTelefon1, "Telefon1");
        dictEditField.Add(tbTelefon2, "Telefon2");
        dictEditField.Add(tbMobil, "Mobil");
        dictEditField.Add(tbFax, "Fax");
        dictEditField.Add(tbInternet, "Internet");

        fileToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
        editToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
        viewToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
        extraToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
        helpToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
    }

    private async void FrmAdressen_Load(object sender, EventArgs e)
    {
        if (File.Exists(_settingsPath))
        {
            await LoadConfiguration();
        }
        else { Directory.CreateDirectory(Path.GetDirectoryName(_settingsPath)!); } // If the folder exists already, the line will be ignored.     
        databaseFilePath = argsPath ? databaseFilePath : recentFiles.Count > 0 ? recentFiles[0] : string.Empty;

        if (!(new int[] { hideColumnArr.Length, hideColumnStd.Length, columnWidths.Length }).All(len => len == dataFields.Length + 1))
        {
            var text = $"Datenfelder: {dataFields.Length + 1}\nhideColumnArr: {hideColumnArr.Length}\nhideColumnStd: {hideColumnStd.Length}\ncolumnWidths: {columnWidths.Length}";
            Utilities.ErrorMsgTaskDlg(Handle, "Fehler bei der Initialisierung", "Nicht alle Arrays haben die gleiche Länge.\n" + text);
        }
        RestoreWindowPlacement();

        try
        {
            if (File.Exists(girlPath))
            {
                var girlNames = await File.ReadAllLinesAsync(girlPath);
                foreach (var name in girlNames)
                {
                    var trimmedName = name.Trim();
                    if (!string.IsNullOrEmpty(trimmedName)) { nameGenderMap[trimmedName] = true; }
                }
            }

            if (File.Exists(boysPath))
            {
                var boyNames = await File.ReadAllLinesAsync(boysPath);
                foreach (var name in boyNames)
                {
                    var trimmedName = name.Trim();
                    if (!string.IsNullOrEmpty(trimmedName)) { nameGenderMap[trimmedName] = false; }
                }
            }
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "Fehler beim Laden der Namenslisten", ex.Message); }

        NativeMethods.SendMessage(searchTSTextBox.TextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 4 << 16);
        NativeMethods.SendMessage(searchTSTextBox.TextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 4);
        NativeMethods.SendMessage(tbNotizen.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 4 << 16);
        NativeMethods.SendMessage(tbNotizen.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 4);
        NativeMethods.SendMessage(maskedTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 4 << 16);
        NativeMethods.SendMessage(maskedTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 4);
        _ = NativeMethods.SendMessage(maskedTextBox.Handle, NativeMethods.EM_SETCUEBANNER, 0, "TT.MM.JJJJ");

        SetColorScheme();
        if ((sReloadRecent || argsPath) && !string.IsNullOrEmpty(databaseFilePath)) { ConnectSQLDatabase(databaseFilePath); }
        else if (!sReloadRecent && !sNoAutoload && !string.IsNullOrEmpty(sStandardFile)) { ConnectSQLDatabase(sStandardFile); }
        tsClearLabel.Visible = false;
        fileSystemWatcher.Path = sLetterDirectory;
        fileSystemWatcher.IncludeSubdirectories = true;
        fileSystemWatcher.Filters.Clear();
        foreach (var pattern in documentTypes) { fileSystemWatcher.Filters.Add(pattern); }
        if (sWatchFolder && !string.IsNullOrEmpty(sLetterDirectory) && Directory.Exists(sLetterDirectory)) { fileSystemWatcher.EnableRaisingEvents = true; }
        else { fileSystemWatcher.EnableRaisingEvents = false; } // stellt sich im Inspector ständig von allein auf true 

        if (_splashScreen != null)
        {
            _splashScreen.Close();
            _splashScreen.Dispose();
        }
        Enabled = true; // UI wieder aktivieren
        Opacity = 1.0; // Jetzt, wo alles bereit ist, das Formular sichtbar machen.
        searchTSTextBox.TextBox.Focus(); // funktioniert nur, wenn SplashScreen weg ist 
    }

    private void RestoreWindowPlacement()
    {
        if (_settings.WindowMaximized)
        {
            WindowState = FormWindowState.Maximized;
            return;
        }
        if (_settings.WindowPosition == null) { return; } // Standardwerte verwenden  
        WindowState = FormWindowState.Normal;
        var savedPlacement = _settings.WindowPosition;
        var screen = Screen.FromPoint(new Point(savedPlacement.X, savedPlacement.Y)); // Screen.FromPoint statt Screen.PrimaryScreen.
        var workingArea = screen.WorkingArea; // Der Bereich ohne Taskleiste
        var width = Math.Max(savedPlacement.Width, MinimumSize.Width);
        var height = Math.Max(savedPlacement.Height, MinimumSize.Height);
        width = Math.Min(width, workingArea.Width);
        height = Math.Min(height, workingArea.Height);
        var x = savedPlacement.X;
        var y = savedPlacement.Y;
        if (x + width > workingArea.Right) { x = workingArea.Right - width; }  // Passt die X-Position an, falls das Fenster nach rechts aus dem Bildschirm ragt.
        if (x < workingArea.Left) { x = workingArea.Left; }  // Passt an, falls es nach links aus dem Bildschirm ragt (oder im negativen Bereich ist).
        if (y + height > workingArea.Bottom) { y = workingArea.Bottom - height; }  // Passt die Y-Position an, falls das Fenster nach unten aus dem Bildschirm ragt.
        if (y < workingArea.Top) { y = workingArea.Top; } // Passt an, falls es nach oben aus dem Bildschirm ragt.
        Location = new Point(x, y);
        Size = new Size(width, height);
    }
    private async Task LoadConfiguration()
    {
        _settings = await SettingsManager.LoadAsync(_settingsPath);
        pDevice = _settings.PrintDevice;
        pSource = _settings.PrintSource;
        pLandscape = _settings.PrintLandscape; // bool ist direkt bool
        pFormat = _settings.PrintFormat;
        pFont = _settings.PrintFont;
        pSenderSize = _settings.SenderFontsize; // int ist direkt int
        pRecipSize = _settings.RecipientFontsize;
        pSenderIndex = _settings.SenderIndex;
        pSenderLines1 = _settings.SenderLines1;
        pSenderLines2 = _settings.SenderLines2;
        pSenderLines3 = _settings.SenderLines3;
        pSenderLines4 = _settings.SenderLines4;
        pSenderLines5 = _settings.SenderLines5;
        pSenderLines6 = _settings.SenderLines6;
        pSenderPrint = _settings.PrintSender;
        pRecipX = _settings.RecipientOffsetX; // decimal ist direkt decimal
        pRecipY = _settings.RecipientOffsetY;
        pSendX = _settings.SenderOffsetX;
        pSendY = _settings.SenderOffsetY;
        pRecipBold = _settings.PrintRecipientBold;
        pSendBold = _settings.PrintSenderBold;
        pSalutation = _settings.PrintRecipientSalutation;
        pCountry = _settings.PrintRecipientCountry;
        sAskBeforeDelete = _settings.AskBeforeDelete;
        sColorScheme = _settings.ColorScheme;
        sContactsAutoload = _settings.ContactsAutoload;
        sAskBeforeSaveSQL = _settings.AskBeforeSaveSQL;
        sReloadRecent = _settings.ReloadRecent;
        sNoAutoload = _settings.NoAutoload;
        sStandardFile = _settings.StandardFile;
        sDailyBackup = _settings.DailyBackup;
        sWatchFolder = _settings.WatchFolder;
        sBackupSuccess = _settings.BackupSuccess;
        sSuccessDuration = _settings.SuccessDuration;
        sBackupDirectory = _settings.BackupDirectory;
        sLetterDirectory = _settings.DocumentFolder;
        sDatabaseFolder = _settings.DatabaseFolder;
        indexCopyPattern = _settings.CopyPatternIndex;
        copyPattern1 = _settings.CopyPattern1;
        copyPattern2 = _settings.CopyPattern2;
        copyPattern3 = _settings.CopyPattern3;
        copyPattern4 = _settings.CopyPattern4;
        copyPattern5 = _settings.CopyPattern5;
        copyPattern6 = _settings.CopyPattern6;
        hideColumnArr = _settings.HideColumnArr.Length > 0 ? _settings.HideColumnArr : hideColumnArr;
        splitterPosition = _settings.SplitterPosition;
        windowMaximized = _settings.WindowMaximized;
        windowPosition = _settings.WindowPosition;
        columnWidths = _settings.ColumnWidths.Length > 0 ? _settings.ColumnWidths : columnWidths;
        birthdayRemindLimit = _settings.BirthdayRemindLimit;
        birthdayRemindAfter = _settings.BirthdayRemindAfter;
        birthdayAddressShow = _settings.BirthdayAddressShow;
        birthdayContactShow = _settings.BirthdayContactShow;
        recentFiles = _settings.RecentFiles;
        sWordProcProg = _settings.WordProcessorProgram;
    }

    private void SaveConfiguration()
    {
        _settings.PrintDevice = pDevice;
        _settings.PrintSource = pSource;
        _settings.PrintLandscape = pLandscape;
        _settings.PrintFormat = pFormat;
        _settings.PrintFont = pFont;
        _settings.SenderFontsize = pSenderSize;
        _settings.RecipientFontsize = pRecipSize;
        _settings.SenderIndex = pSenderIndex;
        _settings.SenderLines1 = pSenderLines1;
        _settings.SenderLines2 = pSenderLines2;
        _settings.SenderLines3 = pSenderLines3;
        _settings.SenderLines4 = pSenderLines4;
        _settings.SenderLines5 = pSenderLines5;
        _settings.SenderLines6 = pSenderLines6;
        _settings.PrintSender = pSenderPrint;
        _settings.RecipientOffsetX = pRecipX;
        _settings.RecipientOffsetY = pRecipY;
        _settings.SenderOffsetX = pSendX;
        _settings.SenderOffsetY = pSendY;
        _settings.PrintRecipientBold = pRecipBold;
        _settings.PrintSenderBold = pSendBold;
        _settings.PrintRecipientSalutation = pSalutation;
        _settings.PrintRecipientCountry = pCountry;
        _settings.AskBeforeDelete = sAskBeforeDelete;
        _settings.ColorScheme = sColorScheme;
        _settings.ContactsAutoload = sContactsAutoload;
        _settings.AskBeforeSaveSQL = sAskBeforeSaveSQL;
        _settings.ReloadRecent = sReloadRecent;
        _settings.NoAutoload = sNoAutoload;
        _settings.StandardFile = sStandardFile;
        _settings.DailyBackup = sDailyBackup;
        _settings.WatchFolder = sWatchFolder;
        _settings.BackupSuccess = sBackupSuccess;
        _settings.SuccessDuration = sSuccessDuration;
        _settings.BackupDirectory = sBackupDirectory;
        _settings.DocumentFolder = sLetterDirectory;
        _settings.DatabaseFolder = sDatabaseFolder;
        _settings.CopyPatternIndex = indexCopyPattern;
        _settings.CopyPattern1 = copyPattern1;
        _settings.CopyPattern2 = copyPattern2;
        _settings.CopyPattern3 = copyPattern3;
        _settings.CopyPattern4 = copyPattern4;
        _settings.CopyPattern5 = copyPattern5;
        _settings.CopyPattern6 = copyPattern6;
        _settings.WindowMaximized = WindowState == FormWindowState.Maximized;
        var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        _settings.WindowPosition = new WindowPlacement { X = bounds.X, Y = bounds.Y, Width = bounds.Width, Height = bounds.Height };
        _settings.SplitterPosition = splitContainer.SplitterDistance;
        _settings.HideColumnArr = hideColumnArr;
        _settings.BirthdayRemindLimit = birthdayRemindLimit;
        _settings.BirthdayRemindAfter = birthdayRemindAfter;
        _settings.BirthdayAddressShow = birthdayAddressShow;
        _settings.BirthdayContactShow = birthdayContactShow;
        _settings.RecentFiles = recentFiles;
        _settings.WordProcessorProgram = sWordProcProg;
        if (tabControl.SelectedTab == contactTabPage) { _settings.ColumnWidths = [.. contactDGV.Columns.Cast<DataGridViewColumn>().Select(c => c.Width)]; }
        else if (tabControl.SelectedTab == addressTabPage) { _settings.ColumnWidths = [.. addressDGV.Columns.Cast<DataGridViewColumn>().Select(c => c.Width)]; }
        SettingsManager.Save(_settings, _settingsPath);
    }

    private async void FrmAdressen_Shown(object sender, EventArgs e)
    {
        splitContainer.SplitterDistance = splitterPosition;
        flexiTSStatusLabel.Width = 244 + splitContainer.SplitterDistance - 536;
        if (birthdayAddressShow) { BirthdayReminder(); }
        if (sContactsAutoload) { await LoadAndDisplayGoogleContactsAsync(); }
    }

    private void ConnectSQLDatabase(string file)
    {
        flexiTSStatusLabel.Text = string.Empty;
        if (string.IsNullOrEmpty(file) || !File.Exists(file))
        {
            Utilities.ErrorMsgTaskDlg(Handle, "Datenbank-Datei nicht gefunden", file, TaskDialogIcon.ShieldWarningYellowBar);
            recentFiles.Remove(file); // erzeugt keinen Fehler, wenn file nicht in der Liste ist - gibt dann nur false zurück
            return;
        }
        try
        {
            databaseFilePath = file;
            if (databaseFilePath.StartsWith(@"\\")) { databaseFilePath = @"\\\\" + databaseFilePath.TrimStart('\\'); } // Workaround für "SQLiteException: unable to open database file if UNC-Path
            _connection = new SQLiteConnection($"Data Source={databaseFilePath};FailIfMissing=True"); // new SQLiteConnection("Data Source=adressen.db;Version=3;");
            _connection.Open();
            MigrateDatabase(_connection); // Datenbank-Migration, wenn nötig

            _adapter = new SQLiteDataAdapter("SELECT * FROM Adressen", _connection) { MissingSchemaAction = MissingSchemaAction.AddWithKey }; // stellt u.a. die AutoIncrement-Eigenschaft der Id-Spalte in der DataTable ein
            var builder = new SQLiteCommandBuilder(_adapter) { ConflictOption = ConflictOption.OverwriteChanges };
            _adapter.UpdateCommand = builder.GetUpdateCommand();
            _adapter.DeleteCommand = builder.GetDeleteCommand();
            _adapter.InsertCommand = builder.GetInsertCommand();
            _adapter.InsertCommand.CommandText += "; SELECT LAST_INSERT_ROWID();";
            _adapter.InsertCommand.UpdatedRowSource = UpdateRowSource.FirstReturnedRecord;

            _dataTable = new DataTable(); // Create a DataTable to hold the data
            _adapter.Fill(_dataTable);           // Use the Fill method to retrieve data into the DataTable
            var sortedRows = from row in _dataTable.AsEnumerable() orderby row.Field<string>("Vorname") ascending orderby row.Field<string>("Nachname") ascending select row; // alphabetisch 
            using (var sortedDT = _dataTable.Clone())
            {
                foreach (var row in sortedRows) { sortedDT.ImportRow(row); }
                _dataTable = sortedDT;
            }
            addressDGV.SuspendLayout();

            addressDGV.DataSource = _dataTable;
            //_bindingSource = new BindingSource { DataSource = _dataTable };
            //addressDGV.DataSource = _bindingSource;

            Utilities.SetColumnWidths(columnWidths, addressDGV);
            foreach (DataGridViewColumn column in addressDGV.Columns) { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
            if (addressDGV.Rows.Count > 0)
            {
                var emptyRows = from DataGridViewRow row in addressDGV.Rows.Cast<DataGridViewRow>()
                                .Where(row => row.Cells.Cast<DataGridViewCell>().SkipLast(2)
                                .All(cell => cell.Value == null || string.IsNullOrEmpty(cell.Value.ToString())))
                                select row;
                foreach (var emptyRow in emptyRows) // //if (emptyRows.Any())
                {
                    if (emptyRow.DataBoundItem is DataRowView drv) { drv.Row.Delete(); }
                }
                using var changes = _dataTable.GetChanges(DataRowState.Deleted);
                if (changes != null) { _adapter.Update(changes); }
            }
            _dataTable.AcceptChanges();
            PopulateMemberships(); // allAddressMemberships auffüllen
            //}
            addressDGV.ResumeLayout(false);

            cbAnrede.Items.Clear();
            cbPräfix.Items.Clear();
            cbPLZ.Items.Clear();
            cbOrt.Items.Clear();
            cbLand.Items.Clear();
            cbSchlussformel.Items.Clear();
            cbAnrede.Items.AddRange([.. _dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Anrede")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbPräfix.Items.AddRange([.. _dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Präfix")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbPLZ.Items.AddRange([.. _dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("PLZ")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbOrt.Items.AddRange([.. _dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Ort")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbLand.Items.AddRange([.. _dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Land")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbSchlussformel.Items.AddRange([.. _dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Schlussformel")!).Where(value => !string.IsNullOrEmpty(value)).Distinct()]);
            addressCbItems_Anrede = [.. cbAnrede.Items.Cast<string>()];
            addressCbItems_Präfix = [.. cbPräfix.Items.Cast<string>()];
            addressCbItems_PLZ = [.. cbPLZ.Items.Cast<string>()];
            addressCbItems_Ort = [.. cbOrt.Items.Cast<string>()];
            addressCbItems_Land = [.. cbLand.Items.Cast<string>()];
            addressCbItems_Schlussformel = [.. cbSchlussformel.Items.Cast<string>()];

            recentFiles.Remove(databaseFilePath);
            recentFiles.Insert(0, databaseFilePath);
            if (recentFiles.Count > maxRecentFiles) { recentFiles = [.. recentFiles.Take(maxRecentFiles)]; }

            newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled
                = deleteTSButton.Enabled = newToolStripMenuItem.Enabled = newTSButton.Enabled = duplicateToolStripMenuItem.Enabled = copyTSButton.Enabled = wordTSButton.Enabled
                = envelopeTSButton.Enabled = true;
            copyToOtherDGVTSMenuItem.Enabled = false;
            tabControl.SelectTab(0);
            AddressEditFields(_dataTable.Rows.Count > 0 ? 0 : -1);
            searchTSTextBox.Focus(); //MessageBox.Show(string.Join(Environment.NewLine, [.. dataTable.Columns.Cast<DataColumn>().Select(x => x.ColumnName)]));
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "ConnectSQLDatabase: " + ex.GetType().ToString(), ex.Message); }
    }

    private void OnAdapterRowUpdated(object? sender, System.Data.Common.RowUpdatedEventArgs e)
    {
        if (e.StatementType == StatementType.Insert && e.Status == UpdateStatus.Continue)
        {
            if (_connection != null && _connection.State == ConnectionState.Open)
            {
                try
                {
                    using var cmd = new SQLiteCommand("SELECT LAST_INSERT_ROWID()", _connection);
                    var newId = Convert.ToInt64(cmd.ExecuteScalar());  // newId abrufen (sicherer mit Convert)
                    e.Row["Id"] = newId; // Die 'Id'-Spalte der DataRow in der DataTable aktualisieren
                    e.Row.AcceptChanges(); // SEHR WICHTIG: Die Änderungen für DIESE Zeile akzeptieren
                }
                catch (Exception ex)
                {
                    Utilities.ErrorMsgTaskDlg(Handle, "Fehler beim Abrufen der neuen ID", ex.Message);
                    e.Row.RejectChanges();
                    e.Status = UpdateStatus.SkipCurrentRow;
                }
            }
        }
    }

    private void MigrateDatabase(SQLiteConnection connection)
    {
        long currentVersion;
        using (var cmd = new SQLiteCommand("PRAGMA user_version;", connection)) { currentVersion = (long)cmd.ExecuteScalar(); }
        if (currentVersion >= latestSchemaVersion) { return; }
        if (currentVersion < 1) // "hart" codieren (nicht latestSchemaVersion)! => bei neuen Versinen hier weitere if-Abfragen ergänzen!
        {
            using var transaction = connection.BeginTransaction(); // Migration von Version 0 auf Version 1
            try
            {
                var existingColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using (var cmd = new SQLiteCommand("PRAGMA table_info(Adressen);", connection)) // fehlende Spalten finden
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read()) { existingColumns.Add(reader.GetString(1)); } // Der Spaltenname steht an Index 1
                }
                foreach (var field in dataFields)
                {
                    if (!existingColumns.Contains(field))
                    {
                        var alterQuery = $"ALTER TABLE Adressen ADD COLUMN {field} TEXT;";  // Spalte fehlt, also hinzufügen
                        using var alterCmd = new SQLiteCommand(alterQuery, connection);
                        alterCmd.ExecuteNonQuery();
                    }
                }
                var createPhotoQuery = @"
                    CREATE TABLE IF NOT EXISTS Fotos (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        AdressId INTEGER NOT NULL UNIQUE,
                        Fotodaten BLOB, 
                        FOREIGN KEY(AdressId) REFERENCES Adressen(Id) ON DELETE CASCADE
                    );";
                using (var createCmd = new SQLiteCommand(createPhotoQuery, connection)) { createCmd.ExecuteNonQuery(); }
                using (var updateVersionCmd = new SQLiteCommand($"PRAGMA user_version = {latestSchemaVersion};", connection)) { updateVersionCmd.ExecuteNonQuery(); }  // Datenbankversion heraufsetzen
                transaction.Commit();
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                Utilities.ErrorMsgTaskDlg(Handle, "Fehler bei der Datenbank-Migration", ex.Message);
            }
        }
    }

    private void CreateNewDatabase(string filePath, bool addSampleRecord = false)
    {
        try
        {
            if (File.Exists(filePath)) { File.Delete(filePath); }
            //using var connection = new SQLiteConnection($"Data Source={filePath};FailIfMissing=False;");
            _connection = new SQLiteConnection($"Data Source={filePath};FailIfMissing=False;");
            _connection.Open();
            var columnDefinitions = string.Join(", ", dataFields.Select(field => $"{field} TEXT"));
            var createTableQuery = $@"CREATE TABLE Adressen ({columnDefinitions}, Id INTEGER PRIMARY KEY AUTOINCREMENT)";
            using (var command = new SQLiteCommand(createTableQuery, _connection)) { command.ExecuteNonQuery(); }
            var createPhotoQuery = @"
                CREATE TABLE Fotos (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                    AdressId INTEGER NOT NULL UNIQUE,
                    Fotodaten BLOB, 
                    FOREIGN KEY(AdressId) REFERENCES Adressen(Id) ON DELETE CASCADE
                );";
            using (var command = new SQLiteCommand(createPhotoQuery, _connection)) { command.ExecuteNonQuery(); }


            if (addSampleRecord)  // Beispieldatensatz einfügen, falls gewünscht
            {
                var insertQuery = "INSERT INTO Adressen (Anrede, Präfix, Nachname, Vorname, Zwischenname, Nickname, Suffix, Straße, PLZ, Ort, Grußformel, Geburtstag, Mail1) " +
                                  "VALUES (@Anrede, @Präfix, @Nachname, @Vorname, @Zwischenname, @Nickname, @Suffix, @Straße, @Plz, @Ort, @Grußformel, @Geburtstag, @Mail1)";
                using var command = new SQLiteCommand(insertQuery, _connection);
                command.Parameters.AddWithValue("@Anrede", "Herrn");
                command.Parameters.AddWithValue("@Präfix", "Dr. h.c.");
                command.Parameters.AddWithValue("@Nachname", "Mustermann");
                command.Parameters.AddWithValue("@Vorname", "Max");
                command.Parameters.AddWithValue("@Zwischenname", "von und zu");
                command.Parameters.AddWithValue("@Nickname", "Maxi");
                command.Parameters.AddWithValue("@Suffix", "Jr. MBA");
                command.Parameters.AddWithValue("@Straße", "Langer Weg 144");
                command.Parameters.AddWithValue("@Plz", "01234");
                command.Parameters.AddWithValue("@Ort", "Entenhausen");
                command.Parameters.AddWithValue("@Grußformel", "Lieber Max");
                command.Parameters.AddWithValue("@Geburtstag", "6.3.1995");
                command.Parameters.AddWithValue("@Mail1", "abc@xyz.com");
                command.ExecuteNonQuery();
            }

            using var cmd = new SQLiteCommand($"PRAGMA user_version = {latestSchemaVersion}", _connection);
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "Fehler beim Erstellen der neuen Datenbank", ex.Message); }
    }

    private void PopulateMemberships()
    {
        const string columnName = "Gruppen";
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        if (_dataTable is null || !_dataTable.Columns.Contains(columnName)) { return; }
        allAddressMemberships.Clear();
        allAddressMemberships.Add("★"); // Favoriten immer als erstes Element hinzufügen
        foreach (var jsonString in _dataTable.AsEnumerable().Select(row => row[columnName] as string).Where(s => !string.IsNullOrWhiteSpace(s)))
        {
            allAddressMemberships.UnionWith(Utilities.DeserializeGroups(Handle, jsonString!, options));
        }
    }

    private void SaveSQLDatabase(bool closeDB = false, bool askNever = false, bool isFormClosing = false)
    {
        if (_dataTable == null || _adapter == null)
        {
            if (closeDB) { CloseDatabaseConnection(); }
            return;
        }
        using (var changes = _dataTable.GetChanges())
        {
            if (!_dataTable.HasErrors && changes == null)
            {
                if (closeDB) { CloseDatabaseConnection(); }
                return;
            }

            //foreach (var kvp in dictEditField) { ApplyControlValue(kvp.Key, kvp.Value); }
            //BindingContext[_dataTable].EndCurrentEdit();
            addressDGV.EndEdit();
            if (addressDGV.DataSource != null && BindingContext != null)
            {
                BindingContext[addressDGV.DataSource].EndCurrentEdit();
            }
            if (changes != null)
            {
                if (tabControl.SelectedTab != addressTabPage) { tabControl.SelectTab(addressTabPage); }

                if (!askNever && sAskBeforeSaveSQL && !Utilities.YesNo_TaskDialog(Handle, appName, "Möchten Sie die Änderungen speichern?",
                    changes.Rows.Count == 1 ? "An einer Adresse wurden Änderungen vorgenommen." : $"Änderungen wurden an {changes.Rows.Count} Adressen vorgenommen.",
                    TaskDialogIcon.ShieldGrayBar, true, "&Speichern", "&Ignorieren"))
                {
                    _dataTable.RejectChanges();  // Benutzer hat "Ignorieren" gewählt, Änderungen verwerfen
                    saveTSButton.Enabled = false;
                    return;
                }
                try
                {
                    _adapter.Update(_dataTable); //_dataTable.AcceptChanges(); // nicht erforderlich, da Update dies bereits aufruft
                    saveTSButton.Enabled = false;
                    foreach (var entry in changedAddressData) { originalAddressData[entry.Key] = entry.Value; }
                    changedAddressData.Clear();
                    flexiTSStatusLabel.Text = $"Letztes Speichern: {DateTime.Now:HH:mm:ss}";
                }
                catch (DBConcurrencyException dbEx)
                {
                    Utilities.ErrorMsgTaskDlg(Handle, "Konflikt beim Speichern", $"Details: {dbEx.Message}\nIhre lokalen Änderungen werden verworfen.");
                    _dataTable?.RejectChanges();  // Änderungen im dataTable verwerfen, da sie nicht gespeichert werden konnten
                    saveTSButton.Enabled = false;
                }
                catch (Exception ex)
                {
                    Utilities.ErrorMsgTaskDlg(Handle, "Fehler beim Speichern der Datenbank", ex.Message);
                    _dataTable?.RejectChanges();
                    saveTSButton.Enabled = false;
                }
                if (sDailyBackup && File.Exists(Utilities.CorrectUNC(databaseFilePath)) && Directory.Exists(sBackupDirectory))
                {
                    if (isFormClosing) { Hide(); } // DailyBackup-Success-TaskDialog soll ohne Hauptfenster erscheinen 
                    Utilities.DailyBackup(Utilities.CorrectUNC(databaseFilePath), sBackupDirectory, sBackupSuccess, sSuccessDuration);
                }
            }
        }
        if (closeDB) { CloseDatabaseConnection(); }
    }

    //private void OnRowUpdating(object? sender, RowUpdatingEventArgs e)
    //{
    //    if (e.StatementType == StatementType.Update)
    //    {
    //        var row = e.Row;

    //        // Sicherstellen, dass idColumnInChanges nicht null ist
    //        var idColumnInChanges = row.Table.Columns["Id"];
    //        var isIdNullOrZero = (idColumnInChanges != null && (row.IsNull(idColumnInChanges, DataRowVersion.Original) ||
    //                                  (row[idColumnInChanges, DataRowVersion.Original] is long originalId && originalId == 0)));

    //        if (isIdNullOrZero)
    //        {
    //            e.Status = UpdateStatus.SkipCurrentRow;

    //            // Manuelle INSERT-Logik, die den Fehler "SetAdded" umgeht:
    //            var connection = e.Command?.Connection;
    //            var transaction = e.Command?.Transaction;

    //            if (connection == null || transaction == null) { e.Status = UpdateStatus.ErrorsOccurred; return; }

    //            // INSERT Command
    //            using var insertCmd = connection.CreateCommand();
    //            insertCmd.Transaction = transaction;

    //            var tableName = "Adressen";
    //            var dataFields = row.Table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).Where(f => f != "Id").ToArray();
    //            var insertFields = string.Join(", ", dataFields.Select(f => $"[{f}]"));
    //            var insertValues = string.Join(", ", dataFields.Select(f => $"@{f}"));
    //            insertCmd.CommandText = $"INSERT INTO {tableName} ({insertFields}) VALUES ({insertValues});";

    //            foreach (var field in dataFields)
    //            {
    //                var param = insertCmd.CreateParameter();
    //                param.ParameterName = $"@{field}";
    //                param.Value = row[field, DataRowVersion.Current];
    //                insertCmd.Parameters.Add(param);
    //            }

    //            var rowsAffected = insertCmd.ExecuteNonQuery();

    //            if (rowsAffected == 1)
    //            {
    //                // ID abrufen (fehlerfrei, da generische Schnittstellen verwendet werden)
    //                using var idCmd = connection.CreateCommand();
    //                idCmd.CommandText = "SELECT last_insert_rowid()";
    //                idCmd.Transaction = transaction;

    //                var newId = (long)(idCmd.ExecuteScalar() ?? 0L);

    //                // Wichtig: ID zuweisen und akzeptieren.
    //                row["Id"] = newId;
    //                row.AcceptChanges();
    //            }
    //            else
    //            {
    //                e.Status = UpdateStatus.ErrorsOccurred;
    //            }

    //            e.Status = UpdateStatus.Continue;
    //            return;
    //        }
    //    }
    //}

    //private void OnRowUpdating(object? sender, RowUpdatingEventArgs e)
    //{
    //    if (e.StatementType == StatementType.Update)
    //    {
    //        var row = e.Row;

    //        // Prüfen, ob der Original-ID-Wert 0 oder Null ist.
    //        var idColumnInChanges = row.Table.Columns["Id"];
    //        var isIdNullOrZero = row.IsNull(idColumnInChanges, DataRowVersion.Original) ||
    //                              (row[idColumnInChanges, DataRowVersion.Original] is long originalId && originalId == 0);

    //        if (isIdNullOrZero)
    //        {
    //            e.Status = UpdateStatus.SkipCurrentRow;

    //            // Manuelle INSERT-Logik, die den Fehler "SetAdded" umgeht:
    //            var connection = e.Command?.Connection;
    //            var transaction = e.Command?.Transaction;

    //            if (connection == null || transaction == null) { e.Status = UpdateStatus.ErrorsOccurred; return; }

    //            // INSERT Command
    //            using var insertCmd = connection.CreateCommand();
    //            insertCmd.Transaction = transaction;

    //            var tableName = "Adressen";
    //            var dataFields = row.Table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).Where(f => f != "Id").ToArray();
    //            var insertFields = string.Join(", ", dataFields.Select(f => $"[{f}]"));
    //            var insertValues = string.Join(", ", dataFields.Select(f => $"@{f}"));
    //            insertCmd.CommandText = $"INSERT INTO {tableName} ({insertFields}) VALUES ({insertValues});";

    //            foreach (var field in dataFields)
    //            {
    //                var param = insertCmd.CreateParameter();
    //                param.ParameterName = $"@{field}";
    //                param.Value = row[field, DataRowVersion.Current];
    //                insertCmd.Parameters.Add(param);
    //            }

    //            var rowsAffected = insertCmd.ExecuteNonQuery();

    //            if (rowsAffected == 1)
    //            {
    //                // ID abrufen (fehlerfrei, da generische Schnittstellen verwendet werden)
    //                using var idCmd = connection.CreateCommand();
    //                idCmd.CommandText = "SELECT last_insert_rowid()";
    //                idCmd.Transaction = transaction;

    //                var newId = (long)(idCmd.ExecuteScalar() ?? 0L);

    //                // Wichtig: ID zuweisen und akzeptieren.
    //                row["Id"] = newId;
    //                row.AcceptChanges();
    //            }
    //            else
    //            {
    //                e.Status = UpdateStatus.ErrorsOccurred;
    //            }

    //            e.Status = UpdateStatus.Continue;
    //            return;
    //        }
    //    }
    //}

    //private void OnRowUpdating(object? sender, RowUpdatingEventArgs e)
    //{
    //    // Wir prüfen, ob der DataAdapter versucht, ein UPDATE auszuführen
    //    if (e.StatementType == StatementType.Update)
    //    {
    //        var row = e.Row;

    //        // Annahme: ID ist Int64 (long) und der Original-Wert 0 bedeutet 'noch nicht in DB gespeichert'.
    //        if (row["Id", DataRowVersion.Original] is long originalId && originalId == 0)
    //        {
    //            // 1. Zwingen Sie den DataAdapter, den fehlerhaften UPDATE-Versuch zu überspringen.
    //            e.Status = UpdateStatus.SkipCurrentRow;

    //            // 2. Extrahieren Sie die Connection und die Transaktion.
    //            var connection = e.Command?.Connection;
    //            var transaction = e.Command?.Transaction;

    //            // Wenn keine Connection oder Transaktion verfügbar ist, brechen wir ab.
    //            if (connection == null || transaction == null)
    //            {
    //                e.Status = UpdateStatus.ErrorsOccurred;
    //                return;
    //            }

    //            // 3. Manuelles Erstellen des INSERT-Befehls.
    //            // Erstellen Sie einen Befehl über die generische IDbConnection, um Typ-Konflikte zu vermeiden.
    //            using var insertCmd = connection.CreateCommand();
    //            insertCmd.Transaction = transaction; // Die IDbTransaction zuweisen

    //            var tableName = "Adressen"; // Tabellename
    //                                        // Alle Spalten außer der ID für INSERT-Befehl verwenden
    //            var dataFields = row.Table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).Where(f => f != "Id").ToArray();

    //            var insertFields = string.Join(", ", dataFields.Select(f => $"[{f}]"));
    //            var insertValues = string.Join(", ", dataFields.Select(f => $"@{f}"));
    //            insertCmd.CommandText = $"INSERT INTO {tableName} ({insertFields}) VALUES ({insertValues});";

    //            // Parameter setzen: WICHTIG: Verwenden Sie hier den aktuellen Wert der Zeile (DataRowVersion.Current).
    //            foreach (var field in dataFields)
    //            {
    //                var param = insertCmd.CreateParameter();
    //                param.ParameterName = $"@{field}";
    //                param.Value = row[field, DataRowVersion.Current];
    //                insertCmd.Parameters.Add(param);
    //            }

    //            // 4. INSERT manuell ausführen und ID abrufen.
    //            var rowsAffected = insertCmd.ExecuteNonQuery();

    //            if (rowsAffected == 1)
    //            {
    //                // ID abrufen: Erneut einen generischen Command verwenden.
    //                using var idCmd = connection.CreateCommand();
    //                idCmd.CommandText = "SELECT last_insert_rowid()";
    //                idCmd.Transaction = transaction;

    //                var newId = (long)(idCmd.ExecuteScalar() ?? 0L);

    //                // 5. ID der DataRow zuweisen und Änderungen akzeptieren.
    //                row["Id"] = newId;
    //                row.AcceptChanges();
    //            }
    //            else
    //            {
    //                // Wenn INSERT fehlschlägt, Fehlerstatus setzen.
    //                e.Status = UpdateStatus.ErrorsOccurred;
    //            }

    //            // 6. Setzen Sie den Status auf Continue für die nächste Zeile.
    //            e.Status = UpdateStatus.Continue;
    //            return;
    //        }
    //    }
    //}

    //private void OnRowUpdated(object? sender, RowUpdatedEventArgs e)
    //{
    //    Console.Beep();
    //    if (e.StatementType == StatementType.Insert && e.Status == UpdateStatus.Continue && e.Command != null
    //        && e.Command.Connection is SQLiteConnection sqliteConn && e.Command.Transaction != null)
    //    {
    //        // 1. Transaktion des aktuellen Commands für den neuen Befehl verwenden
    //        var sqliteTransaction = (SQLiteTransaction)e.Command.Transaction;

    //        // 2. ID abrufen
    //        using var cmd = new SQLiteCommand("SELECT last_insert_rowid()", sqliteConn, sqliteTransaction);
    //        var newId = (long)(cmd.ExecuteScalar() ?? 0L);

    //        // 3. ID in die Zeile schreiben
    //        e.Row["Id"] = newId;

    //        // 4. Änderungen akzeptieren (setzt RowState von Added/Modified auf Unchanged)
    //        e.Row.AcceptChanges();
    //    }
    //}

    private void CloseDatabaseConnection()
    {
        if (_dataTable != null)
        {
            _dataTable.Dispose();
            _dataTable = null;
        }
        addressDGV.DataSource = null;
        addressDGV.Rows.Clear();
        AddressEditFields(-1);
        duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled
            = deleteTSButton.Enabled = newToolStripMenuItem.Enabled = newTSButton.Enabled
            = duplicateToolStripMenuItem.Enabled = copyTSButton.Enabled = wordTSButton.Enabled
            = envelopeTSButton.Enabled = false;
        copyToOtherDGVTSMenuItem.Enabled = false;
        flexiTSStatusLabel.Text = string.Empty;
        searchTSTextBox.TextBox.Clear();
        tsClearLabel.Visible = false;
    }

    private async void OpenToolStripMenuItem_Click(object? sender, EventArgs? e)
    { //openFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp()) { await CreateContactAsync(); }
            if (CheckContactDataChange()) { ShowMultiPageTaskDialog(); }
        }
        var fileName = Path.GetFileName(Utilities.CorrectUNC(databaseFilePath)) ?? string.Empty;
        var dirName = Path.GetDirectoryName(Utilities.CorrectUNC(databaseFilePath)) ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(fileName)) { openFileDialog.FileName = fileName; }
        else { openFileDialog.FileName = "Adressen.adb"; }
        openFileDialog.InitialDirectory = !string.IsNullOrEmpty(sDatabaseFolder) && Directory.Exists(sDatabaseFolder) ? sDatabaseFolder : !string.IsNullOrWhiteSpace(dirName) ? dirName : null;
        openFileDialog.Multiselect = false;

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            if (_dataTable != null) { SaveSQLDatabase(true); }
            ConnectSQLDatabase(openFileDialog.FileName);
            ignoreSearchChange = true;
            searchTSTextBox.TextBox.Clear();
            ignoreSearchChange = false;
            if (birthdayAddressShow) { BirthdayReminder(); }
        }
    }

    private void ExitToolStripMenuItem_Click(object? sender, EventArgs? e)
    {
        if (_dataTable != null) { SaveSQLDatabase(true); }
        Close();
    }

    private void AddressDGV_SelectionChanged(object sender, EventArgs e)
    {   // Leere (neue) Zeilen werden BEIM SPEICHERN gelöscht   
        if (isSelectionChanging) { return; }
        isSelectionChanging = true;
        try
        {
            if (addressDGV.SelectedRows.Count > 0)
            {
                prevSelectedAddressRowIndex = addressDGV.SelectedRows[0].Index;
                AddressEditFields(prevSelectedAddressRowIndex);
            }
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "AddressDGV_SelectionChanged: " + ex.GetType().ToString(), ex.Message); }
        finally { isSelectionChanging = false; }
    }

    private async void AddressDGV_CellClick(object sender, DataGridViewCellEventArgs e)
    {
        if ((NativeMethods.GetKeyState(NativeMethods.VK_CONTROL) & 0x8000) != 0 && e.ColumnIndex >= 0)
        {
            var colName = addressDGV.Columns[e.ColumnIndex].Name;
            if (!string.IsNullOrEmpty(colName))
            {
                using (var row = addressDGV.Rows[e.RowIndex])
                {
                    if (!row.Selected) { row.Selected = true; }
                }
                await Task.Delay(50);
                foreach (Control control in tableLayoutPanel.Controls)
                {
                    if (control.Name != null && control.Name.EndsWith(colName))
                    {
                        control.Focus();
                        break;
                    }
                }
            }
        }
    }

    private void AddressEditFields(int rowIndex) // rowIndex = -1 => ClearFields
    {
        tabulation.SelectedTab = tabPageDetail;
        try
        {
            ignoreTextChange = true; // verhindert, dass TextChanged
            foreach (var (ctrl, colText) in dictEditField) { ctrl.Text = rowIndex < 0 ? "" : addressDGV.Rows[rowIndex].Cells[colText]?.Value?.ToString() ?? ""; }

            if (rowIndex >= 0 && _dataTable != null && !(_dataTable.Rows[rowIndex].RowState == DataRowState.Added)) // NICHT !addressDGV.Rows[rowIndex].IsNewRow) 
            {
                var currentContactId = Convert.ToInt32(addressDGV.Rows[rowIndex].Cells["Id"].Value);
                var kontaktFoto = LadeFotoFuerAddress(currentContactId);
                if (kontaktFoto != null) // && !kontaktFoto.Size.IsEmpty)
                {
                    topAlignZoomPictureBox.Image = kontaktFoto;
                    delPictboxToolStripButton.Enabled = true;
                }
                else
                {
                    topAlignZoomPictureBox.Image = Resources.AddressBild100;
                    delPictboxToolStripButton.Enabled = false;
                }
            }
            cbGrußformel.Items.Clear();
            if (rowIndex >= 0)
            {
                ErzeugeGrußformeln();
                if (DateTime.TryParse(addressDGV.Rows[rowIndex].Cells["Geburtstag"]?.Value?.ToString(), out var date))
                {
                    maskedTextBox.Text = date.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("de-DE"));
                    AgeLabel_SetText(date);
                }
                else
                {
                    AgeLabel_DeleteText();
                    maskedTextBox.Text = string.Empty;
                }
                var membershipsJson = addressDGV.Rows[rowIndex].Cells["Gruppen"].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(membershipsJson))
                {
                    var deserialized = JsonSerializer.Deserialize<List<string>>(membershipsJson) ?? [];
                    curAddressMemberships = new SortedSet<string>(deserialized, StringComparer.OrdinalIgnoreCase);
                    allAddressMemberships.UnionWith(curAddressMemberships);
                    UpdateMembershipTags();
                }
                else
                {
                    curAddressMemberships.Clear();
                    flowLayoutPanel.Controls.Clear();
                    UpdatePlaceholderVis();
                }
                UpdateMembershipCBox();
            }
            else
            {
                curAddressMemberships.Clear();
                flowLayoutPanel.Controls.Clear();
                UpdatePlaceholderVis();
                UpdateMembershipCBox();
            }

            tbNotizen.Text = rowIndex < 0 ? "" : addressDGV.Rows[rowIndex].Cells["Notizen"]?.Value?.ToString();

            dokuListView.Items.Clear();
            if (rowIndex >= 0 && addressDGV.Rows[rowIndex].DataBoundItem is DataRowView rowView)
            {
                var json = rowView.Row["Dokumente"]?.ToString() ?? string.Empty;
                if (!string.IsNullOrEmpty(json))
                {
                    var dateipfade = JsonSerializer.Deserialize<List<string>>(json);
                    if (dateipfade != null)
                    {
                        foreach (var pfad in dateipfade)
                        {
                            var info = new FileInfo(pfad);
                            Add2dokuListView(new FileInfo(pfad), false);
                        }
                    }
                    dokuListView.ListViewItemSorter = new ListViewItemComparer();
                    dokuListView.Sort();
                }
            }

            originalAddressData.Clear();
            if (rowIndex >= 0)
            {
                foreach (DataGridViewCell cell in addressDGV.Rows[rowIndex].Cells)
                {
                    var columnName = cell.OwningColumn.Name; // Spaltenname als Schlüssel verwenden
                    if (!string.IsNullOrEmpty(columnName)) { originalAddressData[columnName] = cell.Value?.ToString() ?? string.Empty; }
                }
            }

            tabPageDoku.ImageIndex = dokuListView.Items.Count > 0 ? 4 : 3;
            LinkLabel_Enabled();
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "AddressEditFields: " + ex.GetType().ToString(), ex.Message); }
        finally { ignoreTextChange = false; } // TextChanged wieder aktivieren
    }

    private Image? LadeFotoFuerAddress(int kontaktId)
    {
        Image? foto = null;
        //using (var connection = new SQLiteConnection((string?)$"Data Source={databaseFilePath};FailIfMissing=False;"))
        //{
        _connection = new SQLiteConnection((string?)$"Data Source={databaseFilePath};FailIfMissing=False;");
        _connection.Open();
        var query = "SELECT Fotodaten FROM Fotos WHERE AdressId = @id";
        using var cmd = new SQLiteCommand(query, _connection);
        cmd.Parameters.AddWithValue("@id", kontaktId);
        var result = cmd.ExecuteScalar();
        if (result != null && result != DBNull.Value)
        {
            using var ms = new MemoryStream((byte[])result);
            foto = Image.FromStream(ms);
        }
        //}
        return foto; // Gibt das Bild oder null zurück
    }

    private void AgeLabel_SetText(DateTime date)
    {
        maskedTextBox.Text = date.ToString("dd.MM.yyyy");  //ToString("d", CultureInfo.CurrentCulture);
        var days = (DateTime.Today - date).Days;
        if (Math.Abs(days) <= 31) { ageLabel.Text = Math.Abs(days).Equals(1) ? days.ToString() + " Tag" : days.ToString() + " Tage"; }
        else
        {
            var ddf = Utilities.CalcDateDiff(DateTime.Today, date);
            ageLabel.Text = (!ddf.years.Equals(0) ? ddf.years.ToString() + (ddf.years.Equals(1) ? " Jahr" : " Jahre") +
                (ddf.months.Equals(0) && ddf.days.Equals(0) ? "" : ", ") : "") + (!ddf.months.Equals(0) ? ddf.months.ToString() +
                (ddf.months.Equals(1) ? " Monat" : " Monate") + (ddf.days.Equals(0) ? "" : ", ") : "") +
                (!ddf.days.Equals(0) ? ddf.days.ToString() + (ddf.days.Equals(1) ? " Tag" : " Tage") : "");
            toolTip.SetToolTip(ageLabel, $"{days} Tage");
        }
    }

    private void AgeLabel_DeleteText()
    {
        maskedTextBox.Text = string.Empty;
        ageLabel.Text = string.Empty;
        toolTip.SetToolTip(ageLabel, string.Empty);
    }

    private void AddressDGV_DataSourceChanged(object sender, EventArgs e)
    {
        if (addressDGV.DataSource != null && hideColumnArr.Length == addressDGV.Columns.Count)
        {
            for (var i = 0; i < addressDGV.Columns.Count; i++) { addressDGV.Columns[i].Visible = !hideColumnArr[i]; }
            Text = appName + " – " + (string?)(string.IsNullOrEmpty(databaseFilePath) ? "unbenannt" : Utilities.CorrectUNC(databaseFilePath));  // Workaround for UNC-Path
        }
        else { Text = appLong; }
    }

    private void OpenTSButton_Click(object sender, EventArgs e) => OpenToolStripMenuItem_Click(sender, e);

    private void FrmAdressen_Resize(object sender, EventArgs e)
    {
        flexiTSStatusLabel.Width = 244 + splitContainer.SplitterDistance - 536;
        searchTSTextBox.Width = 202 + splitContainer.SplitterDistance - 536 - (tsClearLabel.Visible ? tsClearLabel.Width : 0);
        //ResizeListViewColumns();
    }

    private async void SearchTSTextBox_TextChanged(object sender, EventArgs e)
    {
        if (ignoreSearchChange) { return; }
        tsClearLabel.Visible = searchTSTextBox.TextBox.Text.Length > 0;
        try
        {
            var normalizedSearchTerm = Utilities.NormalizeString(searchTSTextBox.TextBox.Text);
            if (tabControl.SelectedTab == addressTabPage)
            {
                if (string.IsNullOrWhiteSpace(normalizedSearchTerm))
                {
                    FilterAddressDGV(row => true);
                    filterRemoveToolStripMenuItem.Visible = false;
                    flexiTSStatusLabel.Text = "";
                }
                else
                {
                    FilterAddressDGV(row =>
                    {
                        var content = string.Join(" ", row.Cells.Cast<DataGridViewCell>()
                        .Take(row.Cells.Count - 1) // Nimmt alle Zellen außer der letzten
                        .Select(c => c.Value?.ToString() ?? ""));
                        return Utilities.NormalizeString(content).Contains(normalizedSearchTerm); // Die eigentliche Filterbedingung
                    });
                }
            }
            else if (tabControl.SelectedTab == contactTabPage)
            {
                var test = contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp();
                var dataChanged = CheckContactDataChange();
                if (test || dataChanged)
                {
                    ignoreSearchChange = true;
                    searchTSTextBox.Text = lastSearchText;
                    searchTSTextBox.SelectionStart = searchTSTextBox.Text.Length; // Cursor ans Ende setzen
                    ignoreSearchChange = false;
                }
                if (dataChanged)
                {
                    ShowMultiPageTaskDialog();
                    return;
                }
                else if (test)
                {
                    await CreateContactAsync();
                    return;
                }
                if (string.IsNullOrWhiteSpace(normalizedSearchTerm)) { FilterContactDGV(row => true); }
                else
                {
                    FilterContactDGV(row =>
                    {
                        var content = string.Join(" ", row.Cells.Cast<DataGridViewCell>()
                        .Take(row.Cells.Count - 1) // Nimmt alle Zellen außer der letzten
                        .Select(c => c.Value?.ToString() ?? "")); //.Select(t => t.Replace(Environment.NewLine, " ")));
                        return Utilities.NormalizeString(content).Contains(normalizedSearchTerm); // Die eigentliche Filterbedingung
                    });
                }
            }
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "SearchTSTextBox_TextChanged: " + ex.GetType().ToString(), ex.Message); }
        finally
        {
            lastSearchText = searchTSTextBox.TextBox.Text;
            flexiTSStatusLabel.Text = "";
        }
    }

    private void AddressDGV_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
    {
        try
        {
            if (addressDGV.Columns["Geburtstag"] != null && e.Value != null && e.ColumnIndex == addressDGV.Columns["Geburtstag"].Index && DateTime.TryParse(e.Value.ToString(), out var dt)) // Angenommen, die Datumsspalte ist die erste
            {
                e.Value = dt.ToString("d.M.yyyy"); // Ändere das Format nach Bedarf
                e.FormattingApplied = true;
            }
        }
        catch (Exception ex)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "AddressDGV_CellFormatting: " + ex.GetType().ToString(), ex.Message);
            Application.Exit();
        }
    }

    private async void SaveTSButton_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _dataTable != null)
        {
            if (_dataTable != null)
            {
                var rowIndex = addressDGV.SelectedRows.Count > 0 ? addressDGV.SelectedRows[0].Index : -1;
                SaveSQLDatabase(false, true);
                if (rowIndex >= 0 && addressDGV.Rows[rowIndex] != null)
                {
                    addressDGV.Rows[rowIndex].Selected = true;
                    addressDGV.FirstDisplayedScrollingRowIndex = rowIndex;
                }
            }
        }
        else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp()) { await CreateContactAsync(); }
            if (CheckContactDataChange()) { ShowMultiPageTaskDialog(); }
        }
        else { Console.Beep(); }
    }

    private bool CheckContactDataChange()  //originalContactData wird in ContactEditFields() gesetzt 
    {
        if (originalContactData == null || originalContactData.Count == 0 || contactDGV == null || prevSelectedContactRowIndex < 0) { return false; }
        changedContactData.Clear();
        foreach (var cell in contactDGV.Rows[prevSelectedContactRowIndex].Cells.Cast<DataGridViewCell>().SkipLast(2).Where(cell => !Equals(originalContactData[cell.OwningColumn.Name], cell.Value)))
        {
            changedContactData[cell.OwningColumn.Name] = cell.Value?.ToString() ?? string.Empty;
        }
        if (changedContactData.Count > 0) { return true; }
        return false;
    }

    private bool CheckNewContactTidyUp()
    {
        if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex)
        {
            if (contactDGV.Rows[contactNewRowIndex].Cells.Cast<DataGridViewCell>().Any(cell => cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))) { return true; }
            else
            {
                contactDGV.Rows.RemoveAt(contactNewRowIndex);
                contactNewRowIndex = -1;
                return false;
            }
        }
        return false;
    }

    private void ShowMultiPageTaskDialog()
    {
        var ressource = contactDGV.Rows[prevSelectedContactRowIndex].Cells["Ressource"]?.Value?.ToString() ?? string.Empty;
        var message = "";
        foreach (var pair in changedContactData) { message += pair.Key + ": " + pair.Value + "\n"; }
        var initialButtonYes = new TaskDialogButton("Hochladen");
        var initialButtonNo = TaskDialogButton.Cancel;
        using TaskDialogIcon questionDialogIcon = new(Resources.question32);
        initialButtonYes.AllowCloseDialog = false; // don't close the dialog when this button is clicked
        var initialPage = new TaskDialogPage()
        {
            Caption = "Google Kontakte",
            Heading = "Möchten Sie die Änderungen speichern?",
            Text = Regex.Replace(message, "[\"\\[\\]]", string.Empty),
            Icon = questionDialogIcon, // TaskDialogIcon.ShieldBlueBar,
            AllowCancel = true,
            SizeToContent = true,
            Buttons = { initialButtonNo, initialButtonYes },
        };

        var inProgressCloseButton = TaskDialogButton.Close;
        inProgressCloseButton.Enabled = false;
        var progressPage = new TaskDialogPage()
        {
            Caption = appCont,
            Heading = "Bitte warten…",
            Text = "Änderungen werden hochgeladen.",
            Icon = TaskDialogIcon.Information,
            ProgressBar = new TaskDialogProgressBar() { State = TaskDialogProgressBarState.Marquee },
            Buttons = { inProgressCloseButton }
        };
        initialButtonYes.Click += (sender, e) => { initialPage.Navigate(progressPage); }; // When the user clicks "Yes", navigate to the second page.
        initialButtonNo.Click += (sender, e) =>
        {
            foreach (var entry in changedContactData)
            {
                var columnName = entry.Key; // Original-Einträge wieder herstellen (Wert aus originalContactData)
                if (contactDGV.Columns.Contains(columnName)) { contactDGV.Rows[prevSelectedContactRowIndex].Cells[columnName].Value = originalContactData[entry.Key]; }
            }
            ContactEditFields(prevSelectedContactRowIndex); // Update the contact edit fields with the original data
            changedContactData.Clear(); // Leeren des Dictionaries nach der Aktualisierung
        };
        progressPage.Created += async (s, e) => { await UpdateContactAsync(ressource, () => progressPage.Buttons.First().PerformClick()); };
        TaskDialog.ShowDialog(Handle, initialPage); // Show the initial page of the TaskDialog
    }

    private void TbNotizen_TextChanged(object sender, EventArgs e)
    {
        NativeMethods.ShowScrollBar(tbNotizen.Handle, 1, TextRenderer.MeasureText(tbNotizen.Text, tbNotizen.Font, new Size(tbNotizen.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue),
        TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height > tbNotizen.Height);
        if (ignoreTextChange) { return; } // verhindert, dass TextChanged bei AddressEditFields aufgerufen wird  
        if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0 && addressDGV.SelectedRows[0].Cells["Notizen"].Value.ToString() != tbNotizen.Text.Trim())
        {
            //addressDGV.SelectedRows[0].Cells["Notizen"].Value = tbNotizen.Text;
            if (addressDGV.SelectedRows[0].DataBoundItem is DataRowView dataBoundItem) { dataBoundItem.Row["Notizen"] = tbNotizen.Text; }
        }
        else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0) { contactDGV.SelectedRows[0].Cells["Notizen"].Value = tbNotizen.Text; }
        CheckSaveButton();
    }

    private void TbNotizen_SizeChanged(object sender, EventArgs e) => NativeMethods.ShowScrollBar(tbNotizen.Handle, 1, TextRenderer.MeasureText(tbNotizen.Text, tbNotizen.Font,
        new Size(tbNotizen.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height > tbNotizen.Height);

    private void BtnResetDate_Click(object sender, EventArgs e)
    {
        maskedTextBox.Clear();
        maskedTextBox.Focus();
    }

    private void DictEditField_TextChanged(object sender, EventArgs e)
    {
        if (ignoreTextChange) { return; }
        if (sender is Control ctrl && dictEditField.TryGetValue(ctrl, out var colName))
        {
            if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0)
            {
                if (addressDGV.SelectedRows[0].DataBoundItem is DataRowView dataBoundItem) { dataBoundItem.Row[colName] = ctrl.Text.Trim(); }
            }
            else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0) { contactDGV.SelectedRows[0].Cells[colName].Value = ctrl.Text.Trim(); }
            CheckSaveButton();
        }
    }

    private async void NewTSButton_Click(object sender, EventArgs e)
    {
        searchTSTextBox.TextBox.Clear();
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            if (contactDGV.SelectedRows[0] != null)
            {
                if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp()) { await CreateContactAsync(); }
                if (CheckContactDataChange()) { ShowMultiPageTaskDialog(); }
            }
            contactNewRowIndex = contactDGV.Rows.Add();
            contactDGV.Rows[contactNewRowIndex].Selected = true;
            contactDGV.FirstDisplayedScrollingRowIndex = contactDGV.Rows[contactNewRowIndex].Index;
            ContactEditFields(contactNewRowIndex);
            //saveTSButton.Enabled = true;
            //Utilities.ErrorMsgTaskDlg(Handle, "Neuer Kontakt", "Bitte füllen Sie die Felder aus. Die Änderungen werden erst übernommen, wenn Sie 'speichern' wählen.", TaskDialogIcon.ShieldBlueBar);
            cbAnrede.Focus();
        }
        else if (tabControl.SelectedTab == addressTabPage && _dataTable != null) // && _bindingSource != null)
        {
            SaveSQLDatabase(false, false); // Woraround für Problem mit den ComboBoxen; sobald sie den Fcous erhalten ohne dass ein Text gewählt wird entsteht der Fehler   
            var row = _dataTable.NewRow(); // Concurrency violation: the UpdateCommand affected 0 of the expected 1 records. 
            foreach (var field in dataFields) { row[field] = string.Empty; } // Id bleibt DBNull => Auto-Inkrement
            _dataTable.Rows.Add(row);
            SaveSQLDatabase(false, true);
            addressDGV.Rows[^1].Selected = true; //addressNewRowIndex = addressDGV.Rows[^1].Index;
            addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.SelectedRows[0].Index;
            cbAnrede.Focus();  // .NET9: InvokeAsync(() => { cbAnrede.Focus(); });

            //_bindingSource.AddNew();
            //var newRowIndex = _bindingSource.Position;
            //addressDGV.Rows[newRowIndex].Selected = true; //addressNewRowIndex = addressDGV.Rows[^1].Index;
            //addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.SelectedRows[newRowIndex].Index;
            //BeginInvoke(() => { cbAnrede.Focus(); }); // .NET9: InvokeAsync(() => { cbAnrede.Focus(); });
        }
    }

    private async void CopyTSButton_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            if (contactDGV.SelectedRows[0] != null)
            {
                if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp())
                {
                    await CreateContactAsync();
                    return;
                }
                if (CheckContactDataChange())
                {
                    ShowMultiPageTaskDialog();
                    return;
                }

            }
            isSelectionChanging = true; // verhindert, dass ContactEditFields() aufgerufen wird 
            contactNewRowIndex = contactDGV.Rows.Add();
            contactDGV.Rows[contactNewRowIndex].Selected = true;
            contactDGV.FirstDisplayedScrollingRowIndex = contactNewRowIndex;
            foreach (var entry in originalContactData)
            {
                var columnName = entry.Key;
                if (contactDGV.Columns.Contains(columnName)) { contactDGV.Rows[contactNewRowIndex].Cells[columnName].Value = entry.Value; }
            }

            ContactEditFields(contactNewRowIndex);
            saveTSButton.Enabled = true;
            //Utilities.ErrorMsgTaskDlg(Handle, "Neuer Kontakt", "Die Änderungen werden erst übernommen, wenn Sie 'speichern' wählen.", TaskDialogIcon.ShieldBlueBar);
            cbAnrede.Focus();
            isSelectionChanging = false; // setzt isSelectionChanging zurück, damit ContactEditFields() wieder aufgerufen wird
        }
        else if (tabControl.SelectedTab == addressTabPage && _dataTable != null && addressDGV.SelectedRows[0] != null)
        {
            var newRow = _dataTable.NewRow();
            //newRow["Id"] = dataTable.Rows.Count > 0 ? dataTable.AsEnumerable().Where(r => r.RowState != DataRowState.Deleted).Max(r => r.Field<long>("Id")) + 1 : 1;
            if (addressDGV.SelectedRows[0].DataBoundItem is DataRowView dataBoundItem) { newRow.ItemArray = dataBoundItem.Row.ItemArray; }
            else { return; }
            _dataTable.Rows.Add(newRow);
            addressDGV.Rows[^1].Selected = true;
            addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.Rows[^1].Index;
            BeginInvoke(() => { cbAnrede.Focus(); }); // .NET9: InvokeAsync(() => { cbAnrede.Focus(); });
        }
        else { Console.Beep(); }
    }

    private async void CopyToOtherDGVMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0 && _dataTable != null)
        {
            using var selectedRow = contactDGV.SelectedRows[0];
            var newRow = _dataTable.NewRow();
            selectedRow.Cells.Cast<DataGridViewCell>().SkipLast(2).ToList().ForEach(cell => newRow[cell.ColumnIndex] = cell.Value);
            _dataTable.Rows.Add(newRow); // <--- Nur einmal hinzufügen!
            //using var liteConnection = new SQLiteConnection($"Data Source={databaseFilePath};FailIfMissing=True");
            _connection = new SQLiteConnection($"Data Source={databaseFilePath};FailIfMissing=True");
            _connection.Open();

            var selectQuery = "SELECT * FROM Adressen"; // oder dein Tabellenname
            //using var dataAdapter = new SQLiteDataAdapter(selectQuery, liteConnection);
            _adapter = new SQLiteDataAdapter(selectQuery, _connection);
            using var builder = new SQLiteCommandBuilder(_adapter);
            _adapter.InsertCommand = builder.GetInsertCommand();
            _adapter.Update(_dataTable);
            _dataTable.AcceptChanges();
            long echteId;
            using (var cmd = new SQLiteCommand("SELECT last_insert_rowid()", _connection))
            {
                echteId = (long)(cmd.ExecuteScalar() ?? 0);
                _dataTable.Rows[^1]["Id"] = (int)echteId; // Setze die echte Id in der DataTable  
                _dataTable.AcceptChanges();
            }
            var photoUrl = selectedRow.Cells["PhotoURL"].Value.ToString();
            if (!string.IsNullOrEmpty(photoUrl))
            {
                if (Uri.IsWellFormedUriString(photoUrl, UriKind.Absolute))
                {
                    try
                    {
                        var imageData = await HttpService.Client.GetByteArrayAsync(photoUrl); // byte[]
                        if (imageData != null && imageData.Length > 0)
                        {
                            using var ms = new MemoryStream(imageData);
                            topAlignZoomPictureBox.Image = Image.FromStream(ms);
                            SpeichereFotoFuerKontakt((int)echteId, ms.ToArray(), databaseFilePath);
                            delPictboxToolStripButton.Enabled = true;
                        }
                    }
                    catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "CopyToOtherDGV: " + ex.GetType().ToString(), ex.Message); }
                }
            }

            tabControl.SelectedTab = addressTabPage;
            searchTSTextBox.TextBox.Clear();
            //dataTable.Rows.Add(newRow);
            if (addressDGV.RowCount > 0)
            {
                addressDGV.Rows[^1].Selected = true;
                addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.Rows[^1].Index;
                //AddressEditFields(addressDGV.Rows[^1].Index); // wird durch AddressDGV_SelectionChanged aufgerufen  
                cbAnrede.Focus();
                saveTSButton.Enabled = true;
            }
        }
        else if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0 && contactDGV != null)
        {
            using var selectedRow = addressDGV.SelectedRows[0];
            contactNewRowIndex = contactDGV.Rows.Add();
            selectedRow.Cells.Cast<DataGridViewCell>().SkipLast(2).ToList().ForEach(cell => contactDGV.Rows[contactNewRowIndex].Cells[cell.ColumnIndex].Value = cell.Value);
            var image = LadeFotoFuerAddress(selectedRow.Index);
            tabControl.SelectedTab = contactTabPage;
            searchTSTextBox.TextBox.Clear();
            contactDGV.Rows[contactNewRowIndex].Selected = true;
            contactDGV.FirstDisplayedScrollingRowIndex = contactNewRowIndex;
            ContactEditFields(contactNewRowIndex, image);
            cbAnrede.Focus();
            saveTSButton.Enabled = true;
        }
        else { Console.Beep(); } // für Tastenkombination Strg+K
    }


    private async void DeleteTSButton_Click(object sender, EventArgs e)
    {
        var delete = false;

        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0 && contactDGV.SelectedRows[0] != null)
        {
            var row = contactDGV.SelectedRows[0];
            (sAskBeforeDelete, delete) = Utilities.AskBeforeDeleteTaskDlg(Handle, row, sAskBeforeDelete, false); // false = keine Verification (ask before delete)  
            if (delete && row != null) { await DeleteGoogleContact(row.Index); }
        }

        else if (addressDGV.SelectedRows.Count > 0 && !addressDGV.SelectedRows[0].IsNewRow && _dataTable != null) // && !addressDGV.SelectedRows[0].IsNewRow
        {
            var row = addressDGV.SelectedRows[0];
            if (sAskBeforeDelete && !row.Cells.Cast<DataGridViewCell>().SkipLast(1).All(c => string.IsNullOrEmpty(c.Value?.ToString()?.Trim())))
            {
                (sAskBeforeDelete, delete) = Utilities.AskBeforeDeleteTaskDlg(Handle, row, sAskBeforeDelete);
            }
            else { delete = true; }
            if (delete)
            {
                var indexToDelete = row.Index; // Schritt 1: Den Index der aktuell ausgewählten Zeile merken, BEVOR sie gelöscht wird.

                if (row.DataBoundItem is DataRowView dataRowView) { dataRowView.Row.Delete(); } // Schritt 2: Die Zeile aus der Datenquelle löschen.
                saveTSButton.Enabled = true;
                if (addressDGV.Rows.Count == 0) // Wenn keine Zeilen mehr vorhanden sind, Auswahl löschen und Felder leeren.
                {
                    ignoreSearchChange = true;
                    searchTextBox.Clear();
                    ignoreSearchChange = false;
                    AddressEditFields(-1);
                    return;
                }
                if (searchTSTextBox.TextLength > 0) { SearchTSTextBox_TextChanged(null!, null!); } // Schritt 3: Den Filter neu anwenden, wenn ein Suchtext vorhanden ist.
                var nextSelectedIndex = -1; // Schritt 4: Die vorherige sichtbare Zeile finden und auswählen.
                for (var i = indexToDelete; i >= 0; i--)  // Wir starten die Suche beim Index der gelöschten Zeile und gehen rückwärts.
                {
                    if (i < addressDGV.Rows.Count && addressDGV.Rows[i].Visible)
                    {
                        nextSelectedIndex = i;
                        break; // Die erste gefundene sichtbare Zeile ist die richtige.
                    }
                }
                addressDGV.ClearSelection(); // Zuerst die alte Auswahl löschen.
                if (nextSelectedIndex != -1) // Schritt 5: Die gefundene Zeile auswählen.
                {
                    addressDGV.Rows[nextSelectedIndex].Selected = true;
                    //addressDGV.FirstDisplayedScrollingRowIndex = nextSelectedIndex; // Optional: Sicherstellen, dass die Zeile auch sichtbar gescrollt wird.
                }
                else if (addressDGV.Rows.GetRowCount(DataGridViewElementStates.Visible) > 0) // Falls keine vorherige Zeile gefunden wurde (z.B. weil die erste Zeile gelöscht wurde),
                {
                    for (var i = 0; i < addressDGV.Rows.Count; i++) // Finde die erste sichtbare Zeile von oben
                    {
                        if (addressDGV.Rows[i].Visible)
                        {
                            addressDGV.Rows[i].Selected = true;
                            addressDGV.FirstDisplayedScrollingRowIndex = i;
                            break;
                        }
                    }
                }
            }
        }
        else { Console.Beep(); }
    }

    private async void FrmAdressen_FormClosing(object sender, FormClosingEventArgs e)
    {
        SaveConfiguration(); // GetColumeWidths funktioniert nur wenn TabPages noch vorhanden sind

        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp()) { await CreateContactAsync(); }
            if (CheckContactDataChange()) { ShowMultiPageTaskDialog(); }
        }
        if (_dataTable != null) { SaveSQLDatabase(true, false, true); } // macht Hide für Datensicherung
        if (wordDoc != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
            wordDoc = null;
        }
        if (wordApp != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            wordApp = null;
        }
    }

    private void AboutToolStripMenuItem_Click(object sender, EventArgs e) => Utilities.HelpMsgTaskDlg(Handle, appLong, Icon);

    private void AddressDGV_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e) => toolStripStatusLabel.Text = addressDGV.RowCount.ToString() + " Adressen";

    private void AddressDGV_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e) => toolStripStatusLabel.Text = addressDGV.RowCount.ToString() + " Adressen";


    private void ErzeugeGrußformeln()
    {
        cbGrußformel.Items.Clear();
        var pt = new List<(string Key, string Value)> { ("#vorname", tbVorname.Text), ("#nickname", tbNickname.Text), ("#nachname", tbNachname.Text), ("#titel", cbPräfix.Text) };
        var gender = GetGender(tbVorname.Text);
        cbGrußformel.Items.AddRange([.. (gender == false ? männlichGrusse : gender == true ? weiblichGrusse : weiblichGrusse.Concat(männlichGrusse))
            .Select(s => { var result = s; foreach (var (key, value) in pt.Where(p => !string.IsNullOrWhiteSpace(p.Value))) { result = result.Replace(key, value); } return result; })
            .Where(text => !text.Contains('#')).Distinct()]);
    }


    private void ImportToolStripMenuItem_Click(object sender, EventArgs e)
    {
        var btnCreateCSV = new TaskDialogButton("Beispiel-CSV-Datei erstellen");
        var firstPage = new TaskDialogPage()
        {
            Caption = Application.ProductName,
            Heading = "Folgende Spaltennamen werden in der ersten Zeile der CSV-Datei erwartet:",
            Text = string.Join(", ", dataFields.SkipLast(1)) + "\n\nNicht alle Spaltennamen sind erforderlich.\nDie Spaltenreihenfolge ist beliebig.", // Dokumente ausgenommen  
            Icon = TaskDialogIcon.Information,
            AllowCancel = true,
            Buttons = { btnCreateCSV, TaskDialogButton.Continue }
        };
        var result = TaskDialog.ShowDialog(this, firstPage);
        if (result == btnCreateCSV)
        {
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var filePath = Path.Combine(desktopPath, "adress.csv");
            try
            {
                using var writer = new StreamWriter(filePath, false, Encoding.UTF8); // false = überschreiben   
                writer.WriteLine(string.Join(";", dataFields.SkipLast(1))); // Dokumente ausgenommen  
                writer.WriteLine(string.Join(";", dataFields.SkipLast(1).Select(s => s.ToLower()))); // Beispielinhalt
                var secondPage = new TaskDialogPage()
                {
                    Caption = Application.ProductName,
                    Heading = "Beispiel-CSV-Datei erstellt",
                    Text = "Die Datei 'adress.csv' wurde auf Ihrem Desktop gespeichert.\nDie Trennung der Spalten erfolgt durch ein Semikolon (;).",
                    Icon = TaskDialogIcon.Information,
                    AllowCancel = true,
                    SizeToContent = true,
                    DefaultButton = TaskDialogButton.Continue,
                    Buttons = { TaskDialogButton.Continue, TaskDialogButton.Cancel }
                };
                if (TaskDialog.ShowDialog(this, secondPage) != TaskDialogButton.Continue) { return; }
            }
            catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "Fehler beim Erstellen der CSV-Datei!", ex.Message); }
        }
        else if (result != TaskDialogButton.Continue) { return; }

        openFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*";
        openFileDialog.FileName = "adress.csv";
        openFileDialog.Title = "CSV-Datei zum Importieren auswählen";
        openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        openFileDialog.Multiselect = false;
        if (openFileDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(openFileDialog.FileName))
        {
            if (_dataTable == null)
            {
                try
                {   //If the database file doesn't exist, the default behaviour is to create a new database file. Use 'FailIfMissing=True' to raise an error instead.
                    databaseFilePath = Path.ChangeExtension(openFileDialog.FileName, ".adb");
                    CreateNewDatabase(databaseFilePath, false);
                    ConnectSQLDatabase(databaseFilePath);
                }
                catch (Exception ex)
                {
                    Utilities.ErrorMsgTaskDlg(Handle, "ImportToolStripMenuItem_Click: " + ex.GetType().ToString(), ex.Message);
                    _dataTable = null;
                    return;
                }
            }

            var allowedColumns = new HashSet<string>(dataFields.SkipLast(1)); // HashSet, Dokumente ausgenommen
            if (allowedColumns == null) { return; }
            var reader = Utilities.ReadAsLines(openFileDialog.FileName);
            if (!reader.Any()) { return; }  // Die Datei ist leer, nichts zu tun.
            var headers = reader.First().Split(';'); // Annahme: Die erste Zeile enthält die Spaltennamen

            var unknownColumns = headers.Where(h => !string.IsNullOrEmpty(h) && !allowedColumns.Contains(h)).ToList();
            if (unknownColumns.Count != 0)
            {
                var unknownColumnsStr = string.Join(", ", unknownColumns);
                Utilities.ErrorMsgTaskDlg(Handle, "Unbekannte Spaltennamen!", $"Die folgenden Spalten in der CSV-Datei sind ungültig: {unknownColumnsStr}\n\nDer Importvorgang wird abgebrochen.");
                return;
            }
            if (addressDGV.Rows.Count > 0 && !Utilities.YesNo_TaskDialog(Handle, appName, // dataTable.Rows.Count geht hier nicht, weil z.B. Löschen nicht gespeichert wurde
                heading: "Daten importieren", text: $"'{databaseFilePath}' enthält bereits Daten.\nMöchten Sie die neuen Daten aus der CSV-Datei hinzufügen?",
                TaskDialogIcon.ShieldWarningYellowBar)) { return; }

            var firstNewRowIndex = addressDGV.Rows.Count; // dataTable.Rows.Count geht hier nicht, weil z.B. Löschen nicht gespeichert wurde    

            var columnIndexMap = new Dictionary<int, string>();  // Mapping vom Spaltenindex in der CSV zum Spaltennamen (Wert)
            for (var i = 0; i < headers.Length; i++)
            {
                if (!string.IsNullOrEmpty(headers[i])) { columnIndexMap.Add(i, headers[i]); }
            }
            var records = reader.Skip(1);
            foreach (var record in records)
            {
                var splitArray = record.Split(';');
                if (splitArray.Length != headers.Length)
                {
                    Utilities.ErrorMsgTaskDlg(Handle, "Inkonsistente Daten!", $"Eine Zeile hat {splitArray.Length} Felder, aber die Kopfzeile hat {headers.Length}.\nDie Zeile wird übersprungen.");
                    continue; // Überspringt diese fehlerhafte Zeile und fährt mit der nächsten fort.
                }
                if (_dataTable == null) { continue; }
                var row = _dataTable.NewRow();
                if (row == null) { continue; }
                foreach (var mapping in columnIndexMap)
                {
                    var csvIndex = mapping.Key;
                    var columnName = mapping.Value;
                    var value = splitArray[csvIndex]; // Der Index wird verwendet, um den Wert aus dem splitArray zu lesen
                    row[columnName] = string.IsNullOrEmpty(value) ? DBNull.Value : value; // Der Name wird verwendet, um die richtige Spalte in der DataRow zu finden
                }
                _dataTable.Rows.Add(row);
                saveTSButton.Enabled = true;

                if (addressDGV.Rows.Count > firstNewRowIndex)
                {
                    addressDGV.Rows[firstNewRowIndex].Selected = true;
                    addressDGV.FirstDisplayedScrollingRowIndex = firstNewRowIndex;
                }
            }
        }
    }

    private void SearchTSTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            if (tabControl.SelectedTab == addressTabPage && addressDGV.Rows.GetRowCount(DataGridViewElementStates.Visible) > 0) { addressDGV.Focus(); }
            else if (tabControl.SelectedTab == contactTabPage && contactDGV.Rows.GetRowCount(DataGridViewElementStates.Visible) > 0)
            {
                var row = contactDGV.Rows.Cast<DataGridViewRow>().Where(row => row.Visible).FirstOrDefault();
                if (row != null)
                {
                    contactDGV.Focus();
                    row.Selected = true;
                    addressDGV.FirstDisplayedScrollingRowIndex = row.Index;
                    ContactEditFields(row.Index);
                }
            }
            e.Handled = e.SuppressKeyPress = true;
        }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        switch (keyData)
        {
            case Keys.Escape:
                {
                    if (addressDGV.CurrentCell != null && addressDGV.IsCurrentCellInEditMode)
                    {
                        addressDGV.EndEdit();
                        addressDGV.CurrentCell.Selected = true;
                    }
                    else if (ActiveControl == searchTSTextBox.Control && searchTSTextBox.TextLength > 0) { Clear_SearchTextBox(); }
                    else { searchTSTextBox.Focus(); }
                    return true;
                }
            case Keys.F11:
                foreach (var key in (string[])[.. addBookDict.Keys]) { addBookDict[key] = string.Empty; }
                Utilities.WordInfoTaskDlg(Handle, [.. addBookDict.Keys], new(Resources.word32));
                return true; // You return true to indicate that you handled the keystroke and don't want it to be passed on to other controls.
            case Keys.F5:
                tabControl.SelectedIndex = 0;
                return true;
            case Keys.F6:
                tabControl.SelectedIndex = 1;
                return true;
            case Keys.F7:
                tabulation.SelectedIndex = 0;
                return true;
            case Keys.F8:
                tabulation.SelectedIndex = 1;
                return true;
            case Keys.F1:
                Utilities.StartFile(Handle, @"AdressenKontakte.pdf");
                return true;
            case Keys.I | Keys.Control:
                Utilities.HelpMsgTaskDlg(Handle, appLong, Icon);
                return true;
            case Keys.F9:
                if (filterRemoveToolStripMenuItem.Visible)
                {
                    FilterRemoveToolStripMenuItem_Click(null!, null!);
                    return true;
                }
                else if ((tabControl.SelectedTab == addressTabPage && addressDGV.Rows.Count > 0) || (tabControl.SelectedTab == contactTabPage && contactDGV.Rows.Count > 0))
                {
                    GroupFilterToolStripMenuItem_Click(null!, null!);
                    return true;
                }
                else { return false; }
            case Keys.F9 | Keys.Control:
                ManageGroupsToolStripMenuItem_Click(null!, EventArgs.Empty);
                return true;
            case Keys.Enter | Keys.Control:
            case Keys.Tab | Keys.Shift: // Keys.Control funktioniert nicht
                tabControl.SelectedIndex = tabControl.SelectedIndex == 1 ? 0 : 1;
                return true;
            case Keys.F | Keys.Control:
                if (dokuListView.Focused)
                {
                    searchTextBox.Focus();
                    searchTextBox.SelectAll();
                }
                else
                {
                    searchTSTextBox.TextBox.Focus();
                    searchTSTextBox.TextBox.SelectAll();
                }
                return true;
            case Keys.N | Keys.Control:
                NewTSButton_Click(null!, null!);
                return true;
            case Keys.D | Keys.Control:
                CopyTSButton_Click(null!, null!);
                return true;
            case Keys.O | Keys.Control:
                OpenTSButton_Click(null!, null!);
                return true;
            case Keys.B | Keys.Control:
                BirthdayReminder();
                return true;
            case Keys.G | Keys.Control:
                GoogleTSButton_Click(null!, null!);
                return true;
            case Keys.E | Keys.Control:
                OptionsToolStripMenuItem_Click(null!, null!);
                return true;
            case Keys.K | Keys.Control:
                CopyToOtherDGVMenuItem_Click(null!, null!);
                return true;
            case Keys.F12:
                foreach (var file in recentFiles)
                {
                    if (file == databaseFilePath) { continue; }
                    if (File.Exists(file))
                    {
                        if (_dataTable != null) { SaveSQLDatabase(true); }
                        ConnectSQLDatabase(file);
                        ignoreSearchChange = true;
                        searchTSTextBox.TextBox.Clear();
                        ignoreSearchChange = false;
                    }
                    break;
                }
                return true;
            case Keys.S | Keys.Control:
                SaveTSButton_Click(null!, null!);
                return true;
            case Keys.T | Keys.Control:
                WordTSButton_Click(wordTSButton!, EventArgs.Empty!);
                return true;
            case Keys.U | Keys.Control:
                EnvelopeTSButton_Click(null!, null!);
                return true;
            case Keys.Z | Keys.Control:
                if (tabControl.SelectedTab == addressTabPage && _dataTable?.GetChanges() != null) { RejectChangesToolStripMenuItem_Click(null!, null!); }
                else { Console.Beep(); }
                return true;
            //case Keys.Delete:
            //    var enabledButton = flowLayoutPanel.Controls.OfType<Panel>().SelectMany(tagPanel => tagPanel.Controls.OfType<Button>()).FirstOrDefault(button => button.Enabled);
            //    if (enabledButton != null && flowLayoutPanel.ContainsFocus)
            //    {
            //        var membershipToRemove = enabledButton?.Tag as string;
            //        if (!string.IsNullOrEmpty(membershipToRemove))
            //        {
            //            if (tabControl.SelectedTab == contactTabPage) { curContactMemberships.Remove(membershipToRemove); } // ToDo: gegebenfalls auch aus allContactMemberships entfernen
            //            else { curAddressMemberships.Remove(membershipToRemove); } // ToDo: gegebenfalls auch aus allAddressMemberships entfernen - wenn andere Adressen sie nicht nutzen
            //            UpdateMembershipTags();  // true = curContactMemberships, false = curAddressMemberships
            //            UpdateMembershipJson(); // muss vor PopulateMemberships aufgerufen werden
            //            if (tabControl.SelectedTab == addressTabPage) { PopulateMemberships(); } // muss vor UpdateMembershipCBox aufgerufen werden
            //            UpdateMembershipCBox();
            //            UpdatePlaceholderVis();
            //        }
            //        return true;
            //    }
            //    else { return false; }
            case Keys.Delete | Keys.Control:
                if (tabControl.SelectedTab == addressTabPage)
                {
                    DeleteTSButton_Click(null!, null!);
                    return true;
                }
                else { return false; }
            case Keys.Enter | Keys.Alt:
                if (contactDGV.Focused)
                {
                    BtnEditContact_Click(null!, null!);
                    return true;
                }
                else { return false; }
            case Keys.F1 | Keys.Control | Keys.Shift:
                {
                    Utilities.StartDir(Handle, Path.GetDirectoryName(_settingsPath) ?? string.Empty);
                    return true;
                }
            case Keys.F2 | Keys.Control | Keys.Shift:
                {
                    Utilities.StartFile(Handle, _settingsPath);
                    return true;
                }
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void Edit_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            e.SuppressKeyPress = true;
            SelectNextControl((Control)sender, true, true, true, true);
        }
    }

    private void MaskedTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            e.SuppressKeyPress = true;
            SelectNextControl((Control)sender, true, true, true, true);
        }
        else if (e.KeyCode == Keys.Space)
        {
            e.SuppressKeyPress = true;
            BtnCalendar_Click(null!, null!);
        }
    }

    private void MaskedTextBox_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
    {
        if (e.KeyCode == Keys.OemPeriod)
        {
            var day = string.Empty;
            var month = string.Empty;
            var year = string.Empty;
            var dateComponents = maskedTextBox.Text.Split('.');
            if (dateComponents.Length > 0) { day = dateComponents[0].Trim(); }
            if (dateComponents.Length > 1) { month = dateComponents[1].Trim(); }
            if (dateComponents.Length > 2) { year = dateComponents[2].Trim(); }
            if (day.Length == 1) { day = "0" + day; }
            if (month.Length == 1) { month = "0" + month; }
            if (year.Length == 2) { year = "20" + year; }
            maskedTextBox.Text = day + "." + month + "." + year;
        }
    }

    private void TbInternet_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            e.SuppressKeyPress = true;
            tagComboBox.Focus(); // SelectNextControl((Control)sender, true, true, true, true);
        }
    }

    private void TbNotizen_Enter(object sender, EventArgs e)
    {
        tbNotizen.Select(tbNotizen.Text.Length, 0);
        tbNotizen.BackColor = Color.LightYellow;
    }

    private void InternetLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo(tbInternet.Text) { UseShellExecute = true }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
    }

    private void Mail1LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "mailto:" + tbMail1.Text }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
    }

    private void Mail2LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "mailto:" + tbMail2.Text }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
    }

    private void Tel1LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "tel:" + Regex.Replace(tbTelefon1.Text, cleanRegex, "") }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
    }

    private void Tel2LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "tel:" + Regex.Replace(tbTelefon2.Text, cleanRegex, "") }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
    }

    private void MobilLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "tel:" + Regex.Replace(tbMobil.Text, cleanRegex, "") }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
    }

    private void WordTSButton_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count == 0)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "Es ist keine Adresse gewählt!", "Es gibt keine Daten zu übertragen.");
            return;
        }
        else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count == 0)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "Es ist kein Kontakt gewählt!", "Es gibt keine Daten zu übertragen.");
            return;
        }

        var isWordInstalled = !(Type.GetTypeFromProgID("Word.Application") == null);
        var isLibreOfficeInstalled = !(Type.GetTypeFromProgID("com.sun.star.ServiceManager") == null); // Utilities.IsLibreOfficeInstalled();  
        //MessageBox.Show("isWordInstalled: " + isWordInstalled.ToString() + Environment.NewLine + "isLibreOfficeInstalled: " + isLibreOfficeInstalled.ToString(), "Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        if (sWordProcProg == true && !isWordInstalled)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "Word wurde nicht gefunden", "Installieren Sie Microsoft Word.");
            return;
        }
        else if (sWordProcProg == false && !isLibreOfficeInstalled)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "LibreOffice wurde nicht gefunden", "Installieren Sie LibreOffice Writer.");
            return;
        }
        else if (sWordProcProg == true) { WordProcess(); }
        else if (sWordProcProg == false) { LibreProcess(); }
        else if (sWordProcProg == null)
        {
            var result = Utilities.AskWordProcessingProgram(Handle);
            if (result == true) { WordProcess(); }
            else if (result == false) { LibreProcess(); }
        }
    }

    private void LibreProcess()
    {
        FillDictionary();
        var helperPath = Path.Combine(Path.GetDirectoryName(appPath) ?? string.Empty, "LibreHelper", "LibreOffice.exe");
        var lastWriterNoDoc = NativeMethods.GetLastVisibleHandleByTitleEnd("LibreOffice"); // Process.GetProcessesByName("soffice.bin") findet immer nur einen Prozess!!
        if (!File.Exists(helperPath)) { Utilities.ErrorMsgTaskDlg(Handle, @"LibreHelper\LibreOffice.exe nicht gefunden", helperPath, TaskDialogIcon.ShieldErrorRedBar); }
        else if (NativeMethods.GetLastVisibleHandleByTitleEnd("– LibreOffice Writer") != IntPtr.Zero) // geöffnentes Writer-Dokument
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = helperPath,
                Arguments = "\"" + JsonSerializer.Serialize(addBookDict).Replace("\"", "\\\"") + "\"",
                UseShellExecute = false,
                CreateNoWindow = true
            });
        }
        else if (lastWriterNoDoc != IntPtr.Zero) { NativeMethods.SetForegroundWindow(lastWriterNoDoc); }
        else // LibreOffice (Writer) ist nicht gestartet 
        {
            try
            {
                var libreOfficeDir = string.Empty;
                using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath");
                libreOfficeDir = key?.GetValue(null) as string;
                if (!string.IsNullOrEmpty(libreOfficeDir))
                {
                    var exePath = Path.Combine(libreOfficeDir, "soffice.exe");
                    if (File.Exists(exePath)) { Process.Start(exePath); }
                    else { Utilities.ErrorMsgTaskDlg(Handle, "soffice.exe wurde nicht gefunden", exePath); }
                }
                else { Utilities.ErrorMsgTaskDlg(Handle, "LibreOffice-Installationspfad nicht gefunden.", @"Computer\HKEY_LOCAL_MACHINE\SOFTWARE\LibreOffice\UNO\InstallPath"); }
            }
            catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "Fehler beim Starten von LibreOffice Writer", ex.Message); }
        }
    }

    private void WordProcess()
    {
        FillDictionary();
        wordDoc = null;
        wordApp = null;
        try
        {
            if (Process.GetProcessesByName("WINWORD").Length == 0)
            {
                wordApp = new Word.Application { Visible = true };
                wordApp?.Dialogs[Word.WdWordDialog.wdDialogFileNew].Show(); // Anzeigen der Vorlagenauswahl
                wordApp?.Activate();
                return;
            }
            wordApp ??= (Word.Application)Marshal2.GetActiveObject("Word.Application");
            if (wordApp == null)
            {
                Utilities.WordInfoTaskDlg(Handle, [.. addBookDict.Keys], new(Resources.word32));
                return;
            }
            else
            {
                var hwnd = new IntPtr(wordApp.ActiveWindow.Hwnd); // int-Handle korrekt in einen IntPtr konvertieren
                if (hwnd == IntPtr.Zero) { hwnd = Process.GetProcessesByName("WINWORD")[0].MainWindowHandle; } // Fallback
                NativeMethods.SetForegroundWindow(hwnd);
                wordApp.Activate();
            }
            if (wordApp.Documents.Count >= 1) { wordDoc = wordApp.ActiveDocument; }
            if (wordDoc == null)
            {
                Utilities.WordInfoTaskDlg(Handle, [.. addBookDict.Keys], new(Resources.word32));
                return;
            }
            foreach (var entry in addBookDict)
            {
                var bookmark = entry.Key;
                if (wordDoc.Bookmarks.Exists(bookmark))
                {
                    var bm = wordDoc.Bookmarks[bookmark];
                    var range = bm.Range;
                    range.Text = entry.Value;
                    wordDoc.Bookmarks.Add(bookmark, range);
                }
            }
            string[] arrayOfAllKeys = [.. addBookDict.Keys];
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); } //  + Environment.NewLine + ex.StackTrace
        finally
        {
            if (wordDoc != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                wordDoc = null;
            }
            if (wordApp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                wordApp = null;
            }
            GC.Collect();
        }
    }
    private void FillDictionary()
    {
        addBookDict["Anrede"] = cbAnrede.Text;
        addBookDict["Präfix"] = cbPräfix.Text;
        addBookDict["Vorname"] = tbVorname.Text;
        addBookDict["Zwischenname"] = tbZwischenname.Text;
        addBookDict["Nickname"] = tbNickname.Text;
        addBookDict["Nachname"] = tbNachname.Text;
        addBookDict["Präfix_Zwischenname_Nachname"] = cbPräfix.Text + (cbPräfix.Text.Length > 0 ? " " : "") + tbZwischenname.Text + (tbZwischenname.Text.Length > 0 ? " " : "") + tbNachname.Text;
        addBookDict["Vorname_Zwischenname_Nachname"] = cbPräfix.Text + (cbPräfix.Text.Length > 0 ? " " : "") + tbZwischenname.Text + (tbZwischenname.Text.Length > 0 ? " " : "") + tbNachname.Text;
        addBookDict["Präfix_Vorname_Zwischenname_Nachname"] = cbPräfix.Text + (cbPräfix.Text.Length > 0 ? " " : "") + tbVorname.Text + (tbVorname.Text.Length > 0 ? " " : "") + tbZwischenname.Text + (tbZwischenname.Text.Length > 0 ? " " : "") + tbNachname.Text;
        addBookDict["Anrede_Präfix_Vorname_Zwischenname_Nachname"] = cbAnrede.Text + (cbAnrede.Text.Length > 0 ? " " : "") + cbPräfix.Text + (cbPräfix.Text.Length > 0 ? " " : "") + tbVorname.Text + (tbVorname.Text.Length > 0 ? " " : "") + tbZwischenname.Text + (tbZwischenname.Text.Length > 0 ? " " : "") + tbNachname.Text;
        addBookDict["Suffix"] = tbSuffix.Text;
        addBookDict["Firma"] = tbFirma.Text;
        addBookDict["StraßeNr"] = tbStraße.Text;
        addBookDict["PLZ"] = cbPLZ.Text;
        addBookDict["Ort"] = cbOrt.Text;
        addBookDict["PLZ_Ort"] = cbPLZ.Text + (cbPLZ.Text.Length > 0 ? " " : "") + cbOrt.Text;
        addBookDict["Land"] = cbLand.Text;
        addBookDict["Betreff"] = tbBetreff.Text;
        addBookDict["Grußformel"] = cbGrußformel.Text;
        addBookDict["Schlussformel"] = cbSchlussformel.Text;
        addBookDict["Mail1"] = tbMail1.Text;
        addBookDict["Mail2"] = tbMail2.Text;
        addBookDict["Telefon1"] = tbTelefon1.Text;
        addBookDict["Telefon2"] = tbTelefon2.Text;
        addBookDict["Mobil"] = tbMobil.Text;
        addBookDict["Fax"] = tbFax.Text;
        addBookDict["Internet"] = tbInternet.Text;
    }

    private void WordHelpToolStripMenuItem_Click(object sender, EventArgs e)
    {
        foreach (var key in (string[])[.. addBookDict.Keys]) { addBookDict[key] = string.Empty; }
        Utilities.WordInfoTaskDlg(Handle, [.. addBookDict.Keys], new(Resources.word32));
    }

    private void StatusbarToolStripMenuItem_Click(object sender, EventArgs e) => statusStrip.Visible = statusbarToolStripMenuItem.Checked = !statusbarToolStripMenuItem.Checked;
    private void NewToolStripMenuItem_Click(object sender, EventArgs e) => NewTSButton_Click(sender, e);
    private void DuplicateToolStripMenuItem_Click(object sender, EventArgs e) => CopyTSButton_Click(sender, e);
    private void DeleteToolStripMenuItem_Click(object sender, EventArgs e) => DeleteTSButton_Click(sender, e);

    private async Task CreateContactAsync()
    {
        var vorname = tbVorname.Text; // ?? string.Empty;
        var nachname = tbNachname.Text; // ?? string.Empty;
        var straße = tbStraße.Text;
        var plz = cbPLZ.Text;
        var ort = cbOrt.Text;
        var taskDialog = new TaskDialogPage
        {
            Caption = "Neuer Kontakt",
            Heading = "Möchten Sie die Änderungen speichern?",
            Text = $"{vorname} {nachname}\n{straße}\n{plz} {ort}",
            Icon = TaskDialogIcon.ShieldBlueBar,
            Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
            SizeToContent = true
        };
        if (TaskDialog.ShowDialog(this, taskDialog) == TaskDialogButton.Yes)
        {
            var currentContactRowIndex = contactNewRowIndex;
            contactNewRowIndex = -1;
            try
            {
                toolStripProgressBar.Style = ProgressBarStyle.Marquee;
                toolStripProgressBar.Visible = true;

                var service = await Utilities.GetPeopleServiceAsync(secretPath, tokenDir);

                List<Name> name =
                [
                    new Name() {
                    HonorificPrefix = string.IsNullOrEmpty(cbPräfix.Text.Trim()) ? "" : cbPräfix.Text,
                    MiddleName = string.IsNullOrEmpty(tbZwischenname.Text.Trim()) ? "" : tbZwischenname.Text,
                    FamilyName = string.IsNullOrEmpty(tbNachname.Text.Trim()) ? "" : tbNachname.Text,
                    GivenName = string.IsNullOrEmpty(tbVorname.Text.Trim()) ? "" : tbVorname.Text,
                    HonorificSuffix = string.IsNullOrEmpty(tbSuffix.Text.Trim()) ? "" : tbSuffix.Text
                },];

                List<Nickname> nickname = [];
                if (!string.IsNullOrEmpty(tbNickname.Text.Trim())) { nickname.Add(new Nickname() { Value = tbNickname.Text }); }

                List<UserDefined> userdefined = [];
                if (!string.IsNullOrEmpty(cbAnrede.Text.Trim())) { userdefined.Add(new UserDefined() { Key = "Anrede", Value = cbAnrede.Text }); }
                if (!string.IsNullOrEmpty(tbBetreff.Text.Trim())) { userdefined.Add(new UserDefined() { Key = "Betreff", Value = tbBetreff.Text }); }
                if (!string.IsNullOrEmpty(cbGrußformel.Text.Trim())) { userdefined.Add(new UserDefined() { Key = "Grußformel", Value = cbGrußformel.Text }); }
                if (!string.IsNullOrEmpty(cbSchlussformel.Text.Trim())) { userdefined.Add(new UserDefined() { Key = "Schlussformel", Value = cbSchlussformel.Text }); }

                List<Organization> organization = [];
                if (!string.IsNullOrEmpty(tbFirma.Text.Trim())) { organization.Add(new Organization() { Name = tbFirma.Text }); }

                List<Address> address = [
                    new Address {
                    StreetAddress = string.IsNullOrEmpty(tbStraße.Text.Trim()) ? "" : tbStraße.Text,
                    PostalCode = string.IsNullOrEmpty(cbPLZ.Text.Trim()) ? "" : cbPLZ.Text,
                    City = string.IsNullOrEmpty(cbOrt.Text.Trim()) ? "" : cbOrt.Text,
                    Country = string.IsNullOrEmpty(cbLand.Text.Trim()) ? "" : cbLand.Text,
                },];

                List<Birthday> birthday = [];

                if (!string.IsNullOrEmpty(maskedTextBox.Text) && DateTime.TryParseExact(maskedTextBox.Text, formats, culture, DateTimeStyles.None, out var geburtsdatum))
                {
                    birthday.Add(new Birthday { Date = new Date { Day = geburtsdatum.Day, Month = geburtsdatum.Month, Year = geburtsdatum.Year } });
                }

                List<EmailAddress> emailAddress = [];
                if (!string.IsNullOrEmpty(tbMail1.Text.Trim())) { emailAddress.Add(new EmailAddress { Value = tbMail1.Text, Type = "home" }); }
                if (!string.IsNullOrEmpty(tbMail2.Text.Trim())) { emailAddress.Add(new EmailAddress { Value = tbMail2.Text, Type = "work" }); }

                List<PhoneNumber> phoneNumber = [];
                if (!string.IsNullOrEmpty(tbTelefon1.Text.Trim())) { phoneNumber.Add(new PhoneNumber { Value = tbTelefon1.Text, Type = "home" }); }
                if (!string.IsNullOrEmpty(tbTelefon2.Text.Trim())) { phoneNumber.Add(new PhoneNumber { Value = tbTelefon2.Text, Type = "work" }); }
                if (!string.IsNullOrEmpty(tbMobil.Text.Trim())) { phoneNumber.Add(new PhoneNumber { Value = tbMobil.Text, Type = "mobile" }); }
                if (!string.IsNullOrEmpty(tbFax.Text.Trim())) { phoneNumber.Add(new PhoneNumber { Value = tbFax.Text, Type = "fax" }); }
                List<Url> url = [];
                if (!string.IsNullOrEmpty(tbInternet.Text.Trim())) { url.Add(new Url { Value = tbInternet.Text }); }
                List<Biography> biography = [];
                if (!string.IsNullOrEmpty(tbNotizen.Text.Trim())) { biography.Add(new Biography { Value = tbNotizen.Text }); }

                var person = new Person
                {
                    Names = name.Count > 0 ? name : null,
                    UserDefined = userdefined.Count > 0 ? userdefined : null,
                    Organizations = organization.Count > 0 ? organization : null,
                    Addresses = address.Count > 0 ? address : null,
                    Birthdays = birthday.Count > 0 ? birthday : null,
                    EmailAddresses = emailAddress.Count > 0 ? emailAddress : null,
                    PhoneNumbers = phoneNumber.Count > 0 ? phoneNumber : null,
                    Urls = url.Count > 0 ? url : null,
                    Biographies = biography.Count > 0 ? biography : null
                };
                var response = await service.People.CreateContact(person).ExecuteAsync();
                if (!string.IsNullOrEmpty(response.ResourceName))
                {
                    saveTSButton.Enabled = false;
                    if (topAlignZoomPictureBox.Image != null)
                    {
                        using var photoStream = new MemoryStream();
                        var image = topAlignZoomPictureBox.Image;
                        image.Save(photoStream, image.RawFormat);  // Originalformat (z.B. PNG oder JPEG)
                        var base64Photo = Convert.ToBase64String(photoStream.ToArray());
                        var peopleService = await Utilities.GetPeopleServiceAsync(secretPath, tokenDir);
                        var updatePhotoRequest = new UpdateContactPhotoRequest
                        {
                            PhotoBytes = base64Photo, // erwartet Base64-codierten String der Bild-Bytes
                            PersonFields = "photos"  // explizit die gewünschten Felder anfordern
                        };
                        var request = peopleService.People.UpdateContactPhoto(updatePhotoRequest, response.ResourceName);
                        var photoResponse = await request.ExecuteAsync();
                        if (photoResponse?.Person?.Photos != null && photoResponse.Person.Photos.Any())
                        {
                            var photoUrl = photoResponse.Person.Photos.FirstOrDefault()?.Url;
                            contactDGV.Rows[currentContactRowIndex].Cells["PhotoURL"].Value = photoUrl;
                        }
                    }
                }

            }
            catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "CreateContactAsync: " + ex.GetType().ToString(), ex.Message); }
            finally
            {
                toolStripProgressBar.Visible = false;
                toolStripProgressBar.Style = ProgressBarStyle.Blocks;
            }
        }
        else
        {
            contactNewRowIndex = -1;
            saveTSButton.Enabled = false;
        }
    }

    private async Task UpdateContactAsync(string ressource, Action onClose)
    {
        try
        {
            var service = await Utilities.GetPeopleServiceAsync(secretPath, tokenDir);
            HashSet<string> personFields = []; // HashSet, um Duplikate zu vermeiden
            var getRequest = service.People.Get(ressource);
            getRequest.PersonFields = "names,memberships,nicknames,userDefined,organizations,addresses,birthdays,emailAddresses,phoneNumbers,urls,biographies";
            var person = await getRequest.ExecuteAsync();
            List<string> emptyResourceNames = [];
            var nameKeys = new[] { "Präfix", "Vorname", "Zwischenname", "Nickname", "Nachname", "Suffix" };
            if (nameKeys.Any(changedContactData.ContainsKey))
            {
                person.Names ??= [];
                var name = person.Names.FirstOrDefault();
                if (name == null)
                {
                    name = new Name();
                    person.Names.Add(name);
                }
                if (changedContactData.TryGetValue("Präfix", out var prefix)) { name.HonorificPrefix = prefix; }
                if (changedContactData.TryGetValue("Vorname", out var given)) { name.GivenName = given; }
                if (changedContactData.TryGetValue("Zwischenname", out var middle)) { name.MiddleName = middle; }
                if (changedContactData.TryGetValue("Nachname", out var family)) { name.FamilyName = family; }
                if (changedContactData.TryGetValue("Suffix", out var suffix)) { name.HonorificSuffix = suffix; }
                personFields.Add("names");
            }

            if (changedContactData.TryGetValue("Nickname", out var nick))
            {
                person.Nicknames ??= [];
                var existing = person.Nicknames.FirstOrDefault();
                if (existing == null) { existing = new Nickname(); person.Nicknames.Add(existing); }
                existing.Value = nick;
                personFields.Add("nicknames");
            }

            var addressKeys = new[] { "Straße", "PLZ", "Ort", "Land" };
            if (addressKeys.Any(changedContactData.ContainsKey))
            {
                person.Addresses ??= [];
                var address = person.Addresses.FirstOrDefault(); // Annahme: Es wird nur die erste Adresse verwaltet
                if (address == null)
                {
                    address = new Address();
                    person.Addresses.Add(address);
                }
                if (changedContactData.TryGetValue("Straße", out var street)) { address.StreetAddress = street; }
                if (changedContactData.TryGetValue("PLZ", out var postal)) { address.PostalCode = postal; }
                if (changedContactData.TryGetValue("Ort", out var city)) { address.City = city; }
                if (changedContactData.TryGetValue("Land", out var country)) { address.Country = country; }
                personFields.Add("addresses");
            }

            void updateUserDefined(string key)
            {
                if (changedContactData.TryGetValue(key, out var value))
                {
                    person.UserDefined ??= [];
                    var existing = person.UserDefined.FirstOrDefault(ud => ud.Key == key);
                    if (existing != null)
                    {
                        if (!string.IsNullOrWhiteSpace(value)) { existing.Value = value; }
                        else { person.UserDefined.Remove(existing); }
                    }
                    else if (!string.IsNullOrWhiteSpace(value)) { person.UserDefined.Add(new UserDefined { Key = key, Value = value }); }
                    personFields.Add("userDefined");
                }
            }
            updateUserDefined("Anrede");
            updateUserDefined("Betreff");
            updateUserDefined("Grußformel");
            updateUserDefined("Schlussformel");

            void updateMail(string key, string type)
            {
                if (changedContactData.TryGetValue(key, out var mailValue))
                {
                    person.EmailAddresses ??= [];
                    var existing = person.EmailAddresses.FirstOrDefault(p => p.Type == type);
                    if (existing != null)
                    {
                        if (!string.IsNullOrWhiteSpace(mailValue)) { existing.Value = mailValue; }
                        else { person.EmailAddresses.Remove(existing); }
                    }
                    else if (!string.IsNullOrWhiteSpace(mailValue)) { person.EmailAddresses.Add(new EmailAddress { Value = mailValue, Type = type }); }
                    personFields.Add("emailAddresses");
                }
            }
            updateMail("Mail1", "home");
            updateMail("Mail2", "work");

            void updatePhone(string key, string type)
            {
                if (changedContactData.TryGetValue(key, out var phoneValue))
                {
                    person.PhoneNumbers ??= [];
                    var existing = person.PhoneNumbers.FirstOrDefault(p => p.Type == type);
                    if (existing != null)
                    {
                        if (!string.IsNullOrWhiteSpace(phoneValue)) { existing.Value = phoneValue; }
                        else { person.PhoneNumbers.Remove(existing); }
                    }
                    else if (!string.IsNullOrWhiteSpace(phoneValue)) { person.PhoneNumbers.Add(new PhoneNumber { Value = phoneValue, Type = type }); }
                    personFields.Add("phoneNumbers");
                }
            }
            updatePhone("Telefon1", "home");
            updatePhone("Telefon2", "work");
            updatePhone("Mobil", "mobile");
            updatePhone("Fax", "fax");

            if (changedContactData.TryGetValue("Firma", out var organization))
            {
                person.Organizations ??= [];
                var org = person.Organizations.FirstOrDefault();
                if (org == null) { org = new Organization(); person.Organizations.Add(org); }
                org.Name = organization;
                personFields.Add("organizations");
            }

            if (changedContactData.TryGetValue("Geburtstag", out var geburtstag) && DateTime.TryParse(geburtstag, out var date))
            {
                person.Birthdays ??= [];
                var birthday = person.Birthdays.FirstOrDefault();
                if (birthday == null) { birthday = new Birthday(); person.Birthdays.Add(birthday); }
                birthday.Date = new Date { Day = date.Day, Month = date.Month, Year = date.Year };
                personFields.Add("birthdays");
            }

            if (changedContactData.TryGetValue("Internet", out var internet))
            {
                person.Urls ??= [];
                var existing = person.Urls.FirstOrDefault(u => u.Value == internet);
                if (existing != null) { existing.Value = internet; }
                else if (person.Urls.Count > 0 && person.Urls[0].Value == null) { person.Urls[0].Value = internet; } // Update first Url if it is empty
                personFields.Add("urls");
            }

            if (changedContactData.TryGetValue("Notizen", out var notes))
            {
                person.Biographies ??= [];
                var bio = person.Biographies.FirstOrDefault();
                if (bio == null) { bio = new Biography(); person.Biographies.Add(bio); }
                bio.Value = notes;
                personFields.Add("biographies");
            }

            if (curContactMemberships.Remove("★")) { curContactMemberships.Add("starred"); } // Remove ist nur erfolgreich, wenn ★ enthalten ist
            var desiredGroupResourceNames = new HashSet<string>();
            var nameToResourceNameDict = contactGroupsDict.ToDictionary(kvp => kvp.Value, kvp => kvp.Key); // Umkehrung: Name -> ResourceName
            foreach (var groupName in curContactMemberships)  // Übersetze Klarnamen in ResourceNames. Lege Gruppen an, falls sie nicht existieren.
            {
                if (nameToResourceNameDict.TryGetValue(groupName, out var resourceName)) { desiredGroupResourceNames.Add(resourceName); } // Die Gruppe existiert bereits, füge ihre ResourceName hinzu.
                else
                {
                    var newResourceName = await CreateContactGroupAsync(service, groupName);  // Die Gruppe existiert NICHT. Wir müssen sie neu anlegen.
                    if (!string.IsNullOrEmpty(newResourceName))
                    {
                        desiredGroupResourceNames.Add(newResourceName);
                        nameToResourceNameDict[groupName] = newResourceName;
                    }
                }
            }
            desiredGroupResourceNames.Add("contactGroups/myContacts"); // grundlegende System-Mitgliedschaft als Teil des Soll-Zustands!
            person.Memberships ??= [];  // Sicherstellen, dass die Liste initialisiert ist.
            var existingGroupResourceNames = person.Memberships.Select(m => m.ContactGroupMembership?.ContactGroupResourceName).Where(name => !string.IsNullOrEmpty(name)).ToHashSet();
            if (!existingGroupResourceNames.SetEquals(desiredGroupResourceNames)) // Vergleiche den Ist-Zustand mit dem Soll-Zustand.
            {
                for (var i = person.Memberships.Count - 1; i >= 0; i--)
                {
                    var membership = person.Memberships[i];
                    var groupResourceName = membership.ContactGroupMembership?.ContactGroupResourceName;
                    if (string.IsNullOrEmpty(groupResourceName) || !desiredGroupResourceNames.Contains(groupResourceName))
                    {
                        person.Memberships.RemoveAt(i);
                        if (!string.IsNullOrEmpty(groupResourceName) && groupResourceName != "contactGroups/starred") { emptyResourceNames.Add(groupResourceName); }
                    }
                }
                foreach (var groupResourceNameToAdd in desiredGroupResourceNames) // Mitgliedschaften HINZUFÜGEN
                {
                    if (!existingGroupResourceNames.Contains(groupResourceNameToAdd))
                    {
                        var newMembership = new Membership
                        {
                            ContactGroupMembership = new ContactGroupMembership { ContactGroupResourceName = groupResourceNameToAdd }
                        };
                        person.Memberships.Add(newMembership);
                    }
                }
                personFields.Add("memberships"); // zur Update-Maske hinzufügen
            }
            if (personFields.Count > 0)
            {
                var updateRequest = service.People.UpdateContact(person, ressource);
                updateRequest.UpdatePersonFields = Utilities.BuildMask([.. personFields]); // Specify the fields to update
                var result = await updateRequest.ExecuteAsync();
                if (emptyResourceNames.Count > 0) //  !string.IsNullOrWhiteSpace(emptyResourceName))
                {
                    foreach (var name in emptyResourceNames)
                    {
                        var request = service.ContactGroups.Get(name);
                        request.MaxMembers = 1; // prüfen, ob überhaupt Mitglieder vorhanden sind
                        var contactGroup = await request.ExecuteAsync();
                        onClose?.Invoke();

                        if ((contactGroup.MemberResourceNames == null || contactGroup.MemberResourceNames.Count == 0)
                            && Utilities.YesNo_TaskDialog(Handle, "Google Kontakte", heading: "Möchten Sie die Gruppe löschen?", text: "Die Gruppe ist leer, sie hat keine Mitglieder.", new(Resources.question32)))
                        {
                            await service.ContactGroups.Delete(name).ExecuteAsync();
                        }
                    }
                }
            }
            onClose?.Invoke(); // TaskDialog schließen, sobald die Aktualisierung abgeschlossen ist 
            saveTSButton.Enabled = false;
        }
        catch (Exception ex)
        {
            onClose?.Invoke();
            Utilities.ErrorMsgTaskDlg(Handle, "UpdateContactAsync: " + ex.GetType().ToString(), ex.Message);
        }
        finally
        {
            foreach (var entry in changedContactData) { originalContactData[entry.Key] = entry.Value; } // Überschreibe den Wert, wenn Schlüssel existiert oder füge einen neuen hinzu
            changedContactData.Clear(); // Leeren des Dictionaries nach der Aktualisierung
        }
    }

    private async Task<string?> CreateContactGroupAsync(PeopleServiceService service, string groupName)
    {
        try
        {
            var newGroup = new ContactGroup { Name = groupName };
            var requestBody = new CreateContactGroupRequest { ContactGroup = newGroup };
            var createdGroup = await service.ContactGroups.Create(requestBody).ExecuteAsync();
            //MessageBox.Show(createdGroup.ResourceName + Environment.NewLine + createdGroup.Name, "Gruppe erstellt", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (!contactGroupsDict.ContainsKey(createdGroup.ResourceName)) { contactGroupsDict.Add(createdGroup.ResourceName, createdGroup.Name); } // neue Gruppe zum lokalen Dictionary hinzufügen
            return createdGroup.ResourceName;
        }
        catch (Exception ex)
        {
            Utilities.ErrorMsgTaskDlg(Handle, $"Fehler beim Erstellen der Gruppe '{groupName}'", ex.Message);
            return null; // Gibt null zurück, wenn die Erstellung fehlschlägt.
        }
    }

    private async Task DeleteGoogleContact(int rowIndex)
    {
        try
        {
            toolStripProgressBar.Style = ProgressBarStyle.Marquee;
            toolStripProgressBar.Visible = true;
            //string[] scopes = [PeopleServiceService.Scope.Contacts]; // für OAuth2-Freigabe, mehrere Eingaben mit Komma gerennt (PeopleServiceService.Scope.ContactsOtherReadonly)
            //UserCredential credential;
            //using (FileStream stream = new(secretPath, FileMode.Open, FileAccess.Read))
            //{
            //    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets, scopes, "user", CancellationToken.None, new FileDataStore(tokenDir, true)).Result;
            //}
            //var service = new PeopleServiceService(new BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = appLong, });
            var service = await Utilities.GetPeopleServiceAsync(secretPath, tokenDir);
            await DeleteContactAsync(service, rowIndex);
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "DeleteGoogleContact: " + ex.GetType().ToString(), ex.Message); }
        finally
        {
            toolStripProgressBar.Visible = false;
            toolStripProgressBar.Style = ProgressBarStyle.Blocks;
        }
    }

    private async Task UpdateContactPhotoAsync(string resourceName, Image imageToUpload, ImageFormat formatToUse, Action onClose)
    {
        try
        {
            var service = await Utilities.GetPeopleServiceAsync(secretPath, tokenDir);
            byte[] photoBytes;
            using (var clonedImage = new Bitmap(imageToUpload)) // Workaround für "A generic error occurred in GDI+."-Fehler
            {
                using var ms = new MemoryStream(); // Kopie des Bildes speichern, da das Originalbild noch blockiert sein könnte
                clonedImage.Save(ms, formatToUse);
                photoBytes = ms.ToArray();
            }
            var base64Photo = Convert.ToBase64String(photoBytes);
            var updatePhotoRequest = new UpdateContactPhotoRequest
            {
                PhotoBytes = base64Photo,
                PersonFields = "photos"
            };
            var request = service.People.UpdateContactPhoto(updatePhotoRequest, resourceName);
            var response = await request.ExecuteAsync();
            if (response?.Person?.Photos != null && response.Person.Photos.Any())
            {
                var photoUrl = response.Person.Photos.FirstOrDefault()?.Url;
                contactDGV.Rows[contactDGV.SelectedRows[0].Index].Cells["PhotoURL"].Value = photoUrl;
            }
            onClose?.Invoke(); // TaskDialog schließen, sobald die Aktualisierung abgeschlossen ist 
        }
        catch (Exception ex)
        {
            onClose?.Invoke();
            Utilities.ErrorMsgTaskDlg(Handle, "UpdateContactPhotoAsync: " + ex.GetType().ToString(), ex.Message);
        }
    }

    private async Task DeleteContactPhotoAsync(string resourceName)
    {
        try
        {
            var service = await Utilities.GetPeopleServiceAsync(secretPath, tokenDir);
            var request = service.People.DeleteContactPhoto(resourceName);
            request.PersonFields = "photos"; // Fordert die Photos-Liste in der Antwort an
            var response = await request.ExecuteAsync();
            if (response?.Person != null && response.Person.Photos != null)
            {
                var photo = response.Person.Photos.FirstOrDefault();
                if (photo != null && !string.IsNullOrEmpty(photo.Url))
                {
                    contactDGV.Rows[contactDGV.SelectedRows[0].Index].Cells["PhotoURL"].Value = photo.Url;
                    RenewContactPhoto(photo.Url);
                }
            }
        }
        catch (Google.GoogleApiException gex) when (gex.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "Es ist kein Foto vorhanden!", "Es wurde online nichts gelöscht.", TaskDialogIcon.Information);
            RenewContactPhoto(contactDGV.SelectedRows[0].Cells["PhotoURL"]?.Value?.ToString() ?? "");
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "DeleteContactPhotoAsync: " + ex.GetType().ToString(), ex.Message); }
    }

    private async Task DeleteContactAsync(PeopleServiceService service, int index)
    {
        try
        {
            var resName = contactDGV.Rows[index].Cells["Ressource"]?.Value?.ToString();
            if (!string.IsNullOrEmpty(resName))
            {
                var response = await service.People.DeleteContact(resName).ExecuteAsync();
                if (response != null)
                {
                    ContactEditFields(-1);
                    contactDGV.Rows.RemoveAt(index); // contactDGV.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
                    index = index >= 1 ? index - 1 : 0;
                    if (contactDGV.RowCount > index) { contactDGV.Rows[index].Selected = true; }
                }
            }
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "DeleteContactAsync: " + ex.GetType().ToString(), ex.Message); }
    }

    private async Task LoadAndDisplayGoogleContactsAsync()
    {
        if (tabControl.SelectedTab == addressTabPage && _dataTable != null)
        {
            if (filterRemoveToolStripMenuItem.Visible) { FilterRemoveToolStripMenuItem_Click(null!, null!); }
            if (searchTSTextBox.TextBox.TextLength > 0)
            {
                lastAddressSearch = searchTSTextBox.TextBox.Text;
                ignoreSearchChange = true;
                searchTSTextBox.TextBox.Clear();
                ignoreSearchChange = false;
            }
        }
        else
        {
            if (CheckContactDataChange()) { ShowMultiPageTaskDialog(); }
            lastContactSearch = searchTSTextBox.TextBox.Text;
            ignoreSearchChange = true;
            searchTSTextBox.TextBox.Clear();
            ignoreSearchChange = false;
        }
        if (!Utilities.GoogleConnectionCheck(Handle, secretPath)) { return; }  // Bricht die Methode ab, wenn keine Verbindung besteht
        try
        {
            string[] scopes = [PeopleServiceService.Scope.Contacts, PeopleServiceService.Scope.UserinfoEmail]; //, PeopleServiceService.Scope.UserinfoProfile];
            UserCredential credential;
            using (FileStream stream = new(secretPath, FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets, scopes, "user", CancellationToken.None, new FileDataStore(tokenDir, true));
            }
            var initializer = new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = appLong
            };

            toolStripStatusLabel.Text = string.Empty;
            toolStripProgressBar.Style = ProgressBarStyle.Marquee;
            toolStripProgressBar.Visible = true;

            using Oauth2Service oauthService = new(initializer);
            var emailUserInfo = await oauthService.Userinfo.Get().ExecuteAsync();
            if (emailUserInfo.VerifiedEmail == true) { userEmail = emailUserInfo.Email; }
            using PeopleServiceService service = new(initializer);
            UpdateContactGroupsDict(service); // Lokales Dictionary der Kontaktgruppen aktualisieren

            var peopleRequest = service.People.Connections.List("people/me"); // Kontakte abrufen
            peopleRequest.PersonFields = "names,memberships,nicknames,addresses,phoneNumbers,emailAddresses,biographies,birthdays,urls,organizations,photos,userDefined";
            peopleRequest.SortOrder = (PeopleResource.ConnectionsResource.ListRequest.SortOrderEnum)3; // 3	LAST_NAME_ASCENDING
            peopleRequest.PageSize = 2000;
            var response = await peopleRequest.ExecuteAsync();
            if (response?.Connections == null || !response.Connections.Any())
            {
                toolStripStatusLabel.Text = "Keine Kontakte gefunden.";
                contactDGV.Rows.Clear();
                contactDGV.Columns.Clear();
                return;
            }
            ContactEditFields(-1);
            contactDGV.Rows.Clear();
            contactDGV.Columns.Clear();
            var people = response.Connections;
            if (people != null && people.Count > 0)
            {
                foreach (var field in dataFields.SkipLast(1)) { contactDGV.Columns.Add(field, field); } // ColumnName, HeaderText
                contactDGV.Columns.Add("PhotoURL", "PhotoURL"); // siehe SkipLast(1) 
                contactDGV.Columns.Add("Ressource", "Ressource");
                for (var i = 0; i < contactDGV.Columns.Count - 1; i++) { contactDGV.Columns[i].Visible = !hideColumnArr[i]; }
                foreach (var person in people)
                {
                    var anrede = string.Empty;
                    var betreff = string.Empty;
                    var grußformel = string.Empty;
                    var schlussformel = string.Empty;
                    var fotoUrl = string.Empty;
                    if (person.UserDefined != null && person.UserDefined.Count > 0)
                    {
                        foreach (var customField in person.UserDefined)
                        {
                            if (customField.Key == "Anrede") { anrede = customField.Value ?? string.Empty; }
                            else if (customField.Key == "Betreff") { betreff = customField.Value ?? string.Empty; }
                            else if (customField.Key == "Grußformel") { grußformel = customField.Value ?? string.Empty; }
                            else if (customField.Key == "Schlussformel") { schlussformel = customField.Value ?? string.Empty; }
                        }
                    }
                    if (person.Photos != null && person.Photos.Any())
                    {
                        var photo = person.Photos.FirstOrDefault(p => !string.IsNullOrEmpty(p.Url));
                        if (photo != null)
                        {
                            if (!photo.Default__ ?? true) { fotoUrl = photo.Url; }
                        }
                    }

                    var groupNames = new HashSet<string>();
                    if (person.Memberships != null && person.Memberships.Any())
                    {
                        foreach (var membership in person.Memberships)
                        {
                            if (membership.ContactGroupMembership?.ContactGroupResourceName != null &&
                                contactGroupsDict.TryGetValue(membership.ContactGroupMembership.ContactGroupResourceName, out var groupName))
                            {
                                if (!excludedGroups.Contains(groupName))
                                {
                                    groupName = groupName.Equals("starred") ? "★" : groupName; // "Starred" in "Favoriten" umbenennen   
                                    //allContactMemberships.Add(groupName); // eigentlich unnötig, da schon beim Laden der Gruppen aus ContactGroups
                                    groupNames.Add(groupName);
                                }
                            }
                        }
                    }
                    var membershipsJson = groupNames.Count > 0 ? JsonSerializer.Serialize(groupNames) : string.Empty;

                    contactDGV.Rows.Add(
                        anrede,
                        person.Names != null ? person.Names[0].HonorificPrefix ?? string.Empty : "",
                        person.Names != null ? person.Names[0].FamilyName ?? string.Empty : "",
                        person.Names != null ? person.Names[0].GivenName ?? string.Empty : "",
                        person.Names != null ? person.Names[0].MiddleName ?? string.Empty : "",
                        person.Nicknames != null ? person.Nicknames[0].Value ?? string.Empty : "",
                        person.Names != null ? person.Names[0].HonorificSuffix ?? string.Empty : "",
                        person.Organizations != null ? person.Organizations[0].Name ?? string.Empty : "",
                        person.Addresses != null ? person.Addresses[0].StreetAddress ?? string.Empty : "",
                        person.Addresses != null ? person.Addresses[0].PostalCode ?? string.Empty : "",
                        person.Addresses != null ? person.Addresses[0].City ?? string.Empty : "",
                        person.Addresses != null ? person.Addresses[0].Country ?? string.Empty : "",
                        betreff, // person.Addresses != null ? person.Addresses[0].Type : "",
                        grußformel,
                        schlussformel,
                        person.Birthdays != null ? person.Birthdays[0].Date.Day + "." + person.Birthdays[0].Date.Month + "." + person.Birthdays[0].Date.Year : "",
                        person.EmailAddresses != null ? person.EmailAddresses[0].Value ?? string.Empty : "",
                        person.EmailAddresses != null && person.EmailAddresses.Count > 1 ? person.EmailAddresses[1].Value ?? string.Empty : "",
                        person.PhoneNumbers != null ? Utilities.GetGooglePhoneByType(person, "home") ?? string.Empty : "",
                        person.PhoneNumbers != null ? Utilities.GetGooglePhoneByType(person, "work") ?? string.Empty : "",
                        person.PhoneNumbers != null ? Utilities.GetGooglePhoneByType(person, "mobile") ?? string.Empty : "",
                        person.PhoneNumbers != null ? Utilities.GetGooglePhoneByType(person, "fax") ?? string.Empty : "",
                        person.Urls != null ? person.Urls[0].Value ?? string.Empty : "",
                        person.Biographies != null ? person.Biographies[0].Value.ReplaceLineEndings() ?? string.Empty : "",
                        membershipsJson, // Mitgliedschaften als JSON-Array
                        fotoUrl, // Photo URL
                        person.ResourceName ?? string.Empty
                     );
                }
                allContactMemberships.Add("★"); // "Favoriten (starred)" immer hinzufügen, auch wenn keine Kontakte diese Gruppe haben    
                toolStripStatusLabel.Text = people.Count.ToString() + " Kontakte";
                response.Connections.Clear();  // dispose people 
                foreach (DataGridViewColumn column in contactDGV.Columns) { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
                Utilities.SetColumnWidths(columnWidths, contactDGV);
                tabControl.SelectedIndex = 1; // erst nachdem contactDGV befüllt ist, wg. TabControl_Selecting-Ereignis
                //contactDGV.ResumeLayout(false);
                Text = $"Kontakte - {userEmail}";
                copyTSButton.Enabled = copyToOtherDGVTSMenuItem.Enabled = wordToolStripMenuItem.Enabled = envelopeToolStripMenuItem.Enabled = wordTSButton.Enabled = envelopeTSButton.Enabled = true;
                duplicateToolStripMenuItem.Enabled = false;
                contactDGV.Rows[0].Selected = btnEditContact.Visible = true;
                if (tabulation.TabPages.Contains(tabPageDoku))
                {
                    deactivatedPage = tabPageDoku;
                    tabulation.TabPages.Remove(tabPageDoku);
                }
                cbAnrede.Items.Clear();
                cbPräfix.Items.Clear();
                cbPLZ.Items.Clear();
                cbOrt.Items.Clear();
                cbLand.Items.Clear();
                cbGrußformel.Items.Clear();
                cbSchlussformel.Items.Clear();
                cbAnrede.Items.AddRange([.. contactDGV.Rows.Cast<DataGridViewRow>().Select(row => row.Cells["Anrede"]?.Value).OfType<string>().Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
                cbPräfix.Items.AddRange([.. contactDGV.Rows.Cast<DataGridViewRow>().Select(row => row.Cells["Präfix"]?.Value).OfType<string>().Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
                cbPLZ.Items.AddRange([.. contactDGV.Rows.Cast<DataGridViewRow>().Select(row => row.Cells["PLZ"]?.Value).OfType<string>().Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
                cbOrt.Items.AddRange([.. contactDGV.Rows.Cast<DataGridViewRow>().Select(row => row.Cells["Ort"]?.Value).OfType<string>().Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
                cbLand.Items.AddRange([.. contactDGV.Rows.Cast<DataGridViewRow>().Select(row => row.Cells["Land"]?.Value).OfType<string>().Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
                cbSchlussformel.Items.AddRange([.. contactDGV.Rows.Cast<DataGridViewRow>().Select(row => row.Cells["Schlussformel"]?.Value).OfType<string>().Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
                contactCbItems_Anrede = [.. cbAnrede.Items.Cast<string>()];
                contactCbItems_Präfix = [.. cbPräfix.Items.Cast<string>()];
                contactCbItems_PLZ = [.. cbPLZ.Items.Cast<string>()];
                contactCbItems_Ort = [.. cbOrt.Items.Cast<string>()];
                contactCbItems_Land = [.. cbLand.Items.Cast<string>()];
                contactCbItems_Schlussformel = [.. cbSchlussformel.Items.Cast<string>()];
                ContactEditFields(0);
                if (birthdayShow && birthdayContactShow)
                {
                    toolStripProgressBar.Visible = false;
                    BirthdayReminder();
                }
                birthdayShow = true;
            }
        }
        catch (TokenResponseException)
        {
            birthdayShow = false;
            Utilities.ErrorMsgTaskDlg(Handle, "Autorisierung erforderlich",
            "Das Zugriffstoken ist abgelaufen oder ungültig.\nDer Google-OAuth-Dialog wird beim nächsten Versuch erneut im Browser aufgerufen,\ndort können Sie den Zugriff auf Ihre Kontakte erlauben.",
            TaskDialogIcon.Information);
        }
        catch (Google.GoogleApiException ex) { Utilities.ErrorMsgTaskDlg(Handle, "Google-API-Fehler: " + ex.GetType().ToString(), ex.Message); }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "Allgemeiner Fehler: " + ex.GetType().ToString(), ex.Message); }
        finally // wird immer ausgeführt, auch bei Fehlern
        {
            toolStripProgressBar.Visible = false;
            toolStripProgressBar.Style = ProgressBarStyle.Blocks;
            toolStripStatusLabel.Visible = true;
        }
    }

    private async void GoogleTSButton_Click(object sender, EventArgs e) => await LoadAndDisplayGoogleContactsAsync();

    private void ContactDGV_SelectionChanged(object sender, EventArgs e)
    {
        if (isSelectionChanging) { return; }
        isSelectionChanging = true;
        try
        {
            if (contactDGV.SelectedRows.Count > 0)
            {
                prevSelectedContactRowIndex = contactDGV.SelectedRows[0].Index;
                ContactEditFields(contactDGV.SelectedRows[0].Index); // impliziert btnEditContact.Visible = true;
            }
            else { btnEditContact.Visible = false; }
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "ContactDGV_SelectionChanged: " + ex.GetType().ToString(), ex.Message); }
        finally { isSelectionChanging = false; }
    }

    private async void ContactDGV_Enter(object sender, EventArgs e)
    {
        if (contactDGV.SelectedRows.Count > 0 && contactDGV.SelectedRows[0] != null)
        {
            if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp())
            {
                contactDGV.Rows[prevSelectedContactRowIndex].Selected = true;
                await CreateContactAsync();
                return;
            }
            if (CheckContactDataChange()) { ShowMultiPageTaskDialog(); }
        }
    }

    private async void ContactDGV_CellClick(object sender, DataGridViewCellEventArgs e)
    {
        if ((NativeMethods.GetKeyState(NativeMethods.VK_CONTROL) & 0x8000) != 0 && e.ColumnIndex >= 0)
        {
            var colName = contactDGV.Columns[e.ColumnIndex].Name;
            if (!string.IsNullOrEmpty(colName))
            {
                using (var row = contactDGV.Rows[e.RowIndex])
                {
                    if (!row.Selected) { row.Selected = true; }
                }
                await Task.Delay(50);
                foreach (Control control in tableLayoutPanel.Controls)
                {
                    if (control.Name != null && control.Name.EndsWith(colName))
                    {
                        control.Focus(); // spring auf das der DataGridViewColumn entsprechende Edit-Feld
                        break;
                    }
                }
            }
        }
    }

    private void ContactEditFields(int rowIndex, Image? image = null) // rowIndex = -1 => ClearFields
    {
        try
        {
            ignoreTextChange = true; // verhindert, dass TextChanged
            foreach (var (ctrl, colText) in dictEditField) { ctrl.Text = rowIndex < 0 ? "" : contactDGV.Rows[rowIndex].Cells[colText]?.Value?.ToString() ?? ""; }

            if (rowIndex >= 0)
            {
                if (image != null)
                {
                    topAlignZoomPictureBox.Image = image;
                    delPictboxToolStripButton.Enabled = true;
                }
                else
                {
                    var photoUrl = contactDGV.Rows[rowIndex].Cells["PhotoURL"]?.Value?.ToString();
                    if (string.IsNullOrEmpty(photoUrl))
                    {
                        topAlignZoomPictureBox.Image = Resources.ContactBild100;
                        delPictboxToolStripButton.Enabled = false;
                    }
                    else { RenewContactPhoto(photoUrl); }
                }
            }
            cbGrußformel.Items.Clear();
            if (rowIndex >= 0) { ErzeugeGrußformeln(); }

            if (rowIndex >= 0 && DateTime.TryParse(contactDGV.Rows[rowIndex].Cells["Geburtstag"]?.Value?.ToString(), out var date))
            {
                maskedTextBox.Text = date.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("de-DE"));
                AgeLabel_SetText(date);
            }
            else
            {
                AgeLabel_DeleteText();
                maskedTextBox.Text = string.Empty;
            }

            if (rowIndex >= 0)
            {
                var membershipsJson = contactDGV.Rows[rowIndex].Cells["Gruppen"].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(membershipsJson))
                {
                    var deserialized = JsonSerializer.Deserialize<List<string>>(membershipsJson) ?? [];
                    curContactMemberships = new SortedSet<string>(deserialized, StringComparer.OrdinalIgnoreCase);
                    allContactMemberships.UnionWith(curContactMemberships);
                    //MessageBox.Show(string.Join(", ", allContactMemberships));
                    UpdateMembershipTags();
                }
                else
                {
                    curContactMemberships.Clear();
                    flowLayoutPanel.Controls.Clear();
                    UpdatePlaceholderVis();
                }
                UpdateMembershipCBox();
            }

            tbNotizen.Text = rowIndex < 0 ? "" : contactDGV.Rows[rowIndex].Cells["Notizen"]?.Value?.ToString() ?? "";
            LinkLabel_Enabled();
            btnEditContact.Visible = true;

            originalContactData.Clear();
            if (rowIndex >= 0)
            {
                foreach (DataGridViewCell cell in contactDGV.Rows[rowIndex].Cells)
                {
                    var columnName = cell.OwningColumn.Name; // Spaltenname als Schlüssel verwenden
                    if (!string.IsNullOrEmpty(columnName)) { originalContactData[columnName] = cell.Value?.ToString() ?? string.Empty; }
                }
            }
        }
        catch (Exception ex)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "ContactEditFields: " + ex.GetType().ToString(), ex.Message);
            Application.Exit();
        }
        finally { ignoreTextChange = false; } // TextChanged wieder aktivieren
    }

    private async void RenewContactPhoto(string photoUrl)
    {
        topAlignZoomPictureBox.Image = null;
        delPictboxToolStripButton.Enabled = false;
        if (!string.IsNullOrEmpty(photoUrl))
        {
            if (Uri.IsWellFormedUriString(photoUrl, UriKind.Absolute))
            {
                try
                {   // Nachfolgenden Code nicht löschen, wird fürs Debugging benötigt
                    //using var response = await HttpService.Client.GetAsync(photoUrl);
                    //response.EnsureSuccessStatusCode();
                    //var contentType = response.Content.Headers.ContentType?.MediaType;
                    //var imageFormat = "Unbekannt";
                    //if (contentType != null)
                    //{
                    //    if (contentType.EndsWith("jpeg") || contentType.EndsWith("jpg")) { imageFormat = "JPEG"; }
                    //    else if (contentType.EndsWith("png")) { imageFormat = "PNG"; }
                    //    else if (contentType.EndsWith("gif")) { imageFormat = "GIF"; }
                    //}
                    //MessageBox.Show($"Das Bildformat ist: {imageFormat}");
                    var imageData = await HttpService.Client.GetByteArrayAsync(photoUrl); // byte[]; photoUrl endet mit =s100 für 100×100 - ist so von Google vorgesehen       
                    if (imageData != null && imageData.Length > 0)
                    {
                        using var ms = new MemoryStream(imageData);
                        topAlignZoomPictureBox.Image = Image.FromStream(ms);
                        delPictboxToolStripButton.Enabled = true;
                    }
                }
                catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "RenewContactPhoto: " + ex.GetType().ToString(), ex.Message); }
            }
        }
    }

    private void LinkLabel_Enabled()
    {
        mail1LinkLabel.Enabled = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,})+)$").IsMatch(tbMail1.Text);
        mail2LinkLabel.Enabled = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,})+)$").IsMatch(tbMail2.Text);
        tel1LinkLabel.Enabled = new Regex(@"^\+?(\([0-9 ]*\))?[-. ]?[0-9 ]+$").IsMatch(tbTelefon1.Text);
        tel2LinkLabel.Enabled = new Regex(@"^\+?(\([0-9 ]*\))?[-. ]?[0-9 ]+$").IsMatch(tbTelefon2.Text);
        mobilLinkLabel.Enabled = new Regex(@"^\+?\(?([0-9]*)\)?[-. ]?([0-9].*)$").IsMatch(tbMobil.Text);
        internetLinkLabel.Enabled = new Regex(@"^((http|https)://|www\.)\S+$").IsMatch(tbInternet.Text);
    }

    private async void TabControl_Selecting(object sender, TabControlCancelEventArgs e)
    {
        if (e.TabPage == contactTabPage && contactDGV.Rows.Count == 0)
        {
            if (Utilities.YesNo_TaskDialog(Handle, "Google Kontakte", heading: "Keine Kontakte vorhanden", text: "Möchten Sie Ihre Kontakte laden?", TaskDialogIcon.ShieldBlueBar))
            {
                await LoadAndDisplayGoogleContactsAsync();
            }
        }
        else if (e.TabPage == addressTabPage && contactDGV.SelectedRows.Count > 0)
        {
            if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp())
            {
                await CreateContactAsync();
                e.Cancel = true; // Abbruch, wenn Daten geändert wurden
            }
            if (CheckContactDataChange())
            {
                ShowMultiPageTaskDialog();
                e.Cancel = true;
            }
        }
        if (filterRemoveToolStripMenuItem.Visible) { FilterRemoveToolStripMenuItem_Click(null!, null!); }
    }

    private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            cbAnrede.Items.Clear();
            cbPräfix.Items.Clear();
            cbPLZ.Items.Clear();
            cbOrt.Items.Clear();
            cbLand.Items.Clear();
            cbSchlussformel.Items.Clear();
            cbAnrede.Items.AddRange([.. addressCbItems_Anrede]);
            cbPräfix.Items.AddRange([.. addressCbItems_Präfix]);
            cbPLZ.Items.AddRange([.. addressCbItems_PLZ]);
            cbOrt.Items.AddRange([.. addressCbItems_Ort]);
            cbLand.Items.AddRange([.. addressCbItems_Land]);
            cbSchlussformel.Items.AddRange([.. addressCbItems_Schlussformel]);

            if (deactivatedPage != null && !tabulation.TabPages.Contains(deactivatedPage))
            {
                tabulation.TabPages.Insert(1, deactivatedPage);
                deactivatedPage = null;
            }
            if (searchTSTextBox.TextBox.TextLength > 0)
            {
                lastContactSearch = searchTSTextBox.Text;
                ignoreSearchChange = true;
                searchTSTextBox.TextBox.Clear();
                ignoreSearchChange = false;
            }
            if (!string.IsNullOrEmpty(lastAddressSearch))
            {
                ignoreSearchChange = true;
                searchTSTextBox.TextBox.Text = lastAddressSearch;
                ignoreSearchChange = false;
                lastAddressSearch = string.Empty;
            }
            //flexiTSStatusLabel.Visible = true;
            if (_dataTable != null && _dataTable.Rows.Count > 0)
            {
                Text = appName + " – " + (string?)(string.IsNullOrEmpty(databaseFilePath) ? "unbenannt" : Utilities.CorrectUNC(databaseFilePath));  // Workaround for UNC-Path
                btnEditContact.Visible = false;
                saveTSButton.Enabled = _dataTable?.GetChanges(DataRowState.Added | DataRowState.Modified | DataRowState.Deleted) != null;
                newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled
                    = deleteTSButton.Enabled = newToolStripMenuItem.Enabled = newTSButton.Enabled = duplicateToolStripMenuItem.Enabled = copyTSButton.Enabled = wordTSButton.Enabled
                    = envelopeTSButton.Enabled = true;
                copyToOtherDGVTSMenuItem.Enabled = false;
                var rowCount = addressDGV.Rows.Count;
                var visibleRowCount = addressDGV.Rows.Cast<DataGridViewRow>().Count(static r => r.Visible);
                toolStripStatusLabel.Text = rowCount == visibleRowCount ? $"{visibleRowCount} Adressen" : $"{visibleRowCount}/{rowCount} Adressen";
                if (addressDGV.SelectedRows.Count > 0) { AddressEditFields(addressDGV.SelectedRows[0].Index); }
            }
            Text = !string.IsNullOrWhiteSpace(databaseFilePath) ? $"Adressen - {databaseFilePath}" : "Adressen";
        }
        if (tabControl.SelectedTab == contactTabPage && contactDGV.RowCount > 1)
        {
            cbAnrede.Items.Clear();
            cbPräfix.Items.Clear();
            cbPLZ.Items.Clear();
            cbOrt.Items.Clear();
            cbLand.Items.Clear();
            cbSchlussformel.Items.Clear();
            cbAnrede.Items.AddRange([.. contactCbItems_Anrede]);
            cbPräfix.Items.AddRange([.. contactCbItems_Präfix]);
            cbPLZ.Items.AddRange([.. contactCbItems_PLZ]);
            cbOrt.Items.AddRange([.. contactCbItems_Ort]);
            cbLand.Items.AddRange([.. contactCbItems_Land]);
            cbSchlussformel.Items.AddRange([.. contactCbItems_Schlussformel]);

            if (tabulation.TabPages.Contains(tabPageDoku))
            {
                deactivatedPage = tabPageDoku;
                tabulation.TabPages.Remove(tabPageDoku);
            }

            if (searchTSTextBox.TextBox.TextLength > 0)
            {
                lastAddressSearch = searchTSTextBox.TextBox.Text;
                ignoreSearchChange = true;
                searchTSTextBox.TextBox.Clear();
                ignoreSearchChange = false;
            }
            if (!string.IsNullOrEmpty(lastContactSearch))
            {
                ignoreSearchChange = true;
                searchTSTextBox.TextBox.Text = lastContactSearch;
                ignoreSearchChange = false;
                lastContactSearch = string.Empty;
            }
            //flexiTSStatusLabel.Visible = false;
            Text = !string.IsNullOrWhiteSpace(userEmail) ? $"Kontakte - {userEmail}" : "Google-Kontakte";
            btnEditContact.Visible = true;
            //saveTSButton.Enabled = changedContactData.Count > 0; // eigentlich immer false, da Änderungen sofort übernommen werden  
            newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled
                = duplicateToolStripMenuItem.Enabled = false;
            copyTSButton.Enabled = newTSButton.Enabled = deleteTSButton.Enabled = copyToOtherDGVTSMenuItem.Enabled = wordTSButton.Enabled = envelopeTSButton.Enabled = true;
            var rowCount = contactDGV.Rows.Count;
            var visibleRowCount = contactDGV.Rows.Cast<DataGridViewRow>().Count(static r => r.Visible);
            toolStripStatusLabel.Text = rowCount == visibleRowCount ? $"{visibleRowCount} Kontakte" : $"{visibleRowCount}/{rowCount} Kontakte";
            if (contactDGV.SelectedRows.Count == 1) { ContactEditFields(contactDGV.SelectedRows[0].Index); }
        }
        flexiTSStatusLabel.Text = string.Empty;
        searchTSTextBox.TextBox.Focus();
    }

    private void AuthentMenuItem_Click(object sender, EventArgs e)
    {
        using TaskDialogIcon questionDialogIcon = new(Resources.question32);
        TaskDialogPage page = new()
        {
            Caption = appCont,
            Heading = "Möchten Sie die Zugangsdaten löschen?",
            Text = "Wenn Sie den Request-Token löschen, können Sie nur nach erneuter Autorisierung Google-Kontakte herunterladen. Hierzu öffnet sich beim nächsten Versuch automatisch die Goolge-Anmeldeseite.",
            Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
            Icon = questionDialogIcon,
            DefaultButton = TaskDialogButton.No
        };
        if (TaskDialog.ShowDialog(this, page) == TaskDialogButton.Yes)
        {
            var tokenFile = Path.Combine(tokenDir, "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            try { if (File.Exists(tokenFile)) { File.Delete(tokenFile); } }
            catch (Exception ex)
            {
                var msg = Environment.NewLine + tokenFile + " konnte nicht gelöscht werden.";
                Utilities.ErrorMsgTaskDlg(Handle, ex.Message, msg, TaskDialogIcon.ShieldErrorRedBar);
            }
        }
    }

    private void ExtraToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        authentMenuItem.Enabled = Directory.Exists(tokenDir);
        manageGroupsToolStripMenuItem.Enabled = tabControl.SelectedTab == contactTabPage ? contactDGV.Rows.Count > 0 : _dataTable != null;
    }

    private void BrowserPeopleMenuItem_Click(object sender, EventArgs e)
    {
        try
        {
            ProcessStartInfo psi = new("https://contacts.google.com/") { UseShellExecute = true };
            Process.Start(psi);
        }
        catch (Exception ex) when (ex is Win32Exception || ex is InvalidOperationException) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
    }

    private async void GoogleToolStripMenuItem_ClickAsync(object sender, EventArgs e) => await LoadAndDisplayGoogleContactsAsync();

    private void EnvelopeTSButton_Click(object sender, EventArgs e)
    {
        Cursor = Cursors.WaitCursor; // Aktiviert den Wartesymbol
        FillDictionary();
        using var frm = new FrmPrintSetting(sColorScheme, addBookDict,
            pDevice, pSource, pLandscape,
            pFormat, pFont, pSenderSize, pRecipSize,
            pSenderIndex, pSenderLines1 ??= [], pSenderLines2 ??= [], pSenderLines3 ??= [], pSenderLines4 ??= [], pSenderLines5 ??= [], pSenderLines6 ??= [], pSenderPrint,
            pRecipX, pRecipY, pSendX, pSendY, pRecipBold, pSendBold, pSalutation, pCountry);
        Cursor = Cursors.Default; // Setzt den Cursor auf den Standardwert zurück
        if (frm.ShowDialog() == DialogResult.OK)
        {
            pDevice = frm.Device;
            pSource = frm.Source;
            pLandscape = frm.Landscape;
            pFormat = frm.Format;
            pFont = frm.Schrift;
            pSenderSize = frm.SenderSize;
            pRecipSize = frm.RecipSize;
            pSenderIndex = frm.SenderIndex;
            pSenderLines1 = frm.SenderLines1;
            pSenderLines2 = frm.SenderLines2;
            pSenderLines3 = frm.SenderLines3;
            pSenderLines4 = frm.SenderLines4;
            pSenderLines5 = frm.SenderLines5;
            pSenderLines6 = frm.SenderLines6;
            pSenderPrint = frm.SenderPrint;
            pRecipX = frm.RecipX;
            pRecipY = frm.RecipY;
            pSendX = frm.SendX;
            pSendY = frm.SendY;
            pRecipBold = frm.RecipBold;
            pSendBold = frm.SendBold;
            pSalutation = frm.Salutation;
            pCountry = frm.Country;
        }
    }

    private void OptionsToolStripMenuItem_Click(object sender, EventArgs e)
    {
        using var frm = new FrmProgSettings();
        frm.AskBeforeDelete = sAskBeforeDelete;
        frm.ColorSchemeBlue = sColorScheme == "blue";
        frm.ColorSchemePale = sColorScheme == "pale";
        frm.ColorSchemeDark = sColorScheme == "dark";
        frm.NoFile = sNoAutoload;
        frm.ReloadRecent = sReloadRecent;
        frm.WordProcProg = sWordProcProg;
        frm.StandardFile = sStandardFile;
        frm.DatabaseFolder = sDatabaseFolder;
        frm.ContactsAutoload = sContactsAutoload;
        frm.AskBeforeSaveSQL = sAskBeforeSaveSQL;
        frm.DailyBackup = sDailyBackup;
        frm.WatchFolder = sWatchFolder;
        frm.BackupDirectory = sBackupDirectory;
        frm.LetterDirectory = sLetterDirectory;
        frm.BackupSuccess = sBackupSuccess;
        frm.NumUpDownSuccess.Value = sSuccessDuration > frm.NumUpDownSuccess.Maximum || sSuccessDuration < frm.NumUpDownSuccess.Minimum ? 2500 : sSuccessDuration; // falls in Konfig-Datei Unsinn steht
        //MessageBox.Show(sNoFile.ToString() + Environment.NewLine + sReloadRecent.ToString() + Environment.NewLine + !string.IsNullOrEmpty(sStandardFile));
        if (frm.ShowDialog() == DialogResult.OK)
        {
            sAskBeforeDelete = frm.AskBeforeDelete;
            sColorScheme = frm.ColorSchemeBlue ? "blue" : frm.ColorSchemePale ? "pale" : frm.ColorSchemeDark ? "dark" : "grey";
            sWordProcProg = frm.WordProcProg;
            sReloadRecent = frm.ReloadRecent;
            sStandardFile = frm.StandardFile;
            sNoAutoload = frm.NoFile || (!sReloadRecent && string.IsNullOrEmpty(sStandardFile));
            //MessageBox.Show(sNoFile.ToString() + Environment.NewLine + sReloadRecent.ToString() + Environment.NewLine + !string.IsNullOrEmpty(sStandardFile));
            sDatabaseFolder = frm.DatabaseFolder;
            sContactsAutoload = frm.ContactsAutoload;
            sAskBeforeSaveSQL = frm.AskBeforeSaveSQL;
            sDailyBackup = frm.DailyBackup;
            sWatchFolder = frm.WatchFolder;
            sBackupDirectory = frm.BackupDirectory;
            sLetterDirectory = frm.LetterDirectory;
            sBackupSuccess = frm.BackupSuccess;
            sSuccessDuration = frm.NumUpDownSuccess.Value;
            SetColorScheme();
            fileSystemWatcher.Path = sLetterDirectory;
            if (sWatchFolder && !string.IsNullOrEmpty(sLetterDirectory) && Directory.Exists(sLetterDirectory)) { fileSystemWatcher.EnableRaisingEvents = true; }
            else { fileSystemWatcher.EnableRaisingEvents = false; }
        }
    }

    private void SetColorScheme()
    {
        switch (sColorScheme)
        {
            case "blue":
                menuStrip.BackColor = SystemColors.GradientInactiveCaption;
                menuStrip.ForeColor = SystemColors.ControlText;
                toolStrip.BackColor = SystemColors.GradientInactiveCaption;
                toolStrip.ForeColor = SystemColors.ControlText;
                statusStrip.BackColor = SystemColors.GradientInactiveCaption;
                statusStrip.ForeColor = SystemColors.ControlText;
                tableLayoutPanel.BackColor = SystemColors.InactiveBorder;
                fileToolStripMenuItem.ForeColor = editToolStripMenuItem.ForeColor = viewToolStripMenuItem.ForeColor = extraToolStripMenuItem.ForeColor = helpToolStripMenuItem.ForeColor = SystemColors.ControlText;
                break;
            case "pale":
                menuStrip.BackColor = SystemColors.ControlLightLight;
                menuStrip.ForeColor = SystemColors.ControlText;
                toolStrip.BackColor = SystemColors.ControlLightLight;
                toolStrip.ForeColor = SystemColors.ControlText;
                statusStrip.BackColor = SystemColors.ControlLightLight;
                statusStrip.ForeColor = SystemColors.ControlText;
                tableLayoutPanel.BackColor = SystemColors.ControlLightLight;
                fileToolStripMenuItem.ForeColor = editToolStripMenuItem.ForeColor = viewToolStripMenuItem.ForeColor = extraToolStripMenuItem.ForeColor = helpToolStripMenuItem.ForeColor = SystemColors.ControlText;
                break;
            case "dark":
                menuStrip.BackColor = SystemColors.ControlDark;
                menuStrip.ForeColor = SystemColors.HighlightText;
                toolStrip.BackColor = SystemColors.ControlDark;
                toolStrip.ForeColor = SystemColors.HighlightText;
                statusStrip.BackColor = SystemColors.ControlDark;
                statusStrip.ForeColor = SystemColors.HighlightText;
                tableLayoutPanel.BackColor = SystemColors.Control;
                fileToolStripMenuItem.ForeColor = editToolStripMenuItem.ForeColor = viewToolStripMenuItem.ForeColor = extraToolStripMenuItem.ForeColor = helpToolStripMenuItem.ForeColor = SystemColors.HighlightText;
                break;
            default:
                menuStrip.BackColor = SystemColors.Control;
                menuStrip.ForeColor = SystemColors.ControlText;
                toolStrip.BackColor = SystemColors.Control;
                toolStrip.ForeColor = SystemColors.ControlText;
                statusStrip.BackColor = SystemColors.Control;
                statusStrip.ForeColor = SystemColors.ControlText;
                tableLayoutPanel.BackColor = SystemColors.ButtonFace;
                fileToolStripMenuItem.ForeColor = editToolStripMenuItem.ForeColor = viewToolStripMenuItem.ForeColor = extraToolStripMenuItem.ForeColor = helpToolStripMenuItem.ForeColor = SystemColors.ControlText;
                break;
        }
    }

    private void BtnEditContact_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            var ressourceCell = contactDGV.SelectedRows[0].Cells["Ressource"];
            var ressourceText = ressourceCell?.Value?.ToString();
            if (!string.IsNullOrEmpty(ressourceText))
            {
                ressourceText = ressourceText.Replace("people", "person");
                try
                {
                    ProcessStartInfo psi = new("https://contacts.google.com/" + ressourceText) { UseShellExecute = true };
                    Process.Start(psi);
                }
                catch (Exception ex) when (ex is Win32Exception || ex is InvalidOperationException) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
            }
            else { Utilities.ErrorMsgTaskDlg(Handle, "Kein Ressource-Wert", "Der Ressource-Wert ist leer oder null."); }
        }
        else { Console.Beep(); }
    }

    private void TsClearLabel_Click(object sender, EventArgs e) => Clear_SearchTextBox();

    private void TsClearLabel_VisibleChanged(object sender, EventArgs e) => searchTSTextBox.Width = 202 + splitContainer.SplitterDistance - 536 - (tsClearLabel.Visible ? tsClearLabel.Width : 0);

    private void TsClearLabel_Paint(object sender, PaintEventArgs e) => BeginInvoke(new Action(() => Graphics.FromHwnd(toolStrip.Handle).DrawRectangle(Pens.Black, tsClearLabel.Bounds.Location.X - 2, tsClearLabel.Bounds.Location.Y + 2, tsClearLabel.Width + 1, tsClearLabel.Height - 4)));
    // private void TsClearLabel_Paint(object sender, PaintEventArgs e) => InvokeAsync(() => Graphics.FromHwnd(toolStrip.Handle).DrawRectangle(Pens.Black, tsClearLabel.Bounds.Location.X - 2, tsClearLabel.Bounds.Location.Y + 2, tsClearLabel.Width + 1, tsClearLabel.Height - 4));

    private void AddressDGV_KeyDown(object sender, KeyEventArgs e)
    {
        var keyValue = e.KeyValue;
        if (e.Control && e.KeyCode == Keys.C)
        {
            ClipboardTSMenuItem_Click(null!, null!);
            e.Handled = true; // Prevent default copy behavior
        }
        else if (e.Modifiers == Keys.None && (keyValue >= (int)Keys.A && keyValue <= (int)Keys.Z || e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9))
        {
            searchTSTextBox.Focus();
            searchTSTextBox.Text += e.Shift ? ((char)keyValue).ToString() : ((char)(keyValue + 32)).ToString();
            searchTSTextBox.SelectionStart = searchTSTextBox.Text.Length;  // Cursor ans Ende stellen
        }
    }

    private void ContactDGV_KeyDown(object sender, KeyEventArgs e)
    {
        var keyValue = e.KeyValue;
        if (e.Control && e.KeyCode == Keys.C)
        {
            ClipboardTSMenuItem_Click(null!, null!);
            e.Handled = true; // Prevent default copy behavior
        }
        else if (e.Modifiers == Keys.None && (keyValue >= (int)Keys.A && keyValue <= (int)Keys.Z || e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9))
        {
            searchTSTextBox.Focus();
            searchTSTextBox.Text += e.Shift ? ((char)keyValue).ToString() : ((char)(keyValue + 32)).ToString();
            searchTSTextBox.SelectionStart = searchTSTextBox.Text.Length;  // Cursor ans Ende stellen
        }
    }

    private void SearchTSTextBox_Enter(object sender, EventArgs e) => searchTSTextBox.BackColor = Color.LightYellow;
    private void SearchTSTextBox_Leave(object sender, EventArgs e) => searchTSTextBox.BackColor = Color.White;
    private void ComboBox_Enter(object sender, EventArgs e) => ((ComboBox)sender).BackColor = Color.LightYellow;

    private void ComboBox_Leave(object sender, EventArgs e)
    {
        var ctrl = (Control)sender;
        //if (addressDGV.SelectedRows.Count > 0 && addressDGV.SelectedRows[0].DataBoundItem is DataRowView dataBoundItem) { _bindingSource.EndEdit(); }
        ctrl.BackColor = Color.White;
        CheckSaveButton();
    }

    private void TextBox_MouseDown(object sender, MouseEventArgs e)
    {
        if (sender is TextBox textBox && e.Button == MouseButtons.Left)
        {
            if (!textBoxClicked)
            {
                textBoxClicked = true;
                textBox.SelectAll();
            }
        }
    }

    private void TextBox_Enter(object sender, EventArgs e)
    {
        if (sender is TextBox tb)
        {
            tb.SelectAll(); // Select all text when entering the textbox
            tb.BackColor = Color.LightYellow;
        }
        textBoxClicked = false;
    }

    private void TextBox_Leave(object sender, EventArgs e)
    {
        ((TextBox)sender).BackColor = Color.White;
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            saveTSButton.Enabled = CheckContactDataChange();
        }
    }

    private void MaskedTextBox_Enter(object sender, EventArgs e)
    {
        ignoreTextChange = true;
        maskedTextBox.Mask = "00/00/0000";
        maskedTextBox.BackColor = Color.LightYellow;
        maskedTextBox.BeginInvoke(new Action(maskedTextBox.SelectAll));
        textBoxClicked = false; // Reset the flag when entering the maskedTextBox
        ignoreTextChange = false;
    }

    private void MaskedTextBox_Leave(object sender, EventArgs e)
    {
        var day = string.Empty;
        var month = string.Empty;
        var year = string.Empty;
        var dateComponents = maskedTextBox.Text.Split('.');
        if (dateComponents.Length > 0) { day = dateComponents[0].Trim(); }
        if (dateComponents.Length > 1) { month = dateComponents[1].Trim(); }
        if (dateComponents.Length > 2) { year = dateComponents[2].Trim(); }
        if (day.Length == 1) { day = "0" + day; }
        if (month.Length == 1) { month = "0" + month; }
        if (year.Length == 2) { year = "20" + year; }
        maskedTextBox.Text = day + "." + month + "." + year;

        maskedTextBox.BackColor = panelBirthdayTextbox.BackColor = Color.White;

        if (!maskedTextBox.MaskFull || !DateTime.TryParseExact(maskedTextBox.Text, formats, culture, DateTimeStyles.None, out var _))
        {
            maskedTextBox.Mask = "";
            maskedTextBox.Text = "";
        }
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            saveTSButton.Enabled = CheckContactDataChange();
        }
    }

    private void MaskedTextBox_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left) // !textBoxClicked  &&   
        {
            //textBoxClicked = true;
            var rawDateString = maskedTextBox.Text.Replace(maskedTextBox.PromptChar.ToString(), "").Trim();
            var charIndex = maskedTextBox.GetCharIndexFromPosition(e.Location);
            switch (charIndex)
            {
                case <= 2:
                    if (rawDateString.Length < 2) { break; }
                    maskedTextBox.SelectionStart = 0;
                    maskedTextBox.SelectionLength = 2;
                    break;
                case >= 3 and <= 5:
                    if (rawDateString.Length < 4) { break; }
                    maskedTextBox.SelectionStart = 3;
                    maskedTextBox.SelectionLength = 2;
                    break;
                case >= 5: // and <= 8:
                    if (rawDateString.Length < 8) { break; }
                    maskedTextBox.SelectionStart = 6;
                    maskedTextBox.SelectionLength = 4;
                    break;
            }
        }
    }

    private void MaskedTextBox_TextChanged(object sender, EventArgs e)
    {
        if (ignoreTextChange) { return; }
        if (DateTime.TryParseExact(maskedTextBox.Text, formats, culture, DateTimeStyles.None, out var geburtsdatum))
        {
            if (geburtsdatum > DateTime.Today)
            {
                Utilities.ErrorMsgTaskDlg(Handle, "Ungültiges Geburtsdatum", "Das Geburtsdatum kann nicht in der Zukunft liegen.");
                maskedTextBox.Clear();
                maskedTextBox.Focus();
                return;
            }
            AgeLabel_SetText(geburtsdatum);
            if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0
                && (addressDGV.SelectedRows[0].Cells["Geburtstag"].Value.ToString() == string.Empty
                || (DateTime.TryParse(addressDGV.SelectedRows[0].Cells["Geburtstag"].Value.ToString(), out var foo) && foo != geburtsdatum)))
            {
                //addressDGV.SelectedRows[0].Cells["Geburtstag"].Value = geburtsdatum.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("de-DE"));
                if (addressDGV.SelectedRows[0].DataBoundItem is DataRowView dataBoundItem) { dataBoundItem.Row["Geburtstag"] = geburtsdatum.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("de-DE")); }
            }

            else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0
                && (contactDGV.SelectedRows[0].Cells["Geburtstag"].Value.ToString() == string.Empty
                || (DateTime.TryParse(contactDGV.SelectedRows[0].Cells["Geburtstag"].Value.ToString(), out var bar) && bar != geburtsdatum)))
            {
                contactDGV.SelectedRows[0].Cells["Geburtstag"].Value = geburtsdatum.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("de-DE"));
            }
        }
        else if (string.IsNullOrEmpty(maskedTextBox.Text.Replace(maskedTextBox.PromptChar.ToString(), "").Trim())) // || maskedTextBox.MaskFull)
        {
            if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0 && !string.IsNullOrEmpty(addressDGV.SelectedRows[0].Cells["Geburtstag"].Value.ToString()))
            {
                if (addressDGV.SelectedRows[0].DataBoundItem is DataRowView dataBoundItem) { dataBoundItem.Row["Geburtstag"] = string.Empty; }
                //addressDGV.SelectedRows[0].Cells["Geburtstag"].Value = string.Empty;
            }
            else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0 && !string.IsNullOrEmpty(contactDGV.SelectedRows[0].Cells["Geburtstag"].Value.ToString()))
            {
                contactDGV.SelectedRows[0].Cells["Geburtstag"].Value = string.Empty;
            }
            AgeLabel_DeleteText();
        }
        CheckSaveButton();
    }

    private void OpenCalendar()
    {
        EnsureCalendar();
        if (Utilities.TryParseInput(maskedTextBox.Text, out var current)) { monthCalendar!.SetDate(current); }
        else { monthCalendar!.SetDate(DateTime.Today); }
        var location = new Point(btnCalendar.Width - monthCalendar.Width, btnCalendar.Height); // Dropdown anzeigen, unterhalb des Buttons
        calendarDropdown!.Show(btnCalendar, location);
    }

    private void EnsureCalendar()
    {
        if (monthCalendar == null)
        {
            monthCalendar = new MonthCalendar { MaxSelectionCount = 1, ShowTodayCircle = true };
            monthCalendar.DateSelected += MonthCalendar_DateSelected;
        }
        if (calendarDropdown == null)
        {
            var host = new ToolStripControlHost(monthCalendar) { Margin = Padding = Padding.Empty, AutoSize = false, Size = monthCalendar.Size };
            calendarDropdown = new ToolStripDropDown { AutoClose = true, DropShadowEnabled = true, Padding = Padding.Empty };
            calendarDropdown.Items.Add(host);
            calendarDropdown.Closed += (_, __) => { if (!maskedTextBox.Focused) { maskedTextBox.Focus(); } };  // Fokus zurück ins Feld
        }
    }

    private void MonthCalendar_DateSelected(object? sender, DateRangeEventArgs e)
    {
        var date = e.Start;
        maskedTextBox.Text = date.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("de-DE"));
        calendarDropdown?.Close();
    }

    private void BtnCalendar_Click(object sender, EventArgs e) => OpenCalendar();

    private void NewDBToolStripMenuItem_Click(object sender, EventArgs e)
    {
        try
        {
            saveFileDialog.Title = "Neue Datenbank anlegen";
            saveFileDialog.InitialDirectory = string.IsNullOrEmpty(sDatabaseFolder) || !Directory.Exists(sDatabaseFolder) ? null : sDatabaseFolder;
            saveFileDialog.DefaultExt = "adb";
            saveFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (_dataTable != null) { SaveSQLDatabase(true); }
                databaseFilePath = saveFileDialog.FileName;
            }
            else { return; }
            CreateNewDatabase(databaseFilePath, true);
            ConnectSQLDatabase(databaseFilePath);
        }
        catch (Exception ex)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "NewDBToolStripMenuItem_Click: " + ex.GetType().ToString(), ex.Message);
            databaseFilePath = string.Empty;
        }
    }

    private void ExportToolStripMenuItem_Click(object sender, EventArgs e)
    {
        saveFileDialog.FileName = "Adressen_Export.csv";
        saveFileDialog.DefaultExt = "csv";
        saveFileDialog.Filter = "CSV-Datei (*.csv)|*.csv|Alle Dateien (*.*)|*.*";
        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            if (_dataTable?.Rows.Count > 0)
            {
                try
                {
                    StringBuilder sb = new();
                    sb.AppendLine(string.Join(";", _dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName)));
                    foreach (DataRow row in _dataTable.Rows) { sb.AppendLine(string.Join(";", row.ItemArray.Select(field => string.Concat("\"", field?.ToString()?.Replace("\"", "\"\""), "\"")))); }
                    File.WriteAllText(saveFileDialog.FileName, sb.ToString(), Encoding.UTF8);
                }
                catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
            }
        }

    }

    private void ColumnSelectToolStripMenuItem_Click(object sender, EventArgs e)
    {
        using var frm = new FrmColumns(hideColumnStd, tabControl.SelectedTab == addressTabPage ? "Dokumente" : "PhotoURL");
        for (var i = 0; i < frm.GetColumnList().Items.Count; i++) { frm.GetColumnList().Items[i].Checked = !hideColumnArr[i]; }
        if (frm.ShowDialog() == DialogResult.OK)
        {
            for (var i = 0; i < frm.GetColumnList().Items.Count; i++)
            {
                if (addressDGV.Columns.Count > i) { addressDGV.Columns[i].Visible = frm.GetColumnList().Items[i].Checked; }
                if (contactDGV.Columns.Count > i) { contactDGV.Columns[i].Visible = frm.GetColumnList().Items[i].Checked; }
                hideColumnArr[i] = !frm.GetColumnList().Items[i].Checked;
            }
        }
    }

    private void ColumnWidthsResetToolStripMenuItem_Click(object sender, EventArgs e)
    {
        for (var i = 0; i < addressDGV.Columns.Count; i++)
        {
            if (addressDGV.Columns[i].Name == "Nachname") { addressDGV.Columns[i].Width = 200; }
            else { addressDGV.Columns[i].Width = 100; }
        }
        for (var i = 0; i < contactDGV.Columns.Count; i++)
        {
            if (contactDGV.Columns[i].Name == "Nachname") { contactDGV.Columns[i].Width = 200; }
            else { contactDGV.Columns[i].Width = 100; }
        }
    }

    private void SplitterAutomaticToolStripMenuItem_Click(object sender, EventArgs e) => splitContainer.SplitterDistance = toolStripSeparator.Bounds.Left;

    private void SplitContainer_SplitterMoved(object sender, SplitterEventArgs e)
    {
        //foreach (var box in tableLayoutPanel.Controls.OfType<ComboBox>()) { box.Select(box.Text.Length, 0); }  //Workaround remove highlight from ComboBox, after assigning SelectedValue
        flexiTSStatusLabel.Width = 244 + splitContainer.SplitterDistance - 536;
    }

    private void WordToolStripMenuItem_Click(object sender, EventArgs e) => WordTSButton_Click(sender, e);

    private void EnvelopeToolStripMenuItem_Click(object sender, EventArgs e) => EnvelopeTSButton_Click(sender, e);

    private void ClipboardTSMenuItem_Click(object sender, EventArgs e)
    {
        FillDictionary();
        using var frm = new FrmCopyScheme(sColorScheme, addBookDict, indexCopyPattern, copyPattern1 ?? [], copyPattern2 ?? [], copyPattern3 ?? [], copyPattern4 ?? [], copyPattern5 ?? [], copyPattern6 ?? []);
        if (frm.ShowDialog() == DialogResult.OK)
        {
            copyPattern1 = JoinPatterns(frm.GetPattern1());
            copyPattern2 = JoinPatterns(frm.GetPattern2());
            copyPattern3 = JoinPatterns(frm.GetPattern3());
            copyPattern4 = JoinPatterns(frm.GetPattern4());
            copyPattern5 = JoinPatterns(frm.GetPattern5());
            copyPattern6 = JoinPatterns(frm.GetPattern6());
            indexCopyPattern = frm.PatternIndex;
        }
    }

    private string[] JoinPatterns(string[] patterns)
    {
        if (patterns == null) { return []; }
        var result = new string[patterns.Length];
        for (var i = 0; i < patterns.Length; i++) { result[i] = string.Join(" ", Regex.Matches(patterns[i], @"\b\w+\b").Cast<Match>().Select(m => addBookDict.ContainsKey(m.Value) ? m.Value : string.Empty)).Trim(); }
        return result;
    }

    private void ContextMenu_Opening(object sender, CancelEventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            if (addressDGV.SelectedRows.Count <= 0) { e.Cancel = true; return; }
            else if (!Utilities.RowIsVisible(addressDGV, addressDGV.SelectedRows[0])) { addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.SelectedRows[0].Index; }
            copy2OtherDGVMenuItem.Text = "Zu Google-Kontakte hinzufügen";
            copy2OtherDGVMenuItem.Visible = contactDGV.Rows.Count > 0;
            move2OtherDGVToolStripMenuItem.Visible = false;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            if (contactDGV.SelectedRows.Count <= 0) { e.Cancel = true; return; }
            else if (!Utilities.RowIsVisible(contactDGV, contactDGV.SelectedRows[0])) { contactDGV.FirstDisplayedScrollingRowIndex = contactDGV.SelectedRows[0].Index; }
            copy2OtherDGVMenuItem.Text = "Nach Lokale Adressen kopieren";
            copy2OtherDGVMenuItem.Visible = addressDGV.Rows.Count > 0;
            move2OtherDGVToolStripMenuItem.Visible = contactDGV.Rows.Count > 0;
        }
        //copy2OtherDGVSeparator.Visible = copy2OtherDGVMenuItem.Visible = tabControl.SelectedTab == contactTabPage;
        //dupTSMenuItem.Enabled = tabControl.SelectedTab == addressTabPage;
    }

    private void NewTSMenuItem_Click(object sender, EventArgs e) => NewTSButton_Click(sender, e);
    private void DupTSMenuItem_Click(object sender, EventArgs e) => CopyTSButton_Click(sender, e);
    private void DelTSMenuItem_Click(object sender, EventArgs e) => DeleteTSButton_Click(sender, e);
    private void ClipTSMenuItem_Click(object sender, EventArgs e) => ClipboardTSMenuItem_Click(sender, e);
    private void Copy2OtherDGVMenuItem_Click(object sender, EventArgs e) => CopyToOtherDGVMenuItem_Click(sender, e);
    private void WordTSMenuItem_Click(object sender, EventArgs e) => WordTSButton_Click(sender, e);
    private void EnvelopeTSMenuItem_Click(object sender, EventArgs e) => EnvelopeTSButton_Click(sender, e);


    private void AddressDGV_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right)
        {
            var rowSelected = e.RowIndex;
            if (e.RowIndex != -1)
            {
                addressDGV.ClearSelection();
                addressDGV.Rows[rowSelected].Selected = true;
                if (e.ColumnIndex >= 0) { addressDGV.CurrentCell = addressDGV.Rows[rowSelected].Cells[e.ColumnIndex]; }
                AddressEditFields(rowSelected);
            }
        }
    }

    private void ContactDGV_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right)
        {
            var rowSelected = e.RowIndex;
            if (e.RowIndex != -1)
            {
                contactDGV.ClearSelection();
                contactDGV.Rows[rowSelected].Selected = true;
                if (e.ColumnIndex >= 0) { contactDGV.CurrentCell = contactDGV.Rows[rowSelected].Cells[e.ColumnIndex]; }
                ContactEditFields(rowSelected);
            }
        }
    }

    private void MainToolStripMenuItem_DropDownOpened(object sender, EventArgs e) => ((ToolStripMenuItem)sender).ForeColor = SystemColors.ControlText;

    private void MainToolStripMenuItem_DropDownClosed(object sender, EventArgs e) => ((ToolStripMenuItem)sender).ForeColor = sColorScheme == "dark" ? SystemColors.HighlightText : SystemColors.ControlText;

    private void AddressDGV_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
    {
        if (e.RowIndex >= 0)
        {
            using var row = addressDGV.Rows[e.RowIndex];
            if (searchTSTextBox.TextBox.TextLength == 0) { row.DefaultCellStyle.BackColor = e.RowIndex % 2 == 0 ? Color.FloralWhite : Color.White; }
            else { row.DefaultCellStyle.BackColor = Color.White; }
        }
    }

    private void ContactDGV_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
    {
        if (e.RowIndex >= 0)
        {
            using var row = contactDGV.Rows[e.RowIndex];
            if (searchTSTextBox.TextBox.TextLength == 0) { row.DefaultCellStyle.BackColor = e.RowIndex % 2 == 0 ? Color.AliceBlue : Color.White; }
            else { row.DefaultCellStyle.BackColor = Color.White; }
        }
    }

    private void RejectChangesToolStripMenuItem_Click(object sender, EventArgs e)
    {
        _dataTable?.RejectChanges();
        if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0) { AddressEditFields(addressDGV.SelectedRows[0].Index); }
    }

    private void EditToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        rejectChangesToolStripMenuItem.Enabled = tabControl.SelectedTab == addressTabPage && _dataTable?.GetChanges() != null;
        if (tabControl.SelectedTab == addressTabPage)
        {
            copyToOtherDGVTSMenuItem.Text = "Zu Google-&Kontakte hinzufügen";
            copyToOtherDGVTSMenuItem.Enabled = addressDGV.SelectedRows.Count > 0 && contactDGV.Rows.Count > 0;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            copyToOtherDGVTSMenuItem.Text = "Nach Lokale Adressen &kopieren";
            copyToOtherDGVTSMenuItem.Enabled = contactDGV.SelectedRows.Count > 0 && addressDGV.Rows.Count > 0;
        }
    }

    private void GooglebackupToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (contactDGV.Rows.Count == 0)
        {
            Utilities.ErrorMsgTaskDlg(Handle, "Keine Daten zum Speichern", "Es sind keine Google-Kontaktdaten vohanden.");
            return;
        }
        saveFileDialog.Filter = "SQLite Database File (*.adb)|*.adb|All files (*.*)|*.*"; // using var sfd = new SaveFileDialog();
        saveFileDialog.Title = "Wählen Sie einen Speicherort";
        saveFileDialog.FileName = "GoogleKontakte.adb";
        saveFileDialog.InitialDirectory = Directory.Exists(sDatabaseFolder) ? sDatabaseFolder : Path.GetDirectoryName(databaseFilePath);
        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            var backupPath = saveFileDialog.FileName;
            tabControl.SelectedTab = addressTabPage;
            try
            {
                var readyPage = new TaskDialogPage
                {
                    Caption = appLong,
                    Heading = "Backup erfolgreich",
                    Text = $"Die Google-Kontakte wurden erfolgreich in\n{backupPath} gespeichert.\n\nMöchten Sie die Datei jetzt öffnen?",
                    Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
                    Footnote = "Bitte beachten Sie, dass das Backup insofern unvollständig ist, dass\nnur die in diesem Programm verwendeten Felder gesichert wurden.",
                    AllowCancel = true,
                    Icon = TaskDialogIcon.ShieldSuccessGreenBar,
                    SizeToContent = true
                };

                var inProgressCloseButton = TaskDialogButton.Close;
                inProgressCloseButton.Enabled = false;
                var progressPage = new TaskDialogPage()
                {
                    Caption = appLong,
                    Heading = "Bitte warten…",
                    Text = "Fotos werden heruntergeladen…",
                    Icon = TaskDialogIcon.None,
                    ProgressBar = new TaskDialogProgressBar() { State = TaskDialogProgressBarState.Marquee },
                    Buttons = { inProgressCloseButton }
                };
                progressPage.Created += async (s, e) =>
                {
                    try
                    {
                        await SpeichereKontakteUndFotosAsync(backupPath);
                        progressPage.Navigate(readyPage);
                    }
                    catch (Exception ex)
                    {
                        if (progressPage.BoundDialog != null) { progressPage.BoundDialog?.Close(); } // läuft im UI-Thread
                        var displayException = ex;
                        if (ex is AggregateException aggEx && aggEx.InnerExceptions.Count > 0) { displayException = aggEx.InnerExceptions[0]; }
                        Utilities.ErrorMsgTaskDlg(Handle, displayException.GetType().Name, $"{displayException.Message}\nDer Backupvorgang wird abgebrochen!", TaskDialogIcon.ShieldWarningYellowBar);
                    }
                };
                if (TaskDialog.ShowDialog(Handle, progressPage) == TaskDialogButton.Yes)
                {
                    {
                        if (_dataTable != null) { SaveSQLDatabase(true); }
                        ConnectSQLDatabase(backupPath);
                        ignoreSearchChange = true;
                        searchTSTextBox.TextBox.Clear();
                        ignoreSearchChange = false;
                        if (birthdayAddressShow) { BirthdayReminder(); }
                    }
                }
            }
            catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message, TaskDialogIcon.ShieldErrorRedBar); }
        }
    }

    public async Task SpeichereKontakteUndFotosAsync(string dbPath)
    {
        if (File.Exists(dbPath)) { File.Delete(dbPath); }
        CreateNewDatabase(dbPath, false);
        var fotoInfos = new List<(int KontaktId, string Url)>();
        using var connection = new SQLiteConnection($"Data Source={dbPath};Version=3;");
        await connection.OpenAsync();
        using var transaction = connection.BeginTransaction();
        var columnNames = string.Join(", ", dataFields.SkipLast(1));
        var columnParams = string.Join(", ", dataFields.SkipLast(1).Select(field => "@" + field));
        var insertCommandText = $"INSERT INTO Adressen ({columnNames}) VALUES ({columnParams})";
        foreach (DataGridViewRow row in contactDGV.Rows)
        {
            if (row.IsNewRow) { continue; }
            using var insertCommand = new SQLiteCommand(insertCommandText, connection, transaction);
            foreach (var field in dataFields.SkipLast(1))
            {
                var cellValue = row.Cells[field].Value ?? DBNull.Value;
                insertCommand.Parameters.AddWithValue($"@{field}", cellValue);
            }
            await insertCommand.ExecuteNonQueryAsync();
            using var cmdGetId = new SQLiteCommand("SELECT last_insert_rowid()", connection, transaction);
            var kontaktId = Convert.ToInt32(await cmdGetId.ExecuteScalarAsync());
            if (row.Cells["PhotoURL"].Value is string photoUrl && !string.IsNullOrWhiteSpace(photoUrl)) { fotoInfos.Add((kontaktId, photoUrl)); }
        }
        transaction.Commit();
        connection.Close();
        await LadeUndSpeichereFotosAsync(fotoInfos, dbPath); // Phase 2: Fotos laden und speichern
    }

    public static async Task LadeUndSpeichereFotosAsync(IEnumerable<(int KontaktId, string Url)> fotos, string dbPath)
    {
        var fotoList = fotos.ToList();
        var total = fotoList.Count;
        var done = 0;
        using var connection = new SQLiteConnection($"Data Source={dbPath};Version=3;");
        await connection.OpenAsync();
        using var transaction = connection.BeginTransaction();
        await Parallel.ForEachAsync(fotoList, new ParallelOptions { MaxDegreeOfParallelism = 8 }, async (f, ct) =>
            {
                try
                {
                    var fotodaten = await LadeFotoAsync(f.Url, ct);
                    if (fotodaten?.Length > 0) { await SpeichereFotoFuerKontaktAsync(f.KontaktId, fotodaten, connection, ct); }
                }
                catch (Exception ex) { throw new InvalidOperationException($"Fehler bei Foto {f.KontaktId}", ex); } // Parallel.ForEachAsync würde Fehler in einer AggregateException verpacken
                finally { var current = Interlocked.Increment(ref done); }
            });
        transaction.Commit();
    }

    private static async Task<byte[]?> LadeFotoAsync(string url, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Accept.Add(new("image/*"));
        using var response = await HttpService.Client.SendAsync(request, ct);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsByteArrayAsync(ct);
    }

    private static async Task SpeichereFotoFuerKontaktAsync(int kontaktId, byte[] fotodaten, SQLiteConnection connection, CancellationToken ct)
    {
        using var cmd = connection.CreateCommand();
        cmd.CommandText = @"INSERT INTO Fotos (AdressId, Fotodaten) VALUES (@id, @foto) ON CONFLICT(AdressId) DO UPDATE SET Fotodaten = excluded.Fotodaten;";
        cmd.Parameters.AddWithValue("@id", kontaktId);
        cmd.Parameters.AddWithValue("@foto", fotodaten);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private void ComboBox_DrawItem(object sender, DrawItemEventArgs e) // ComboBox disable highlighting
    {
        if (sender is ComboBox comboBox && comboBox.IsHandleCreated)
        {
            Color BgClr;
            Color TxClr;
            if ((e.State & DrawItemState.ComboBoxEdit) == DrawItemState.ComboBoxEdit)  // Do not highlight main display
            {
                BgClr = comboBox.BackColor;
                TxClr = comboBox.ForeColor;
            }
            else
            {
                BgClr = e.BackColor;
                TxClr = e.ForeColor;
            }
            e.Graphics.FillRectangle(new SolidBrush(BgClr), e.Bounds);
            if (e.Index >= 0 && comboBox.Items.Count > e.Index && comboBox.Items[e.Index] != null)
            {
                var itemText = comboBox.Items[e.Index]?.ToString() ?? string.Empty;
                TextRenderer.DrawText(e.Graphics, itemText, e.Font, e.Bounds, TxClr, BgClr, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            }
        }
    }

    private void BirthdaysToolStripMenuItem_Click(object sender, EventArgs e) => BirthdayReminder(true);

    private void BirthdayReminder(bool showIfEmpty = false)
    {
        DataGridView? dgv;
        string idRessource;
        var isLocal = false;
        if (tabControl.SelectedTab == addressTabPage)
        {
            dgv = addressDGV;
            idRessource = "Id";
            isLocal = true;
        }
        else
        {
            dgv = contactDGV;
            idRessource = "Ressource";
        }
        if (dgv.Rows.Count > 0)
        {
            List<(DateTime Datum, string Name, int Alter, int Tage, string Id)> bevorstehendeGeburtstage = [];
            var heute = DateTime.Today;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) { continue; }
                if (!DateTime.TryParse(row.Cells["Geburtstag"].Value?.ToString(), out var geburtsdatum)) { continue; }
                var geburtstagDiesesJahr = new DateTime(heute.Year, geburtsdatum.Month, geburtsdatum.Day);
                var geburtstagNaechstesJahr = geburtstagDiesesJahr.AddYears(1);
                var spanDiesesJahr = geburtstagDiesesJahr - heute;
                var spanNaechstesJahr = geburtstagNaechstesJahr - heute;
                var relevanterSpan = Math.Abs(spanDiesesJahr.TotalDays) < Math.Abs(spanNaechstesJahr.TotalDays) ? spanDiesesJahr : spanNaechstesJahr;
                var relevanteTage = (int)relevanterSpan.TotalDays;
                if (relevanteTage >= -birthdayRemindAfter && relevanteTage <= birthdayRemindLimit)
                {
                    var vorname = row.Cells["Vorname"].Value?.ToString() ?? string.Empty;
                    var nachname = row.Cells["Nachname"].Value?.ToString() ?? string.Empty;
                    var name = vorname + (string.IsNullOrEmpty(vorname) ? "" : " ") + nachname;
                    var alter = heute.Year - geburtsdatum.Year;
                    if (geburtsdatum.Date > heute.AddYears(-alter)) { alter--; }
                    var id = row.Cells[idRessource].Value?.ToString() ?? string.Empty;
                    bevorstehendeGeburtstage.Add((Datum: geburtsdatum, Name: name, Alter: alter, Tage: relevanteTage, Id: id));
                }
            }

            bevorstehendeGeburtstage = [.. bevorstehendeGeburtstage.OrderBy(g => g.Tage)];
            if (bevorstehendeGeburtstage.Count > 0 || showIfEmpty)
            {
                using var frm = new FrmBirthdays(sColorScheme, bevorstehendeGeburtstage, birthdayRemindLimit, birthdayRemindAfter, isLocal);
                frm.BirthdayAutoShow = tabControl.SelectedTab == addressTabPage ? birthdayAddressShow : birthdayContactShow;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    if (frm.SelectionIndex >= 0 && frm.SelectionIndex < bevorstehendeGeburtstage.Count)
                    {
                        var selectedBirthday = bevorstehendeGeburtstage[frm.SelectionIndex];
                        foreach (DataGridViewRow row in dgv.Rows)
                        {
                            if (row.Cells[idRessource].Value?.ToString() == (string?)selectedBirthday.Id)
                            {
                                row.Selected = true;
                                dgv.FirstDisplayedScrollingRowIndex = row.Index;
                                if (tabControl.SelectedTab == addressTabPage) { AddressEditFields(row.Index); }
                                else if (tabControl.SelectedTab == contactTabPage) { ContactEditFields(row.Index); }
                                break;
                            }
                        }
                    }
                }
                birthdayRemindLimit = frm.BirthdayRemindLimit;
                birthdayRemindAfter = frm.BirthdayRemindAfter;
                if (tabControl.SelectedTab == addressTabPage) { birthdayAddressShow = frm.BirthdayAutoShow; }
                else if (tabControl.SelectedTab == contactTabPage) { birthdayContactShow = frm.BirthdayAutoShow; }
            }
        }
    }

    private void ComboBox_Resize(object sender, EventArgs e)
    {
        var comboBox = (ComboBox)sender;
        if (!comboBox.IsHandleCreated) { return; }
        comboBox.BeginInvoke(new Action(() => comboBox.SelectionLength = 0));
        //comboBox.InvokeAsync(() => comboBox.SelectionLength = 0);
    }

    private void AddressDGV_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right) { ColumnSelectToolStripMenuItem_Click(addressDGV, e); }
    }

    private void ContactDGV_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right) { ColumnSelectToolStripMenuItem_Click(contactDGV, e); }
    }

    private void AddressDGV_RowContextMenuStripNeeded(object sender, DataGridViewRowContextMenuStripNeededEventArgs e) => e.ContextMenuStrip = contextDgvMenu;

    private void ContactDGV_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right)
        {
            var hitTestInfo = contactDGV.HitTest(e.X, e.Y);
            if (hitTestInfo.Type == DataGridViewHitTestType.Cell)
            {
                contactDGV.Rows[hitTestInfo.RowIndex].Selected = true;
                contextDgvMenu.Show(contactDGV, new System.Drawing.Point(e.X, e.Y));
            }
        }
    }

    private async void MainDropDown_Opening(object? sender, CancelEventArgs e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0 && contactDGV.SelectedRows[0] != null)
        {
            if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp())
            {
                await CreateContactAsync();
                e.Cancel = true;
            }
            if (CheckContactDataChange())
            {
                ShowMultiPageTaskDialog();
                e.Cancel = true;
            }
        }
    }

    private void RecentToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        recentToolStripMenuItem.DropDownItems.Clear();
        var first = true;
        foreach (var file in recentFiles)
        {
            if (file == databaseFilePath) { continue; }
            var item = new ToolStripMenuItem(file) { Image = Resources.address_book16, ShortcutKeyDisplayString = first ? "F12" : string.Empty };
            first = false;
            item.Click += (s, e) =>
            {
                if (_dataTable != null) { SaveSQLDatabase(true); }
                ConnectSQLDatabase(file);
                ignoreSearchChange = true;
                searchTSTextBox.TextBox.Clear();
                ignoreSearchChange = false;
            };
            recentToolStripMenuItem.DropDownItems.Add(item);
        }
    }
    private void DokuListView_Resize(object sender, EventArgs e)
    {
        var totalWidth = dokuListView.ClientSize.Width;
        var column2Width = 80;
        var column3Width = 120;
        var column1Width = totalWidth - column2Width - column3Width;
        dokuListView.Columns[0].Width = column1Width > 0 ? column1Width : 0;
        dokuListView.Columns[1].Width = column2Width;
        dokuListView.Columns[2].Width = column3Width;
    }

    private void Tabulation_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabulation.SelectedTab == tabPageDetail)
        {
            newTSButton.Visible = copyTSButton.Visible = deleteTSButton.Visible = wordTSButton.Visible = envelopeTSButton.Visible = detailSeparator1.Visible = detailSeparator2.Visible = true;
            dokuPlusTSButton.Visible = dokuMinusTSButton.Visible = dokuShowTSButton.Visible = dokuSeparator1.Visible = dokuSeparator2.Visible = false;
        }
        else if (tabulation.SelectedTab == tabPageDoku)
        {
            newTSButton.Visible = copyTSButton.Visible = deleteTSButton.Visible = wordTSButton.Visible = envelopeTSButton.Visible = detailSeparator1.Visible = detailSeparator2.Visible = false;
            dokuPlusTSButton.Visible = dokuMinusTSButton.Visible = dokuShowTSButton.Visible = dokuSeparator1.Visible = dokuSeparator2.Visible = true;
        }
    }

    private void DokuListView_SelectedIndexChanged(object sender, EventArgs e) => dokuMinusTSButton.Enabled = dokuShowTSButton.Enabled = dokuListView.SelectedItems.Count > 0;

    private void DokuMinusTSButton_Click(object sender, EventArgs e)
    {
        if (dokuListView.SelectedItems.Count == 1)
        {
            var index = dokuListView.SelectedItems[0].Index;
            dokuListView.Items.RemoveAt(index);
            if (dokuListView.Items.Count > 0)
            {
                if (index >= dokuListView.Items.Count) { index = dokuListView.Items.Count - 1; }
                dokuListView.Items[index].Selected = true;
            }
            if (tabControl.SelectedTab == addressTabPage) { ListView2DataTable(); }
            //else if (tabControl.SelectedTab == contactTabPage) { ListView2ChangedList(); }
        }
    }

    private void DokuShowTSButton_Click(object sender, EventArgs e)
    {
        if (dokuListView.SelectedItems.Count == 1)
        {
            var dateipfad = dokuListView.SelectedItems[0].Text;
            Utilities.StartFile(Handle, dateipfad);
        }
    }

    private void DokuPlusTSButton_Click(object sender, EventArgs e)
    {
        openFileDialog.Title = "Datei auswählen";
        var documentFilter = string.Join(";", documentTypes);
        openFileDialog.Filter = $"Dokumente ({documentFilter})|{documentFilter}|Alle Dateien (*.*)|*.*";
        openFileDialog.Multiselect = true;
        openFileDialog.FileName = string.Empty;
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            foreach (var pfad in openFileDialog.FileNames) { Add2dokuListView(new FileInfo(pfad), false); }
            dokuListView.ListViewItemSorter = new ListViewItemComparer();
            dokuListView.Sort();
            ListView2DataTable();
            //else if (tabControl.SelectedTab == contactTabPage) { ListView2ChangedList(); }
        }
    }

    private void ListView2DataTable()
    {
        if (_dataTable == null) { return; }
        List<string> dateipfade = [];
        foreach (ListViewItem item in dokuListView.Items) { dateipfade.Add(item.Text); }
        var json = JsonSerializer.Serialize(dateipfade);
        if (addressDGV.SelectedRows.Count > 0 && addressDGV.SelectedRows[0].DataBoundItem is DataRowView rowView) { rowView.Row["Dokumente"] = json; }
        tabPageDoku.ImageIndex = dokuListView.Items.Count > 0 ? 4 : 3;
    }

    private void StartPictureBox_Click(object sender, EventArgs e)
    {
        if (searchTextBox.Text.Length > 0) { searchTextBox.Clear(); }
        else { ActiveControl = searchTextBox; }

    }

    private void SearchTextBox_Enter(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(searchTextBox.Text)) { allDokuLVItems = [.. dokuListView.Items.Cast<ListViewItem>()]; }
        searchTextBox.BackColor = Color.White;
        searchTextBox.BorderStyle = searchPictureBox.BorderStyle = BorderStyle.FixedSingle;
        NativeMethods.SendMessage(searchTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        NativeMethods.SendMessage(searchTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);

    }

    private void SearchTextBox_Leave(object sender, EventArgs e)
    {
        searchTextBox.BackColor = Color.WhiteSmoke;
        searchTextBox.BorderStyle = searchPictureBox.BorderStyle = BorderStyle.None;
    }

    private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter && dokuListView.SelectedItems.Count > 0)
        {
            e.Handled = e.SuppressKeyPress = true;
            Utilities.StartFile(Handle, dokuListView.SelectedItems[0].Text);
        }
    }

    private void DokuListView_ColumnClick(object sender, ColumnClickEventArgs e)
    {
        if (e.Column == lastColumn) { lastOrder = lastOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending; }
        else
        {
            lastOrder = SortOrder.Ascending;
            lastColumn = e.Column;
        }
        dokuListView.ListViewItemSorter = new ListViewItemComparer(e.Column, lastOrder);
        dokuListView.Sort();
    }

    private void SearchTextBox_TextChanged(object sender, EventArgs e)
    {
        var filter = searchTextBox.Text.Trim();
        dokuListView.BeginUpdate();
        dokuListView.Items.Clear();
        if (string.IsNullOrEmpty(filter)) { dokuListView.Items.AddRange([.. allDokuLVItems]); }
        else
        {
            var gefiltert = allDokuLVItems.Where(item => item.Text.Contains(filter, StringComparison.OrdinalIgnoreCase)).ToArray();
            dokuListView.Items.AddRange(gefiltert);
        }
        dokuListView.EndUpdate();
        if (searchTextBox.Text.Length > 0)
        {
            searchPictureBox.Image = Resources.DeleteFilter16;
            searchPictureBox.Cursor = Cursors.Hand;
            if (dokuListView.Items.Count > 0) { dokuListView.Items[0].Selected = true; }
        }
        else
        {
            searchPictureBox.Image = Resources.Search_16;
            searchPictureBox.Cursor = Cursors.Default;
        }
    }

    private void DokuListView_MouseDoubleClick(object sender, MouseEventArgs e)
    {
        if (e is not MouseEventArgs me) { return; }
        var senderList = (ListView)sender;
        var hit = senderList.HitTest(me.Location);
        if (hit.Item != null && hit.SubItem != null && hit.Item.SubItems.IndexOf(hit.SubItem) == 0) { Utilities.StartFile(Handle, hit.Item.Text); }
    }


    private void FileSystemWatcher_OnChanged(object sender, FileSystemEventArgs e)
    {
        debounceTimer.Stop(); // Stop the timer to prevent multiple triggers
        Debug.WriteLine($"ChangedEvent: {e.ChangeType} - {e.FullPath} - {e.Name}");
        if (e.Name is { Length: > 2 } name && name.StartsWith("~$")) { debounceTimer.Start(); } // vorhandenes Tag bleibt; Workaround für neue Word-Dokumente
        else
        {
            debounceTimer.Tag = e.FullPath;
            if (!string.IsNullOrEmpty(e.FullPath)) { debounceTimer.Start(); }
        }
    }

    private void FileSystemWatcher_OnRenamed(object sender, RenamedEventArgs e)
    {
        debounceTimer.Stop(); // Stop the timer to prevent multiple triggers
        Debug.WriteLine($"RenamedEvent: {e.ChangeType} - {e.FullPath}");
        if (e is not RenamedEventArgs me || me.Name == null) { return; }
        debounceTimer.Tag = e.FullPath;
        if (!string.IsNullOrEmpty(e.FullPath)) { debounceTimer.Start(); }
    }

    private void Add2dokuListView(FileInfo info, bool sortAndSave = true)
    {
        ListViewItem item;
        var extension = info.Extension.ToLower();
        if (info.Exists)
        {
            if (!dokuImages.Images.ContainsKey(extension))
            {
                var icon = Icon.ExtractAssociatedIcon(info.FullName);
                if (icon != null) { dokuImages.Images.Add(extension, icon); }
            }
            item = new ListViewItem(info.FullName);
            item.SubItems.Add(Utilities.FormatBytes(info.Length));
            item.SubItems.Add(info.LastWriteTime.ToString("dd.MM.yyyy HH:mm"));
            item.ImageKey = extension;
        }
        else { item = new ListViewItem([info.FullName, string.Empty, string.Empty]); }
        var vorhandenesItem = dokuListView.Items.Cast<ListViewItem>().FirstOrDefault(item => string.Equals(item.Text, info.FullName, StringComparison.OrdinalIgnoreCase));
        if (vorhandenesItem != null && vorhandenesItem.SubItems[1] != null && vorhandenesItem.SubItems[2] != null)
        {
            vorhandenesItem.SubItems[1].Text = item.SubItems[1].Text;
            vorhandenesItem.SubItems[2].Text = item.SubItems[2].Text;
        }
        else { dokuListView.Items.Add(item); }
        if (sortAndSave)
        {
            dokuListView.ListViewItemSorter = new ListViewItemComparer();
            dokuListView.Sort();
            ListView2DataTable();
        }
    }

    private void DebounceTimer_Tick(object sender, EventArgs e)
    {
        debounceTimer.Stop(); // Stop the timer until the next event    
        var text = debounceTimer.Tag as string ?? string.Empty;
        if (string.IsNullOrEmpty(text)) { return; } //  || !File.Exists(text)
        NativeMethods.SetForegroundWindow(Handle);
        var ort = cbOrt.Text;
        var nameEtc = string.Join(" ", new[] { tbVorname.Text, tbNachname.Text, tbFirma.Text }.Where(s => !string.IsNullOrWhiteSpace(s)));
        var inOrt = string.IsNullOrWhiteSpace(ort) ? "" : $" in {ort}";
        TaskDialogButton linkButton = new TaskDialogCommandLinkButton("Mit Adresse verknüpfen", $"{nameEtc}{inOrt}");
        TaskDialogButton nextButton = new TaskDialogCommandLinkButton("Eine andere Adresse wählen…", "… und neuen Dialog bestätigen");
        TaskDialogButton copyButton = new TaskDialogCommandLinkButton("In Zwischenablage kopieren", "Briefe lassen sich auch manuell hinzügen.");
        var page = new TaskDialogPage
        {
            Caption = appName,
            Heading = "Änderung im Briefordner erkannt",
            Text = $"Datei: {text}",
            Icon = TaskDialogIcon.ShieldWarningYellowBar,
            Buttons = { linkButton, nextButton, copyButton, TaskDialogButton.Cancel },
            AllowCancel = true,
            SizeToContent = true
        };
        var result = TaskDialog.ShowDialog(Handle, page);
        if (result == linkButton)
        {
            if (tabControl.SelectedTab == addressTabPage)
            {
                Add2dokuListView(new FileInfo(text));
                tabulation.SelectedTab = tabPageDoku;
                BringToFront();
            }
        }
        else if (result == nextButton)
        {
            BringToFront();
            ActiveControl = searchTextBox;
            using TaskDialogIcon questionDialogIcon = new(Resources.question32);
            var next = new TaskDialogPage
            {
                Caption = appName,
                Heading = "Möchten Sie die Datei verknüpfen?",
                Text = $"{text}",
                Icon = questionDialogIcon,
                Footnote = $"Wählen Sie die passende Adresse, bevor Sie auf 'Ja' klicken.",
                Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
                AllowCancel = true,
                SizeToContent = true
            };
            if (TaskDialog.ShowDialog(next) == TaskDialogButton.Yes)
            {
                if (tabControl.SelectedTab == addressTabPage)
                {
                    Add2dokuListView(new FileInfo(text));
                    tabulation.SelectedTab = tabPageDoku;
                }
                else if (tabControl.SelectedTab == contactTabPage)
                {
                    Utilities.ErrorMsgTaskDlg(Handle, "Funktion nicht verfügbar", "Google-Kontakte haben beschränkte Feldgrößen", TaskDialogIcon.Information);
                }
            }
        }

        else if (result == copyButton)
        {
            try { Clipboard.SetText(text); }
            catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message, TaskDialogIcon.ShieldErrorRedBar); }
        }
    }

    private void DokuListView_MouseMove(object sender, MouseEventArgs e)
    {
        var info = dokuListView.HitTest(e.Location);
        if (info.Item != null)
        {
            var text = info.Item.Text;
            if (TextRenderer.MeasureText(text, dokuListView.Font).Width > dokuListView.Columns[0].Width)
            {
                if (text != lastTooltipText)
                {
                    lastTooltipText = text;
                    toolTip.SetToolTip(dokuListView, string.Empty);
                    toolTip.Show(text, dokuListView, e.Location.X + 15, e.Location.Y + 15, 2000);
                }
                return;
            }
        }
        lastTooltipText = string.Empty;
        toolTip.SetToolTip(dokuListView, string.Empty);
    }

    private void CheckSaveButton()
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            changedAddressData.Clear();
            if (addressDGV.Rows[prevSelectedAddressRowIndex].DataBoundItem is DataRowView currentRowView)
            {
                var row = currentRowView.Row;
                if (row.RowState == DataRowState.Modified)
                {
                    foreach (DataColumn column in row.Table.Columns)
                    {
                        if (column.ColumnName == "Id") { continue; }
                        var currentValue = row[column.ColumnName, DataRowVersion.Current];
                        var originalValue = row[column.ColumnName, DataRowVersion.Original];
                        if (!Utilities.ValuesEqual(originalValue, currentValue)) { changedAddressData[column.ColumnName] = currentValue?.ToString() ?? string.Empty; }
                    }
                }
                if (changedAddressData.Count > 0) { saveTSButton.Enabled = true; }
                else
                {
                    if (row.RowState == DataRowState.Modified) { row.RejectChanges(); } // Sicherer: Alle ausstehenden Änderungen verwerfen.
                    else { row.AcceptChanges(); }
                    saveTSButton.Enabled = _dataTable?.GetChanges() != null;
                }
            }
            //changedAddressData.Clear();
            //foreach (var cell in addressDGV.Rows[prevSelectedAddressRowIndex].Cells.Cast<DataGridViewCell>().SkipLast(1).Where(cell => !Utilities.ValuesEqual(originalAddressData[cell.OwningColumn.Name], cell.Value)))
            //{
            //    changedAddressData[cell.OwningColumn.Name] = cell.Value?.ToString() ?? string.Empty;
            //    //MessageBox.Show($"Geändert: {cell.OwningColumn.Name} = {cell.Value} ({cell.Value.GetType().ToString()})", "Debug");
            //}
            //if (changedAddressData.Count > 0) { saveTSButton.Enabled = true; }
            //else if (addressDGV.Rows[prevSelectedAddressRowIndex].DataBoundItem is DataRowView dataBoundItem)
            //{
            //    dataBoundItem.Row.AcceptChanges();  // keine Änderung, RowState auf Unchanged setzen 
            //    saveTSButton.Enabled = _dataTable?.GetChanges() != null; // andere Rows könnten geändert sein
            //}
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            saveTSButton.Enabled = CheckContactDataChange() || contactNewRowIndex >= 0;
        }
    }

    private void Clear_SearchTextBox()
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            if (addressDGV.SelectedRows.Count == 0) { searchTSTextBox.TextBox.Clear(); return; }
            var rowIndex = addressDGV.SelectedRows[0].Index;
            searchTSTextBox.TextBox.Clear(); // löst SearchTSTextBox_TextChanged aus    
            if (rowIndex >= 0)
            {
                addressDGV.Rows[rowIndex].Selected = true;
                addressDGV.FirstDisplayedScrollingRowIndex = rowIndex;
            }
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            if (contactDGV.SelectedRows.Count == 0) { searchTSTextBox.TextBox.Clear(); return; }
            var rowIndex = contactDGV.SelectedRows[0].Index;
            searchTSTextBox.TextBox.Clear();
            if (rowIndex >= 0)
            {
                contactDGV.Rows[rowIndex].Selected = true;
                contactDGV.FirstDisplayedScrollingRowIndex = rowIndex;
            }
        }
        searchTSTextBox.Focus();
    }

    public static bool? GetGender(string name) // true für weiblich, false für männlich, null wenn unbekannt
    {
        if (string.IsNullOrWhiteSpace(name)) { return null; }
        return nameGenderMap.TryGetValue(name.Trim(), out var isFemale) ? isFemale : null;
    }

    private void WeiblicheVornamenToolStripMenuItem_Click(object sender, EventArgs e) => Utilities.StartFile(Handle, girlPath);

    private void MännlicheVornamenToolStripMenuItem_Click(object sender, EventArgs e) => Utilities.StartFile(Handle, boysPath);

    private void WebsiteToolStripMenuItem_Click(object sender, EventArgs e) => Utilities.StartLink(Handle, @"https://www.netradio.info/address");

    private void GithubToolStripMenuItem_Click(object sender, EventArgs e) => Utilities.StartLink(Handle, @"https://github.com/ophthalmos/Adressen");

    private void HelpdokuTSMenuItem_Click(object sender, EventArgs e) => Utilities.StartFile(Handle, "AdressenKontakte.pdf");

    private void SortNamesToolStripMenuItem_Click(object sender, EventArgs e)
    {
        SortNameFiles(girlPath);
        SortNameFiles(boysPath);
    }

    private void SortNameFiles(string path)
    {
        if (!File.Exists(path)) { Utilities.ErrorMsgTaskDlg(Handle, "Datei existiert nicht", boysPath); }
        else
        {
            try
            {
                var lines = File.ReadAllLines(boysPath);
                var duplicates = lines.GroupBy(z => z).Count(g => g.Count() > 1);
                File.Copy(boysPath, Path.ChangeExtension(boysPath, ".bak"), true); // Backup erstellen, falls Datei existiert   
                var cleanedLines = lines.Select(static line => line.Trim()).Where(static line => !string.IsNullOrWhiteSpace(line)).Distinct()
                    .OrderBy(static line => line, StringComparer.OrdinalIgnoreCase).ToList();
                File.WriteAllLines(boysPath, cleanedLines);
                var message = duplicates > 0 ? $"{duplicates} doppelte Zeilen wurden entfernt." : "Die Vornamen wurden alphabetisch sortiert";
                Utilities.ErrorMsgTaskDlg(Handle, message, path, TaskDialogIcon.Information);
            }
            catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "Fehler beim Sortieren der Vornamen", ex.Message); }
        }
    }

    private static List<string> ReadNamesFromFile(string filePath)
    {
        if (!File.Exists(filePath)) { return []; } // Leere Liste zurückgeben
        return [.. File.ReadAllLines(filePath).Select(line => line.Trim()).Where(line => !string.IsNullOrEmpty(line))];
    }

    private void NameduplicatesToolStripMenuItem_Click(object sender, EventArgs e)
    {
        var girlNamesList = ReadNamesFromFile(girlPath);
        var boyNamesList = ReadNamesFromFile(boysPath);
        if (girlNamesList != null && boyNamesList != null)
        {
            var girlNamesSet = new HashSet<string>(girlNamesList, StringComparer.OrdinalIgnoreCase);
            var namesInBothFiles = boyNamesList.Where(girlNamesSet.Contains).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (namesInBothFiles != null && namesInBothFiles.Count > 0)
            {
                Utilities.ErrorMsgTaskDlg(Handle, "Namen in beiden Dateien gefunden", string.Join(Environment.NewLine, namesInBothFiles), TaskDialogIcon.Information);
            }
            else
            {
                Utilities.ErrorMsgTaskDlg(Handle, "Keine Duplikate", "Es wurden keine Namen gefunden, die in beiden Dateien vorkommen.", TaskDialogIcon.Information);
            }
        }
        else { Utilities.ErrorMsgTaskDlg(Handle, "Dateien nicht gefunden", girlPath + Environment.NewLine + boysPath); }
    }

    private void TermsofuseToolStripMenuItem_Click(object sender, EventArgs e) => Utilities.StartLink(Handle, "https://www.netradio.info/adressen-terms-of-use/");
    private void PrivacypolicyToolStripMenuItem_Click(object sender, EventArgs e) => Utilities.StartLink(Handle, "https://www.netradio.info/adressen-privacy-policy/");
    private void LicenseTxtToolStripMenuItem_Click(object sender, EventArgs e) => Utilities.StartFile(Handle, "Lizenzvereinbarung.txt");

    private void AdressenMitBriefToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            static bool filter(DataGridViewRow row) => row.Cells["Dokumente"].Value is string json && !string.IsNullOrWhiteSpace(json);
            //&& (JsonSerializer.Deserialize<List<string>>(json) ?? []).Any(pfad =>
            //!pictureBoxExtensions.Contains(Path.GetExtension(pfad), StringComparer.OrdinalIgnoreCase));
            FilterAddressDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
            flexiTSStatusLabel.Text = "… mit Briefverweis";
        }
    }

    private void PhotoPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            var idsMitFoto = LadeAlleAdressIdsMitFoto();
            bool filter(DataGridViewRow row) { return row.Cells["Id"].Value is not null && int.TryParse(row.Cells["Id"].Value!.ToString(), out var rowId) && idsMitFoto.Contains(rowId); }
            FilterAddressDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
            flexiTSStatusLabel.Text = "… mit Bild";
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            static bool filter(DataGridViewRow row) => row.Cells["PhotoURL"].Value is string url && !string.IsNullOrWhiteSpace(url);
            FilterContactDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
            flexiTSStatusLabel.Text = "… mit Bild";
        }
    }

    private void PhotoMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            var idsMitFoto = LadeAlleAdressIdsMitFoto();
            bool filter(DataGridViewRow row) { return row.Cells["Id"].Value is not null && int.TryParse(row.Cells["Id"].Value!.ToString(), out var rowId) && !idsMitFoto.Contains(rowId); }
            FilterAddressDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
            flexiTSStatusLabel.Text = "… ohne Bild";
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            static bool filter(DataGridViewRow row) => row.Cells["PhotoURL"].Value is not string url || string.IsNullOrWhiteSpace(url);
            FilterContactDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
            flexiTSStatusLabel.Text = "… ohne Bild";
        }
    }

    private void MailPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => !string.IsNullOrEmpty((string?)(row.Cells["Mail1"].Value?.ToString() + row.Cells["Mail2"].Value?.ToString()));
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
        }
        else if (tabControl.SelectedTab == contactTabPage) { FilterContactDGV(filter); }
        flexiTSStatusLabel.Text = "… mit E-Mailadresse";
    }

    private void MailMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => string.IsNullOrEmpty((string?)(row.Cells["Mail1"].Value?.ToString() + row.Cells["Mail2"].Value?.ToString()));
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
        }
        else if (tabControl.SelectedTab == contactTabPage) { FilterContactDGV(filter); }
        flexiTSStatusLabel.Text = "… ohne E-Mailadresse";
    }

    private void TelephonePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => !string.IsNullOrEmpty((string?)(row.Cells["Telefon1"].Value?.ToString() + row.Cells["Telefon2"].Value?.ToString()));
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
        }
        else if (tabControl.SelectedTab == contactTabPage) { FilterContactDGV(filter); }
        flexiTSStatusLabel.Text = "… mit Telefonnummer";
    }

    private void TelephoneMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => string.IsNullOrEmpty((string?)(row.Cells["Telefon1"].Value?.ToString() + row.Cells["Telefon2"].Value?.ToString()));
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
        }
        else if (tabControl.SelectedTab == contactTabPage) { FilterContactDGV(filter); }
        flexiTSStatusLabel.Text = "… ohne Telefonnummer";
    }

    private void MobilePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => !string.IsNullOrEmpty(row.Cells["Mobil"].Value?.ToString());
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
        }
        else if (tabControl.SelectedTab == contactTabPage) { FilterContactDGV(filter); }
        flexiTSStatusLabel.Text = "… mit Mobilfunknummer";
    }

    private void MobileMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => string.IsNullOrEmpty(row.Cells["Mobil"].Value?.ToString());
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            filterRemoveToolStripMenuItem.Visible = true;
        }
        else if (tabControl.SelectedTab == contactTabPage) { FilterContactDGV(filter); }
        flexiTSStatusLabel.Text = "… ohne Mobilfunknummer";
    }

    private void FilterlToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        adressenMitBriefToolStripMenuItem.Enabled = tabControl.SelectedTab == addressTabPage && addressDGV.Rows.Count > 0;
        photoPlusFilterToolStripMenuItem.Enabled = photoMinusFilterToolStripMenuItem.Enabled =
            mailPlusFilterToolStripMenuItem.Enabled = mailMinusFilterToolStripMenuItem.Enabled =
            telephonePlusFilterToolStripMenuItem.Enabled = telephoneMinusFilterToolStripMenuItem.Enabled =
            mobilePlusFilterToolStripMenuItem.Enabled = mobileMinusFilterToolStripMenuItem.Enabled =
            (tabControl.SelectedTab == addressTabPage && addressDGV.Rows.Count > 0) || (tabControl.SelectedTab == contactTabPage && contactDGV.Rows.Count > 0);
    }

    private void FilterAddressDGV(Func<DataGridViewRow, bool> filterCondition)
    {
        if (addressDGV.DataSource == null || BindingContext == null) { return; }
        var manager = (CurrencyManager)BindingContext[addressDGV.DataSource]; // Holen des CurrencyManagers, um 
        manager.SuspendBinding(); // Datenbindung während der Bearbeitung pausieren
        addressDGV.SuspendLayout(); // UI-Updates anhalten für bessere Performance
        addressDGV.ClearSelection();
        AddressEditFields(-1); // Ihre Methode zum Zurücksetzen der Felder

        //foreach (DataGridViewRow row in addressDGV.Rows) // Alle Zeilen zunächst sichtbar machen    
        //{
        //    if (row.IsNewRow) { continue; }
        //    else { row.Visible = true; }
        //}

        var visibleRowCount = 0;
        foreach (DataGridViewRow row in addressDGV.Rows)
        {
            if (row.IsNewRow) { continue; }
            if (filterCondition(row)) // Hier wird die übergebene Filter-Logik aufgerufen!
            {
                row.Visible = true;
                visibleRowCount++;
            }
            else { row.Visible = false; }
        }
        if (visibleRowCount > 0)
        {
            var firstVisibleIndex = Utilities.GetFirstVisibleRowIndex(addressDGV);  // Die erste sichtbare Zeile finden und selektieren
            if (firstVisibleIndex != -1)
            {
                addressDGV.Rows[firstVisibleIndex].Selected = true;
                //AddressEditFields(firstVisibleIndex); // wird durch AddressDGV_SelectionChanged aufgerufen
            }
        }
        manager.ResumeBinding(); // Datenbindung wieder aufnehmen
        addressDGV.ResumeLayout(); // UI-Updates wieder erlauben
        var rowCount = addressDGV.Rows.Count;
        toolStripStatusLabel.Text = rowCount == visibleRowCount ? $"{visibleRowCount} Adressen" : $"{visibleRowCount}/{rowCount} Adressen";
    }

    private void FilterContactDGV(Func<DataGridViewRow, bool> filterCondition)
    {
        contactDGV.SuspendLayout(); // UI-Updates anhalten für bessere Performance
        contactDGV.ClearSelection();
        ContactEditFields(-1); // Ihre Methode zum Zurücksetzen der Felder

        //foreach (DataGridViewRow row in contactDGV.Rows) // Alle Zeilen zunächst sichtbar machen    
        //{
        //    if (row.IsNewRow) { continue; }
        //    else { row.Visible = true; }
        //}

        var visibleRowCount = 0;
        foreach (DataGridViewRow row in contactDGV.Rows)
        {
            if (row.IsNewRow) { continue; }
            if (filterCondition(row)) // Hier wird die übergebene Filter-Logik aufgerufen!
            {
                row.Visible = true;
                visibleRowCount++;
            }
            else { row.Visible = false; }
        }
        if (visibleRowCount > 0)
        {
            var firstVisibleIndex = Utilities.GetFirstVisibleRowIndex(contactDGV);  // Die erste sichtbare Zeile finden und selektieren
            if (firstVisibleIndex != -1)
            {
                contactDGV.Rows[firstVisibleIndex].Selected = true;
                ContactEditFields(firstVisibleIndex); // Ihre Methode zum Füllen der Felder
            }
        }
        contactDGV.ResumeLayout(); // UI-Updates wieder erlauben    
        var rowCount = contactDGV.Rows.Count;
        toolStripStatusLabel.Text = rowCount == visibleRowCount ? $"{visibleRowCount} Kontakte" : $"{visibleRowCount}/{rowCount} Kontakte";

    }

    private void TopAlignZoomPictureBox_DoubleClick(object sender, EventArgs e)
    {
        if ((tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count == 0) || (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count == 0)) { return; }
        openFileDialog.Title = "Foto auswählen";
        openFileDialog.Filter = $"Bilddateien|{string.Join(";", pictureBoxExtensions.Select(ext => "*" + ext))}|Alle Dateien|*.*";
        openFileDialog.Multiselect = false;
        openFileDialog.FileName = string.Empty;
        openFileDialog.CheckFileExists = true;
        if (openFileDialog.ShowDialog(this) == DialogResult.OK)
        {
            if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0 && addressDGV.SelectedRows[0].DataBoundItem is DataRowView rowView)
            {
                var bildDaten = File.ReadAllBytes(openFileDialog.FileName);
                if (bildDaten.Length == 0)
                {
                    Utilities.ErrorMsgTaskDlg(Handle, "Fehler beim Laden der Bilddatei", "Die ausgewählte Datei ist leer.", TaskDialogIcon.ShieldErrorRedBar);
                    return;
                }

                Image? loadedImage = null; // Das von Image.FromStream(ms) erstellte Objekt
                Image? scaledImage = null; // Image-Objekte außerhalb deklarieren, um im catch-Block Zugriff zu haben
                try
                {
                    if (topAlignZoomPictureBox.Image is not null)
                    {
                        topAlignZoomPictureBox.Image.Dispose();
                        topAlignZoomPictureBox.Image = null;
                    }
                    using var ms = new MemoryStream(bildDaten);
                    loadedImage = Image.FromStream(ms);
                    var originalFormat = loadedImage.RawFormat;
                    Utilities.WendeExifOrientierungAn(loadedImage);
                    Image finalImage;
                    if (loadedImage.Width > 100)
                    {
                        scaledImage = Utilities.SkaliereBildDaten(loadedImage, 100);
                        finalImage = scaledImage;
                    }
                    else { finalImage = loadedImage; }
                    topAlignZoomPictureBox.Image = finalImage!;
                    delPictboxToolStripButton.Enabled = true;
                    byte[] datenZumSpeichern;
                    using (var outputMs = new MemoryStream())
                    {
                        var saveFormat = originalFormat.Equals(ImageFormat.Png) ? ImageFormat.Png : ImageFormat.Jpeg;
                        finalImage!.Save(outputMs, saveFormat);
                        datenZumSpeichern = outputMs.ToArray();
                    }
                    SpeichereFotoFuerKontakt(Convert.ToInt32(rowView.Row["Id"]), datenZumSpeichern, databaseFilePath);
                    loadedImage = null;  // ** Erfolgreiche Übergabe: Referenzen auf null setzen **
                    scaledImage = null;
                }
                catch (Exception ex)
                {
                    loadedImage?.Dispose();  // Freigabe, falls vor der Übergabe an die PictureBox ein Fehler auftrat
                    scaledImage?.Dispose();
                    Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message);
                }
            }
            else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
            {
                Image? workingImage = null;           // Das Bild, mit dem wir arbeiten (skaliert oder Klon)
                Image? finalImageToUpload = null;     // Das Bild, das final hochgeladen wird
                Image? finalImageForDisplay = null;   // Das Bild für die PictureBox (ggf. "wie Google")
                var origImgFormat = ImageFormat.Jpeg; // Standard für UpdateContactPhotoAsync
                try
                {
                    using (var fs = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    using (var originalImage = Image.FromStream(fs)) // originalImage nur in diesem Block gültig
                    {
                        origImgFormat = originalImage.RawFormat;
                        Utilities.WendeExifOrientierungAn(originalImage);
                        if (fs.Length > 1024 * 1024)
                        {
                            Utilities.ErrorMsgTaskDlg(Handle, "Automatische Größenreduzierung", $"Die Dateigröße ist größer als 1 MB ({Utilities.FormatBytes(fs.Length)}).\nEs erfolgt eine Skalierung auf 250 Pixel Breite.", TaskDialogIcon.ShieldWarningYellowBar);
                            workingImage = Utilities.SkaliereBildDaten(originalImage, 250);
                        }
                        else { workingImage = (Image)originalImage.Clone(); }
                    }
                    var ressource = contactDGV.Rows[contactDGV.SelectedRows[0].Index].Cells["Ressource"]?.Value?.ToString() ?? string.Empty;
                    var initialButtonYes = new TaskDialogButton("Hochladen") { AllowCloseDialog = false };
                    var initialButtonNo = TaskDialogButton.Cancel;
                    var caveText = string.Empty;
                    var radioButtons = new List<TaskDialogRadioButton>();
                    TaskDialogRadioButton? centerRadio = null;
                    TaskDialogRadioButton? topRadio = null;
                    TaskDialogRadioButton? downRadio = null;
                    TaskDialogRadioButton? skipRadio = null;
                    if (workingImage.Height > workingImage.Width && workingImage.Width > topAlignZoomPictureBox.Width)
                    {
                        topRadio = new TaskDialogRadioButton("&Oben priorisieren, nur unten abschneiden");
                        centerRadio = new TaskDialogRadioButton("&Mitte priorisieren, oben/unten abschneiden") { Checked = true };
                        downRadio = new TaskDialogRadioButton("&Unten priorisieren, nur oben abschneiden");
                        skipRadio = new TaskDialogRadioButton("&Nicht beschneiden (nicht empfohlen)");
                        radioButtons.AddRange([topRadio, centerRadio, downRadio, skipRadio]);
                        caveText =
                            "\n\nDas Bild ist höher als breit. Es wird beim Download\n" +
                            "gemäß den Google-Vorgaben in einer auf 100 Pixel" + Environment.NewLine +
                            "Höhe skalierten Version ausgegeben. Dies führt da-\n" +
                            "zu, dass das Foto den horizontal verfügbaren Platz\n" +
                            "nicht vollständig ausfüllen wird.\n\n" +
                            "Sie können das hochzuladende Bild mit einer der\n" +
                            "folgenden Optionen zum Quadrat beschneiden:";
                    }
                    var initialPage = new TaskDialogPage()
                    {
                        Caption = "Google Kontakte",
                        Heading = "Möchten Sie die Änderung speichern?",
                        Text = $"Falls ein Foto vorhanden ist, wird es überschrieben.\n\nUpload-Information: Abmessung {workingImage.Width}×{workingImage.Height} Pixel.{caveText}",
                        Icon = TaskDialogIcon.ShieldWarningYellowBar,   // new(Resources.question32),
                        AllowCancel = true,
                        SizeToContent = true,
                        Buttons = { initialButtonNo, initialButtonYes }
                    };
                    foreach (var rb in radioButtons) { initialPage.RadioButtons.Add(rb); }
                    var inProgressCloseButton = TaskDialogButton.Close;
                    inProgressCloseButton.Enabled = false;
                    var progressPage = new TaskDialogPage()
                    {
                        Caption = appCont,
                        Heading = "Bitte warten…",
                        Text = "Das Foto wird hochgeladen.",
                        Icon = TaskDialogIcon.Information,
                        ProgressBar = new TaskDialogProgressBar() { State = TaskDialogProgressBarState.Marquee },
                        Buttons = { inProgressCloseButton }
                    };
                    initialButtonYes.Click += (sender, e) =>
                    {
                        Image? intermediateImageToDispose = null;
                        if (topRadio?.Checked == true)
                        {
                            intermediateImageToDispose = workingImage;
                            workingImage = Utilities.BeschneideZuQuadrat(workingImage, null);
                            finalImageToUpload = workingImage;
                            finalImageForDisplay = (Image)workingImage.Clone();
                        }
                        else if (centerRadio?.Checked == true)
                        {
                            intermediateImageToDispose = workingImage; // Das alte 'workingImage' zum Dispose vormerken
                            workingImage = Utilities.BeschneideZuQuadrat(workingImage, false); // 'workingImage' ist jetzt das *neue* beschnittene
                            finalImageToUpload = workingImage; // Hochladen
                            finalImageForDisplay = (Image)workingImage.Clone(); // Anzeigen
                        }
                        else if (downRadio?.Checked == true)
                        {
                            intermediateImageToDispose = workingImage;
                            workingImage = Utilities.BeschneideZuQuadrat(workingImage, true);
                            finalImageToUpload = workingImage;
                            finalImageForDisplay = (Image)workingImage.Clone();
                        }
                        else if (skipRadio?.Checked == true)
                        {
                            finalImageToUpload = workingImage; // 'workingImage' wird *nicht* ersetzt
                            finalImageForDisplay = Utilities.ReduziereWieGoogle(workingImage, 100);
                        }
                        else  // Fall: Keine RadioButtons (Bild war nicht hochkant)
                        {
                            finalImageToUpload = workingImage;
                            finalImageForDisplay = (Image)workingImage.Clone();
                        }

                        topAlignZoomPictureBox.Image = finalImageForDisplay;
                        initialPage.Navigate(progressPage);
                        intermediateImageToDispose?.Dispose(); // Das zwischenzeitliche Bild (skaliert oder Klon) entsorgen, *falls* es ersetzt wurde
                    };
                    progressPage.Created += async (s, e) =>
                    {
                        try { await UpdateContactPhotoAsync(ressource, finalImageToUpload!, origImgFormat, () => progressPage.Buttons.First().PerformClick()); }
                        finally { workingImage?.Dispose(); }  // finalImageForDisplay wird von PictureBox verwaltet, darf hier nicht disposed werden    
                    };
                    TaskDialog.ShowDialog(Handle, initialPage);
                    delPictboxToolStripButton.Enabled = true;
                }
                catch (Exception ex)
                {
                    Utilities.ErrorMsgTaskDlg(Handle, $"Fehler beim Laden: {ex.GetType().ToString()}", $"Bild konnte nicht geladen werden: {ex.Message}", TaskDialogIcon.Error);
                    workingImage?.Dispose();
                    finalImageForDisplay?.Dispose();
                }
            }
        }
    }

    private static void SpeichereFotoFuerKontakt(int kontaktId, byte[] fotodaten, string dbPath)
    {
        using var liteConnection = new SQLiteConnection($"Data Source={dbPath};FailIfMissing=True");
        liteConnection.Open();
        var checkQuery = "SELECT COUNT(*) FROM Fotos WHERE AdressId = @id"; // Prüfen, ob für diesen Kontakt bereits ein Foto existiert
        long count;
        using (var checkCmd = new SQLiteCommand(checkQuery, liteConnection))
        {
            checkCmd.Parameters.AddWithValue("@id", kontaktId);
            count = (long)checkCmd.ExecuteScalar();
        }
        string query;
        if (count > 0) { query = "UPDATE Fotos SET Fotodaten = @foto WHERE AdressId = @id"; } // Existiert bereits -> UPDATE
        else { query = "INSERT INTO Fotos (AdressId, Fotodaten) VALUES (@id, @foto)"; } // Existiert nicht -> INSERT
        using var cmd = new SQLiteCommand(query, liteConnection);
        cmd.Parameters.AddWithValue("@id", kontaktId);
        cmd.Parameters.AddWithValue("@foto", fotodaten);
        cmd.ExecuteNonQuery();
    }

    private void EntferneFotoFuerKontakt(int kontaktId)
    {
        using var liteConnection = new SQLiteConnection($"Data Source={databaseFilePath};FailIfMissing=True");
        liteConnection.Open();
        using var cmd = new SQLiteCommand("DELETE FROM Fotos WHERE AdressId = @id", liteConnection);
        cmd.Parameters.AddWithValue("@id", kontaktId); // Parameter um SQL-Injection zu vermeiden.
        cmd.ExecuteNonQuery(); // gibt die Anzahl der betroffenen Zeilen zurück
    }

    private HashSet<int> LadeAlleAdressIdsMitFoto()
    {
        var idsMitFoto = new HashSet<int>();
        using (var liteConnection = new SQLiteConnection($"Data Source={databaseFilePath}"))
        {
            liteConnection.Open();
            using var cmd = new SQLiteCommand("SELECT AdressId FROM Fotos", liteConnection);
            using var reader = cmd.ExecuteReader();
            while (reader.Read()) { _ = idsMitFoto.Add(reader.GetInt32(0)); } // GetInt32(0) liest nur die erste Spalte der aktuellen Zeile als int.
        }
        return idsMitFoto;
    }

    private void FilterRemoveToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => true; // lokale Funktion statt Lambda-Ausdruck (Func<>)
        if (tabControl.SelectedTab == addressTabPage)
        {
            var rowIndex = addressDGV.SelectedRows.Count > 0 ? addressDGV.SelectedRows[0].Index : -1;
            FilterAddressDGV(filter);
            if (rowIndex >= 0 && addressDGV.Rows[rowIndex] != null)
            {
                addressDGV.Rows[rowIndex].Selected = true;
                addressDGV.FirstDisplayedScrollingRowIndex = rowIndex;
            }
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            if (CheckContactDataChange()) { ShowMultiPageTaskDialog(); }
            var rowIndex = contactDGV.SelectedRows.Count > 0 ? contactDGV.SelectedRows[0].Index : -1;
            FilterContactDGV(filter);
            if (rowIndex >= 0 && contactDGV.Rows[rowIndex] != null)
            {
                contactDGV.Rows[rowIndex].Selected = true;
                contactDGV.FirstDisplayedScrollingRowIndex = rowIndex;
            }
        }
        filterRemoveToolStripMenuItem.Visible = false;
        ignoreSearchChange = true; // F9 löst SearchTSTextBox_TextChanged aus
        searchTSTextBox.TextBox.Clear();
        ignoreSearchChange = false;
        flexiTSStatusLabel.Text = "";
    }

    private void AddPictboxToolStripButton_Click(object sender, EventArgs e)
    {
        TopAlignZoomPictureBox_DoubleClick(addPictboxToolStripButton, EventArgs.Empty);
    }

    private async void DelPictboxToolStripButton_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0 && addressDGV.SelectedRows[0].Index is int rowIndex &&
            addressDGV.Rows[rowIndex].DataBoundItem is DataRowView rowView && Utilities.YesNo_TaskDialog(Handle, "Adressen", "Möchten Sie das Bild wirklich löschen?",
            "Es wird unwiderruflich aus der Datenbank entfernt.", new(Resources.question32), false, "&Löschen", "&Belassen"))
        {
            EntferneFotoFuerKontakt(Convert.ToInt32(rowView.Row["Id"]));
            delPictboxToolStripButton.Enabled = false;
            topAlignZoomPictureBox.Image = Resources.AddressBild100;
        }
        else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0 && Utilities.YesNo_TaskDialog(Handle, "Kontakte", "Möchten Sie das Bild  löschen?",
            "Das Bild wird unwiderruflich gelöscht.", new(Resources.question32), false, "&Löschen", "&Belassen"))
        {
            await DeleteContactPhotoAsync(contactDGV.Rows[contactDGV.SelectedRows[0].Index].Cells["Ressource"]?.Value?.ToString() ?? string.Empty);
        }
    }

    private async void Move2OtherDGVToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0 && Utilities.YesNo_TaskDialog(Handle, "Google Kontakte", "Möchten Sie den Kontakt wirklich löschen?",
            "Verschieben löscht den Google-Kontakt unwiderruflich!", new(Resources.question32), false, "&Verschieben", "&Abbrechen"))
        {
            CopyToOtherDGVMenuItem_Click(move2OtherDGVToolStripMenuItem, EventArgs.Empty);
            var row = contactDGV.SelectedRows[0];
            if (row != null) { await DeleteGoogleContact(row.Index); }
            Application.DoEvents();
            AddressEditFields(addressDGV.Rows[^1].Index); // Letzte Zeile auswählen
        }
    }

    private void UpdateMembershipTags()
    {
        var groupsList = tabControl.SelectedTab == contactTabPage ? curContactMemberships : curAddressMemberships;
        flowLayoutPanel.Controls.Clear();
        foreach (var membership in groupsList)
        {
            var tagControl = new TagControl
            {
                Text = membership, // keine automatische Zuweisung einer Eigenschaft zu einer anderen
                Membership = membership // Text = Membership = membership geht also nicht!
            };
            tagControl.DeleteClick += (sender, e) =>
            {
                var ctrl = sender as TagControl;
                var membershipToRemove = ctrl?.Membership; // Daten aus dem Control holen
                if (!string.IsNullOrEmpty(membershipToRemove))
                {
                    if (tabControl.SelectedTab == contactTabPage) { curContactMemberships.Remove(membershipToRemove); }
                    else { curAddressMemberships.Remove(membershipToRemove); }
                    UpdateMembershipTags();
                    UpdateMembershipJson();
                    if (tabControl.SelectedTab == addressTabPage) { PopulateMemberships(); }
                    UpdateMembershipCBox();
                    UpdatePlaceholderVis();
                    CheckSaveButton();
                }
            };
            flowLayoutPanel.Controls.Add(tagControl);
        }
    }

    private void TagPanel_MouseDeactivation(object? sender, EventArgs e)
    {
        var currentControl = sender as Control;
        var currentPanel = (currentControl as Panel) ?? (currentControl?.Parent as Panel);
        if (currentPanel == null) { return; }
        var clientPoint = currentPanel.PointToClient(Cursor.Position);
        if (currentPanel.ClientRectangle.Contains(clientPoint)) { return; }
        var currentButton = currentPanel.Controls.OfType<Button>().FirstOrDefault();
        if (currentButton != null) { currentButton.Enabled = false; }
    }

    private void TagButton_Click(object sender, EventArgs e)
    {
        var newMembership = tagComboBox.Text.Trim();
        if (string.IsNullOrEmpty(newMembership)) { return; }
        if (newMembership == "*") { newMembership = "★"; }
        if (tabControl.SelectedTab == contactTabPage)
        {
            if (curContactMemberships.Contains(newMembership))
            {
                tagComboBox.SelectAll();
                tagComboBox.Focus();
                return;
            }
            curContactMemberships.Add(newMembership);
            allContactMemberships.Add(newMembership);
        }
        else if (tabControl.SelectedTab == addressTabPage)
        {
            if (curAddressMemberships.Contains(newMembership))
            {
                tagComboBox.SelectAll();
                tagComboBox.Focus();
                return;
            }
            curAddressMemberships.Add(newMembership);
            allAddressMemberships.Add(newMembership);
        }
        UpdateMembershipTags();
        UpdateMembershipCBox();
        UpdateMembershipJson();
        CheckSaveButton();
    }

    private void UpdateMembershipJson()
    {
        if (tabControl.SelectedTab == contactTabPage)
        {
            var rowIndex = contactDGV.SelectedRows.Count > 0 ? contactDGV.SelectedRows[0].Index : -1;
            if (rowIndex >= 0)
            {
                var newJson = JsonSerializer.Serialize(curContactMemberships);
                contactDGV.Rows[rowIndex].Cells["Gruppen"].Value = curContactMemberships.Count > 0 ? newJson : "";
            }
        }
        else if (tabControl.SelectedTab == addressTabPage)
        {
            var rowIndex = addressDGV.SelectedRows.Count > 0 ? addressDGV.SelectedRows[0].Index : -1;
            if (rowIndex >= 0 && addressDGV.Rows[rowIndex].DataBoundItem is DataRowView rowView)
            {
                var newJson = JsonSerializer.Serialize(curAddressMemberships);
                rowView.Row["Gruppen"] = curAddressMemberships.Count > 0 ? newJson : "";
            }
        }
    }

    private void UpdateMembershipCBox()
    {
        if (tabControl.SelectedTab == contactTabPage) { tagComboBox.DataSource = allContactMemberships.ToList(); }
        else { tagComboBox.DataSource = allAddressMemberships.ToList(); }
        tagComboBox.Text = ""; // Text zurücksetzen
    }

    private void UpdatePlaceholderVis()
    {
        if (flowLayoutPanel.Controls.Count == 0)
        {
            var lblPlaceholder = new Label
            {
                Text = "Gruppen",
                AutoSize = true,
                ForeColor = Color.Gray,
                BackColor = Color.Transparent,
                Name = "lblPlaceholder",
                Location = new Point(0, 0)
            };
            flowLayoutPanel.Controls.Add(lblPlaceholder);
        }
    }

    private void TagComboBox_TextChanged(object sender, EventArgs e)
    {
        tagButton.Enabled = !string.IsNullOrWhiteSpace(tagComboBox.Text);
        if (tagButton.Enabled)
        {
            tagButton.BackColor = SystemColors.MenuBar;
            tagButton.ForeColor = Color.Black;
            tagButton.Text = "Übernehmen";
        }
        else
        {
            tagButton.BackColor = SystemColors.InactiveBorder;
            tagButton.ForeColor = Color.Gray;
            tagButton.Text = string.Empty;
        }
    }

    private void TagComboBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            if (tagButton.Enabled) { TagButton_Click(tagButton, EventArgs.Empty); }
            else { SelectNextControl((Control)sender, true, true, true, true); }
            e.SuppressKeyPress = true;  //e.Handled = true;
        }
    }

    private void GroupFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        using var frm = new FrmGroupFilter(tabControl.SelectedTab == contactTabPage ? allContactMemberships : allAddressMemberships);
        if (frm.ShowDialog(this) == DialogResult.OK)
        {
            var includedGroups = frm.IncludedGroups;
            var excludedGroups = frm.ExcludedGroups;

            if (tabControl.SelectedTab == addressTabPage)
            {
                if (includedGroups.Count == 0 && excludedGroups.Count == 0)
                {
                    FilterAddressDGV(row => true); // Filter zurücksetzen
                    return;
                }
                var jsonOptions = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                bool rowFilter(DataGridViewRow row)
                {
                    if (row.DataBoundItem is not DataRowView dataRowView) { return false; } // Sollte nicht passieren, aber sicherheitshalber 
                    var jsonString = dataRowView.Row["Gruppen"]?.ToString();
                    if (string.IsNullOrWhiteSpace(jsonString)) { return includedGroups.Count == 0; } // Keine Gruppen
                    List<string> addressGroups;
                    try { addressGroups = JsonSerializer.Deserialize<List<string>>(jsonString, jsonOptions) ?? []; }
                    catch (JsonException) { return false; } // Ungültiges JSON, nicht anzeigen!?
                    var includeCondition = includedGroups.Count == 0 || addressGroups.Any(includedGroups.Contains);
                    var excludeCondition = addressGroups.Count == 0 || !addressGroups.Any(excludedGroups.Contains);
                    return includeCondition && excludeCondition;
                }
                FilterAddressDGV(rowFilter);
                filterRemoveToolStripMenuItem.Visible = true;
            }
            else if (tabControl.SelectedTab == contactTabPage)
            {
                if (includedGroups.Count == 0 && excludedGroups.Count == 0)
                {
                    FilterContactDGV(row => true); // Filter zurücksetzen
                    return;
                }
                var jsonOptions = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                bool rowFilter(DataGridViewRow row)
                {
                    var jsonString = row.Cells["Gruppen"].Value.ToString();
                    if (string.IsNullOrWhiteSpace(jsonString)) { return includedGroups.Count == 0; } // Keine Gruppen
                    List<string> contactGroups;
                    try { contactGroups = JsonSerializer.Deserialize<List<string>>(jsonString, jsonOptions) ?? []; }
                    catch (JsonException) { return false; } // Ungültiges JSON, nicht anzeigen!?
                    var includeCondition = includedGroups.Count == 0 || contactGroups.Any(includedGroups.Contains);
                    var excludeCondition = contactGroups.Count == 0 || !contactGroups.Any(excludedGroups.Contains);
                    return includeCondition && excludeCondition;
                }
                FilterContactDGV(rowFilter);
                filterRemoveToolStripMenuItem.Visible = true;
            }
            flexiTSStatusLabel.Text = "… mit Gruppenfilter";
        }
    }

    private void ManageGroupsToolStripMenuItem_Click(object sender, EventArgs e)
    {
        const string columnName = "Gruppen";
        Dictionary<string, int> groupDict = [];
        var jsonOptions = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        if (tabControl.SelectedTab == addressTabPage)
        {
            if (_dataTable is null || !_dataTable.Columns.Contains(columnName)) { return; }
            foreach (DataRow row in _dataTable.Rows)
            {
                if (row[columnName] is string jsonGroups && !string.IsNullOrWhiteSpace(jsonGroups))
                {
                    foreach (var groupName in Utilities.DeserializeGroups(Handle, jsonGroups, jsonOptions).Where(static g => !string.IsNullOrWhiteSpace(g)))
                    {
                        if (groupDict.TryGetValue(groupName, out var value)) { groupDict[groupName] = ++value; } // Key existiert: Wert um 1 erhöhen.
                        else { groupDict.Add(groupName, 1); } // Key existiert nicht: Hinzufügen mit dem Wert 1.
                    }
                }
            }
            using var frm = new FrmGroupsEdit(groupDict);
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                var groupChanges = frm.groupNameMap.Where(kvp => kvp.Key != kvp.Value || string.IsNullOrEmpty(kvp.Value)).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                foreach (var kvp in groupChanges)
                {
                    if (kvp.Key == "★") { continue; }
                    allAddressMemberships.Remove(kvp.Key);
                    if (kvp.Value is { Length: > 0 } newName) { allAddressMemberships.Add(newName); }
                }
                foreach (DataRow row in _dataTable.Rows)
                {
                    if (row[columnName] is string jsonGroups && !string.IsNullOrWhiteSpace(jsonGroups))
                    {
                        var currentGroups = Utilities.DeserializeGroups(Handle, jsonGroups, jsonOptions).ToList();
                        var newGroupsQuery = currentGroups.Select(group =>
                        {
                            if (groupChanges.TryGetValue(group, out var newName)) { return newName is { Length: > 0 } ? newName : null; }
                            return group;
                        })
                        .Where(g => g is not null).Distinct().ToList();
                        if (currentGroups.Any(groupChanges.ContainsKey)) { row[columnName] = JsonSerializer.Serialize(newGroupsQuery, jsonOptions); }
                    }
                }
                if (addressDGV.SelectedRows.Count > 0) { AddressEditFields(addressDGV.SelectedRows[0].Index); }
            }
        }
        else if (tabControl.SelectedTab == contactTabPage && contactDGV.Rows.Count > 0)
        {
            foreach (DataGridViewRow row in contactDGV.Rows)
            {
                if (row.Cells[columnName].Value is string jsonGroups && !string.IsNullOrWhiteSpace(jsonGroups))
                {
                    foreach (var groupName in Utilities.DeserializeGroups(Handle, jsonGroups, jsonOptions).Where(static g => !string.IsNullOrWhiteSpace(g)))
                    {
                        if (groupDict.ContainsKey(groupName)) { groupDict[groupName]++; } // Key existiert: Wert um 1 erhöhen.
                        else { groupDict.Add(groupName, 1); } // Key existiert nicht: Hinzufügen mit dem Wert 1.
                    }
                }
            }
            using var frm = new FrmGroupsEdit(groupDict);
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                List<string> deleteChanges = [];
                List<string> renameChanges = [];
                var groupChanges = frm.groupNameMap.Where(kvp => kvp.Key != kvp.Value || string.IsNullOrEmpty(kvp.Value)).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                foreach (var kvp in groupChanges)
                {
                    if (kvp.Value is { Length: > 0 } newName) { renameChanges.Add(newName); }
                    else { deleteChanges.Add(kvp.Key); }
                }
                var initialButtonYes = new TaskDialogButton("Hochladen");
                var initialButtonNo = TaskDialogButton.Cancel;
                using TaskDialogIcon questionDialogIcon = new(Resources.question32);
                initialButtonYes.AllowCloseDialog = false; // don't close the dialog when this button is clicked
                var initialPage = new TaskDialogPage()
                {
                    Caption = "Google Kontakte",
                    Heading = "Möchten Sie die Änderungen dauerhaft speichern?",
                    Text = "Änderungen an den Gruppen können nicht zurückgenommen werden.\nMitglieder einer Gruppe werden übrigens beim Löschen nicht entfernt.", // + Environment.NewLine +
                    Footnote = (renameChanges.Count > 0 ? $"Umbenennen: {string.Join(", ", renameChanges)}" : string.Empty) +
                    (deleteChanges.Count > 0 ? (renameChanges.Count > 0 ? Environment.NewLine : "") + $"zu Löschen: {string.Join(", ", deleteChanges)}" : string.Empty),
                    Icon = questionDialogIcon, // TaskDialogIcon.ShieldBlueBar,
                    AllowCancel = true,
                    SizeToContent = true,
                    Buttons = { initialButtonNo, initialButtonYes },
                };

                var inProgressCloseButton = TaskDialogButton.Close;
                inProgressCloseButton.Enabled = false;
                var progressPage = new TaskDialogPage()
                {
                    Caption = appCont,
                    Heading = "Bitte warten…",
                    Text = "Die Gruppenänderungen werden ausgeführt.",
                    Icon = TaskDialogIcon.Information,
                    ProgressBar = new TaskDialogProgressBar() { State = TaskDialogProgressBarState.Marquee },
                    Buttons = { inProgressCloseButton }
                };
                initialButtonYes.Click += (sender, e) => { initialPage.Navigate(progressPage); }; // When the user clicks "Yes", navigate to the second page.
                progressPage.Created += async (s, e) =>
                {
                    try
                    {
                        var service = await Utilities.GetPeopleServiceAsync(secretPath, tokenDir);
                        deleteChanges.Clear(); // jetzt ernsthaft angehen
                        renameChanges.Clear();
                        foreach (var kvp in groupChanges)
                        {
                            allContactMemberships.Remove(kvp.Key);
                            if (kvp.Value is { Length: > 0 } newName) // Kurze Form für !string.IsNullOrEmpty in C#/.NET
                            {
                                allContactMemberships.Add(newName);
                                renameChanges.Add(newName);
                            }
                            else { deleteChanges.Add(kvp.Key); }
                        }
                        foreach (DataGridViewRow row in contactDGV.Rows)
                        {
                            if (row.Cells[columnName].Value is string jsonGroups && !string.IsNullOrWhiteSpace(jsonGroups))
                            {
                                var currentGroups = Utilities.DeserializeGroups(Handle, jsonGroups, jsonOptions).ToList();
                                var newGroupsQuery = currentGroups.Select(group =>
                                {
                                    if (groupChanges.TryGetValue(group, out var newName)) { return newName is { Length: > 0 } ? newName : null; } // group wurde nicht geändert
                                    return group;
                                })
                                .Where(g => g is not null).Distinct().ToList();
                                if (currentGroups.Any(groupChanges.ContainsKey)) { row.Cells[columnName].Value = JsonSerializer.Serialize(newGroupsQuery, jsonOptions); }
                            }
                        }
                        if (contactDGV.SelectedRows.Count > 0) { ContactEditFields(contactDGV.SelectedRows[0].Index); }
                        var nameToResourceNameDict = contactGroupsDict.ToDictionary(kvp => kvp.Value, kvp => kvp.Key); // Umkehrung: Name -> ResourceName
                        foreach (var kvp in frm.groupNameMap)
                        {
                            if (nameToResourceNameDict.TryGetValue(kvp.Key, out var resourceName))
                            {
                                var newName = kvp.Value;
                                if (string.IsNullOrEmpty(newName))
                                {
                                    using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(9));
                                    await service.ContactGroups.Delete(resourceName).ExecuteAsync(cts.Token);  // Gruppe löschen
                                }
                                else if (resourceName != newName)
                                {
                                    var group = await service.ContactGroups.Get(resourceName).ExecuteAsync();
                                    var updateRequest = new UpdateContactGroupRequest
                                    {
                                        ContactGroup = new ContactGroup
                                        {
                                            ETag = group.ETag,
                                            ResourceName = group.ResourceName,
                                            Name = newName,
                                        },
                                        UpdateGroupFields = "name"
                                    };
                                    using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(9));
                                    await service.ContactGroups.Update(updateRequest, resourceName).ExecuteAsync(cts.Token);  // Gruppe umbenennen
                                }
                            }
                        }
                        UpdateContactGroupsDict(service);
                    }
                    catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
                    finally { progressPage.Buttons.First().PerformClick(); }
                };
                TaskDialog.ShowDialog(Handle, initialPage); // Show the initial page of the TaskDialog
            }
        }
    }

    private async void UpdateContactGroupsDict(PeopleServiceService service)
    {
        contactGroupsDict.Clear(); // Kontaktgruppen abrufen, auf den neuesten Stand bringen    
        allContactMemberships.Clear(); // Auch die Liste der Mitgliedschaften leeren
        var groupResponse = await service.ContactGroups.List().ExecuteAsync();
        if (groupResponse?.ContactGroups != null)
        {
            foreach (var group in groupResponse.ContactGroups)
            {
                var gName = group.Name;
                if (!contactGroupsDict.ContainsValue(gName)) { contactGroupsDict.Add(group.ResourceName, gName); }
                if (!excludedGroups.Contains(gName))
                {
                    gName = gName.Equals("starred") ? "★" : gName;
                    allContactMemberships.Add(gName);
                }
            }
        }
    }

    private void FlowLayoutPanel_MouseDoubleClick(object sender, MouseEventArgs e) => ManageGroupsToolStripMenuItem_Click(null!, EventArgs.Empty);

    private void CopyCellToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage) { CopyCurrentCellToClipboard(addressDGV); }
        else if (tabControl.SelectedTab == contactTabPage) { CopyCurrentCellToClipboard(contactDGV); }
    }

    private void CopyCurrentCellToClipboard(DataGridView myDataGridView)
    {
        if (myDataGridView.CurrentCell != null && myDataGridView.CurrentCell.Value is string strValue && !string.IsNullOrWhiteSpace(strValue))
        {
            try { Utilities.SetClipboardText(strValue); }
            catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
        }
    }

    //    protected override void WndProc(ref Message m)  // nicht löschen, wird unter .NET 9 funktionieren
    //    {
    //        const int WM_SETTINGCHANGE = 0x001A;

    //        base.WndProc(ref m);

    //        if (m.Msg == WM_SETTINGCHANGE)
    //        {
    //#pragma warning disable WFO5001
    //            Application.SetColorMode(SystemColorMode.System);
    //#pragma warning restore WFO5001
    //            this.Invalidate(true); // Repaint erzwingen
    //        }
    //    }
}
