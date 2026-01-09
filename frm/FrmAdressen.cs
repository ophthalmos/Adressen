using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
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
using Microsoft.EntityFrameworkCore;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;

namespace Adressen;

public partial class FrmAdressen : Form
{
    private readonly FrmSplashScreen? _splashScreen;
    private static readonly string appPath = Application.ExecutablePath; // EXE-Pfad
    private string _databaseFilePath = string.Empty; // Path.ChangeExtension(appPath, ".adb");
    private bool sAskBeforeSaveSQL = true; // false = Änderungen automatisch speichern
    private AppSettings _settings = new(); // Ein einziges Objekt für alle Einstellungen
    private AdressenDbContext? _context;
    private readonly string _settingsPath;
    private readonly string tokenDir;
    private readonly string secretPath;
    private readonly string boysPath;
    private readonly string girlPath;
    private readonly string cleanRegex = @"[^\+0-9]";
    private readonly string appLong = Application.ProductName ?? "Adressen & Kontakte";
    private readonly string appName = "Adressen";
    private readonly string appCont = "Kontakte";
    private readonly Dictionary<string, string> bookmarkTextDictionary = [];  // wird aus den Edit-Controls befüllt, Datenbank unabhängig
    private readonly Dictionary<Control, string> editControlsDictionary = [];
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
    private readonly string[] dataFields = ["Anrede", "Praefix", "Nachname", "Vorname", "Zwischenname", "Nickname",
        "Suffix", "Firma", "Strasse", "PLZ", "Ort", "Land", "Betreff", "Grussformel", "Schlussformel", "Geburtstag",
        "Mail1", "Mail2", "Telefon1", "Telefon2", "Mobil", "Fax", "Internet", "Notizen"]; // Id fehlt absichtlich  
    private bool[] hideColumnArr = new bool[25]; // muss angepasst werden, wenn Felder/Spalten hinzugefügt werden
    private readonly bool[] hideColumnStd = [true, true, false, false, true, true, true, false, false, false, false, false, true, true, true, false, false, false, false, false, false, false, false, true, true]; // muss angepasst werden, wenn Felder/Spalten hinzugefügt werden
    private int[] columnWidths = [100, 100, 200, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100];
    private int splitterPosition;
    private WindowPlacement? windowPosition;
    private bool windowMaximized = false;
    private readonly bool argsPath = false;
    private Word.Document? wordDoc;
    private dynamic? wordApp; // Word.Application
    private int contactNewRowIndex = -1;
    private bool isSelectionChanging = false;
    private int birthdayRemindLimit = 30;
    private int birthdayRemindAfter = 3;
    private bool birthdayAddressShow = false;
    private bool birthdayContactShow = false;
    private bool ignoreTextChange = false; // ignore when changing text in ContactEditFields
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
    private readonly string lastSearchText = string.Empty;
    private TabPage? deactivatedPage = null;
    private List<ListViewItem> allDokuLVItems = [];
    private int lastColumn = -1;
    private SortOrder lastOrder = SortOrder.None;
    private string lastTooltipText = string.Empty;
    private bool contactBirthdayFlag = true; // false wenn Zugriffstoken für Google-Kontakte fehlt oder abgelaufen ist
    private readonly string[] documentTypes = ["*.doc", "*.dot", "*.docx", "*.doct", "*.docm", "*.odt", "*.ott", "*.fodt", "*.uot", "*.pdf", "*.txt"];
    private readonly List<string> addressCbItems_Anrede = [];
    private readonly List<string> addressCbItems_Präfix = [];
    private readonly List<string> addressCbItems_PLZ = [];
    private readonly List<string> addressCbItems_Ort = [];
    private readonly List<string> addressCbItems_Land = [];
    private readonly List<string> addressCbItems_Schlussformel = [];
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
    private readonly SortedSet<string> curAddressMemberships = [];
    private SortedSet<string> curContactMemberships = [];
    private Contact? _lastActiveContact; // Merkt sich den Kontakt, der VOR dem Wechsel aktiv war
    private Contact? _originalContactSnapshot;
    private readonly Dictionary<string, string> contactGroupsDict = [];
    private static readonly HashSet<string> excludedGroups = ["myContacts", "all", "blocked", "chatBuddies", "coworkers", "family", "friends"];
    private string userEmail = string.Empty;
    private bool _isClosing = false; // Flag, um Endlosschleife zu verhindern
    private bool _isFiltering = false; // Verhindert Speichern während der Suche
    private BindingList<Contact> _allGoogleContacts = []; // Klassenvariable
    private bool _isDarkMode;

    public FrmAdressen(FrmSplashScreen? splashScreen, string[] args)
    {
        if (args.Length >= 1)
        {
            if (File.Exists((string?)args[0]))
            {
                _databaseFilePath = (string?)args[0] ?? string.Empty;
                if (!string.IsNullOrEmpty(_databaseFilePath)) { argsPath = true; }
            }
        }

        InitializeComponent();
        _splashScreen = splashScreen;  // Splash Screen speichern um ihn beenden zu können (s. Load Event)  
        typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, addressDGV, [true]);
        typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, contactDGV, [true]);
        _isDarkMode = DefaultBackColor.R < 128;
        UpdateAppearanceStatus(); // Basis-Farben setzen
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

        if (Utils.IsInnoSetupValid(Path.GetDirectoryName(appPath)!))
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
        foreach (var field in dataFields.Concat(extraFields)) { bookmarkTextDictionary[field] = string.Empty; }

        editControlsDictionary.Add(cbAnrede, "Anrede");
        editControlsDictionary.Add(cbPräfix, "Praefix");
        editControlsDictionary.Add(tbNachname, "Nachname");
        editControlsDictionary.Add(tbVorname, "Vorname");
        editControlsDictionary.Add(tbZwischenname, "Zwischenname");
        editControlsDictionary.Add(tbNickname, "Nickname");
        editControlsDictionary.Add(tbSuffix, "Suffix");
        editControlsDictionary.Add(tbFirma, "Firma");
        editControlsDictionary.Add(tbStraße, "Strasse");
        editControlsDictionary.Add(cbPLZ, "PLZ");
        editControlsDictionary.Add(cbOrt, "Ort");
        editControlsDictionary.Add(cbLand, "Land");
        editControlsDictionary.Add(tbBetreff, "Betreff");
        editControlsDictionary.Add(cbGrußformel, "Grussformel");
        editControlsDictionary.Add(cbSchlussformel, "Schlussformel");
        editControlsDictionary.Add(tbMail1, "Mail1");
        editControlsDictionary.Add(tbMail2, "Mail2");
        editControlsDictionary.Add(tbTelefon1, "Telefon1");
        editControlsDictionary.Add(tbTelefon2, "Telefon2");
        editControlsDictionary.Add(tbMobil, "Mobil");
        editControlsDictionary.Add(tbFax, "Fax");
        editControlsDictionary.Add(tbInternet, "Internet");
        editControlsDictionary.Add(tbNotizen, "Notizen");

        fileToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
        editToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
        viewToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
        extraToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
        helpToolStripMenuItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening);
    }

    private async void FrmAdressen_Load(object sender, EventArgs e)
    {
        if (File.Exists(_settingsPath)) { await LoadConfiguration(); }
        else { Directory.CreateDirectory(Path.GetDirectoryName(_settingsPath)!); } // If the folder exists already, the line will be ignored.     
        _databaseFilePath = argsPath ? _databaseFilePath : recentFiles.Count > 0 ? recentFiles[0] : string.Empty;
        if (!(new int[] { hideColumnArr.Length, hideColumnStd.Length, columnWidths.Length }).All(len => len == dataFields.Length + 1))
        {
            var text = $"Datenfelder: {dataFields.Length + 1}\nhideColumnArr: {hideColumnArr.Length}\nhideColumnStd: {hideColumnStd.Length}\ncolumnWidths: {columnWidths.Length}";
            Utils.MsgTaskDlg(Handle, "Fehler bei der Initialisierung", "Nicht alle Arrays haben die gleiche Länge.\n" + text);
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
        catch (Exception ex) { Utils.MsgTaskDlg(Handle, "Fehler beim Laden der Namenslisten", ex.Message); }

        NativeMethods.SendMessage(searchTSTextBox.TextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 4 << 16);
        NativeMethods.SendMessage(searchTSTextBox.TextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 4);
        NativeMethods.SendMessage(tbNotizen.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 4 << 16);
        NativeMethods.SendMessage(tbNotizen.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 4);
        NativeMethods.SendMessage(maskedTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 4 << 16);
        NativeMethods.SendMessage(maskedTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 4);
        _ = NativeMethods.SendMessage(maskedTextBox.Handle, NativeMethods.EM_SETCUEBANNER, 0, "TT.MM.JJJJ");

        SetColorScheme();
        if ((sReloadRecent || argsPath) && !string.IsNullOrEmpty(_databaseFilePath)) { ConnectSQLDatabase(_databaseFilePath); }
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
        hideColumnArr = (_settings.HideColumnArr.Length > 0 && _settings.HideColumnArr.Length <= hideColumnArr.Length) ? _settings.HideColumnArr : hideColumnArr;
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
        if (sContactsAutoload) { await LoadAndDisplayGoogleContactsAsync(); }
    }

    private bool MigrateLegacyData(AdressenDbContext context)
    {
        if (context == null) { return false; }
        var changesMade = false;

        try
        {
            // 1. Aktuelle Spalten in der DB ermitteln (Case-Insensitive HashSet für schnelle Suche)
            var dbColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var command = context.Database.GetDbConnection().CreateCommand())
            {
                command.CommandText = "SELECT name FROM pragma_table_info('Adressen')";
                context.Database.OpenConnection();
                using var reader = command.ExecuteReader();
                while (reader.Read()) { dbColumns.Add(reader.GetString(0)); }
            }

            // --- SCHRITT A: Spalten umbenennen (Legacy German -> English) ---
            // Wir führen das sofort aus und aktualisieren unser HashSet
            if (dbColumns.Contains("Grußformel"))
            {
                context.Database.ExecuteSqlRaw("ALTER TABLE Adressen RENAME COLUMN \"Grußformel\" TO Grussformel");
                dbColumns.Remove("Grußformel"); dbColumns.Add("Grussformel"); changesMade = true;
            }
            if (dbColumns.Contains("Straße"))
            {
                context.Database.ExecuteSqlRaw("ALTER TABLE Adressen RENAME COLUMN \"Straße\" TO Strasse");
                dbColumns.Remove("Straße"); dbColumns.Add("Strasse"); changesMade = true;
            }
            if (dbColumns.Contains("Präfix"))
            {
                context.Database.ExecuteSqlRaw("ALTER TABLE Adressen RENAME COLUMN \"Präfix\" TO Praefix");
                dbColumns.Remove("Präfix"); dbColumns.Add("Praefix"); changesMade = true;
            }

            // --- SCHRITT B: Fehlende Spalten ergänzen (z.B. Nickname, Zwischenname) ---
            // Wir nutzen Reflection, um alle Properties der Klasse Adresse zu prüfen, die in der DB sein sollten.
            var entityProperties = typeof(Adresse).GetProperties()
                .Where(p => p.Name != "Id"
                         && !Attribute.IsDefined(p, typeof(NotMappedAttribute))
                         && !p.GetAccessors().Any(x => x.IsVirtual)); // Navigation Properties ignorieren

            foreach (var prop in entityProperties)
            {
                if (!dbColumns.Contains(prop.Name))
                {
#pragma warning disable EF1002 // Spalte existiert nicht -> Anlegen! Disable Warning, weil ALTER TABLE keine Parameter unterstützt und prop.Name (Reflection) sicher ist
                    context.Database.ExecuteSqlRaw($"ALTER TABLE Adressen ADD COLUMN \"{prop.Name}\" TEXT");
#pragma warning restore EF1002
                    changesMade = true;
                }
            }

            // --- SCHRITT C: Tabellenstruktur für Relationen sicherstellen ---
            // Falls die Migration von einer sehr alten Version kommt, fehlen diese Tabellen vielleicht trotz EnsureCreated,
            // da EnsureCreated nichts macht, wenn die DB (wegen Tabelle Adressen) schon existiert.
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""Gruppen"" (""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, ""Name"" TEXT NOT NULL);");
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""AdresseGruppen"" (""AdressenId"" INTEGER NOT NULL, ""GruppenId"" INTEGER NOT NULL, PRIMARY KEY(""AdressenId"", ""GruppenId""), FOREIGN KEY(""AdressenId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE, FOREIGN KEY(""GruppenId"") REFERENCES ""Gruppen""(""Id"") ON DELETE CASCADE);");
            context.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""Dokumente"" (""Id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, ""Dateipfad"" TEXT NOT NULL, ""AdressId"" INTEGER NOT NULL, FOREIGN KEY(""AdressId"") REFERENCES ""Adressen""(""Id"") ON DELETE CASCADE);");


            // --- SCHRITT D: Daten aus alten JSON/Format-Spalten migrieren ---
            var hasOldGruppen = dbColumns.Contains("Gruppen");
            var hasOldDokumente = dbColumns.Contains("Dokumente");

            // Check auf alte Datumsformate
            var hasOldDateFormats = false;
            using (var command = context.Database.GetDbConnection().CreateCommand())
            {
                // Prüfen ob Punkte im Datum sind (deutsch) oder Leerstrings (altes Legacy)
                command.CommandText = "SELECT 1 FROM Adressen WHERE Geburtstag LIKE '%.%' OR Geburtstag = '' LIMIT 1";
                using var reader = command.ExecuteReader();
                hasOldDateFormats = reader.HasRows;
            }

            if (hasOldGruppen || hasOldDokumente || hasOldDateFormats)
            {
                // SQL Dynamisch bauen, um "no such column" Fehler zu vermeiden
                var sbSql = new System.Text.StringBuilder();
                sbSql.Append("SELECT Id, NULLIF(CAST(Geburtstag AS TEXT), '') AS Geburtstag");

                if (hasOldGruppen) sbSql.Append(", Gruppen");
                else sbSql.Append(", NULL AS Gruppen"); // Dummy für Record

                if (hasOldDokumente) sbSql.Append(", Dokumente");
                else sbSql.Append(", NULL AS Dokumente"); // Dummy für Record

                sbSql.Append(" FROM Adressen");

                var legacyData = context.Database.SqlQueryRaw<LegacyRawData>(sbSql.ToString()).ToList();

                // Geburtstag temporär nullen für sauberes EF-Laden
                context.Database.ExecuteSqlRaw("UPDATE Adressen SET Geburtstag = NULL;");

                var gruppenCache = new Dictionary<string, Gruppe>(StringComparer.OrdinalIgnoreCase);

                // Kontext neu laden oder Tracking beachten ist hier schwierig, wir arbeiten direkt.
                // Um EF-Konflikte zu vermeiden, laden wir die Adressen frisch.
                var allAdressen = context.Adressen.Include(a => a.Gruppen).Include(a => a.Dokumente).ToList();

                foreach (var row in legacyData)
                {
                    var adresse = allAdressen.FirstOrDefault(a => a.Id == row.Id);
                    if (adresse == null) continue;

                    var dataChanged = false;

                    // 1. Geburtstag fixen
                    if (!string.IsNullOrWhiteSpace(row.Geburtstag))
                    {
                        // Versuche verschiedene Formate
                        DateOnly parsedDate;
                        if (DateOnly.TryParseExact(row.Geburtstag, "d.M.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate) ||
                            DateOnly.TryParse(row.Geburtstag, CultureInfo.GetCultureInfo("de-DE"), DateTimeStyles.None, out parsedDate))
                        {
                            adresse.Geburtstag = parsedDate;
                            dataChanged = true;
                        }
                    }

                    // 2. Gruppen migrieren
                    if (hasOldGruppen && !string.IsNullOrWhiteSpace(row.Gruppen))
                    {
                        try
                        {
                            var namen = System.Text.Json.JsonSerializer.Deserialize<List<string>>(row.Gruppen);
                            if (namen != null)
                            {
                                foreach (var name in namen.Where(n => !string.IsNullOrWhiteSpace(n)))
                                {
                                    if (!gruppenCache.TryGetValue(name, out var gruppe))
                                    {
                                        gruppe = context.Gruppen.Local.FirstOrDefault(g => g.Name == name)
                                                 ?? context.Gruppen.FirstOrDefault(g => g.Name == name)
                                                 ?? new Gruppe { Name = name };

                                        if (gruppe.Id == 0 && !context.Gruppen.Local.Contains(gruppe))
                                        {
                                            context.Gruppen.Add(gruppe); // Neu hinzufügen
                                        }
                                        gruppenCache[name] = gruppe;
                                    }
                                    if (!adresse.Gruppen.Any(g => g.Name == name))
                                    {
                                        adresse.Gruppen.Add(gruppe);
                                        dataChanged = true;
                                    }
                                }
                            }
                        }
                        catch { /* Ignore invalid JSON */ }
                    }

                    // 3. Dokumente migrieren
                    if (hasOldDokumente && !string.IsNullOrWhiteSpace(row.Dokumente))
                    {
                        try
                        {
                            var pfade = System.Text.Json.JsonSerializer.Deserialize<List<string>>(row.Dokumente);
                            if (pfade != null)
                            {
                                foreach (var pfad in pfade.Where(p => !string.IsNullOrWhiteSpace(p)))
                                {
                                    if (!adresse.Dokumente.Any(d => d.Dateipfad == pfad))
                                    {
                                        adresse.Dokumente.Add(new Dokument { Dateipfad = pfad });
                                        dataChanged = true;
                                    }
                                }
                            }
                        }
                        catch { /* Ignore invalid JSON */ }
                    }

                    // Wir speichern im Batch am Ende, aber markieren hier Änderungen falls nötig
                    if (dataChanged)
                    {
                        context.Entry(adresse).State = EntityState.Modified;
                    }
                }

                context.SaveChanges();
                changesMade = true;

                // --- SCHRITT E: Aufräumen ---
                if (hasOldGruppen) { context.Database.ExecuteSqlRaw("ALTER TABLE Adressen DROP COLUMN Gruppen"); }
                if (hasOldDokumente) { context.Database.ExecuteSqlRaw("ALTER TABLE Adressen DROP COLUMN Dokumente"); }
            }

            if (changesMade)
            {
                context.Database.ExecuteSqlRaw("VACUUM;");
                Utils.MsgTaskDlg(Handle, "Migration erfolgreich", "Die Datenbank wurde auf das neue Format aktualisiert.", TaskDialogIcon.ShieldSuccessGreenBar);
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            return false;
        }
    }

    internal record LegacyRawData(int Id, string? Gruppen, string? Dokumente, string? Geburtstag); // Hilfsrecord für die Migration

    private void ConnectSQLDatabase(string file)
    {
        flexiTSStatusLabel.Text = string.Empty;
        if (string.IsNullOrEmpty(file) || !File.Exists(file))
        {
            Utils.MsgTaskDlg(Handle, "Datenbank-Datei nicht gefunden", file, TaskDialogIcon.ShieldWarningYellowBar);
            recentFiles.Remove(file);
            return;
        }
        try
        {
            CloseDatabaseConnection();
            _databaseFilePath = file;
            _context = new AdressenDbContext(_databaseFilePath);
            _context.Database.EnsureCreated(); // Tabellenstruktur sicherstellen
            var migrationDone = MigrateLegacyData(_context);  // Migration durchführen (gibt true zurück, wenn migriert wurde)
            _context.Adressen.Include(a => a.Gruppen).Include(a => a.Dokumente).OrderBy(a => a.Nachname).ThenBy(a => a.Vorname).Load(); // ohne Fotos
            addressBindingSource.DataSource = _context.Adressen.Local.ToBindingList();
            addressDGV.DataSource = addressBindingSource;
            // Deaktiviert das Sortieren für alle Spalten
            foreach (DataGridViewColumn column in addressDGV.Columns) { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
            PopulateMemberships();  // ComboBoxen und Listen füllen
            SwitchDataBinding(addressBindingSource); // EditControls mit BindingSource verbinden
            if (_context != null)
            {
                recentFiles.Remove(_databaseFilePath);
                recentFiles.Insert(0, _databaseFilePath);
                if (recentFiles.Count > maxRecentFiles) { recentFiles = [.. recentFiles.Take(maxRecentFiles)]; }
                newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = deleteTSButton.Enabled = newTSButton.Enabled = duplicateToolStripMenuItem.Enabled = copyTSButton.Enabled = wordTSButton.Enabled = envelopeTSButton.Enabled = true; copyToOtherDGVTSMenuItem.Enabled = false;
                Utils.ApplyColumnSettings(addressDGV, columnWidths, hideColumnArr);
                tabControl.SelectTab(0);
                //addressBindingSource.CurrentChanged -= AddressBindingSource_CurrentChanged;
                //addressBindingSource.CurrentChanged += AddressBindingSource_CurrentChanged;

                //// 2. ListChanged (Änderung im Grid/Daten -> Save Button)
                //// Wir nutzen hier ein Lambda, daher können wir es schwer "detachen". 
                //// Sauberer ist es, wenn du eine Methode UpdateSaveButtonWrapper(object s, EventArgs e) hättest.
                //// Für hier reicht es, das Event bei CloseDatabaseConnection() zu leeren oder zu akzeptieren.
                //// Besser: Wir hängen es nur an, wenn wir sicher sind.
                //addressBindingSource.ListChanged += (s, e) => UpdateSaveButton();

                // 3. EF Core ChangeTracker (Reagiert auf .Add/.Remove im Code)
                _context.ChangeTracker.StateChanged += (s, e) => UpdateSaveButton();
                // Tracked ist oft redundant zu StateChanged, kann aber nicht schaden
                // _context.ChangeTracker.Tracked += (s, e) => UpdateSaveButton(); 
                addressBindingSource.CurrentChanged += AddressBindingSource_CurrentChanged; // Wurde in CloseDatabaseConnection entfernt, um Doppelbindungen zu vermeiden


                // UI initialisieren (erstes Element laden)
                if (addressBindingSource.Count > 0) { AddressBindingSource_CurrentChanged(this, EventArgs.Empty); }

                // Geburtstags-Reminder (nur wenn keine Migration lief, um User nicht zu nerven)
                if (!migrationDone && birthdayAddressShow) { BeginInvoke(new Action(() => { BirthdayReminder(addressDGV); })); }
                Utils.StartSearchCacheWarmup(_context.Adressen.Local);
            }
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            //addressBindingSource.CurrentChanged -= AddressBindingSource_CurrentChanged;
            //addressBindingSource.CurrentChanged += AddressBindingSource_CurrentChanged;
        }
    }

    // Optimierte PopulateMemberships (C# 14)
    private void PopulateMemberships()
    {
        if (addressBindingSource is null || _context is null) { return; }
        allAddressMemberships.Clear();
        allAddressMemberships.Add("★"); // Favoriten immer zuerst
        var dbGruppen = _context.Gruppen.Select(g => g.Name).Distinct().ToList();
        allAddressMemberships.UnionWith(dbGruppen);
        UpdateMembershipCBox();
    }

    private void CreateNewDatabase(string filePath, bool addSampleRecord = false)
    {
        try
        {
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools(); // bestehende Pools leeren, um Dateisperren zu vermeiden
            if (File.Exists(filePath)) { File.Delete(filePath); }
            using var dbContext = new AdressenDbContext(filePath);
            dbContext.Database.EnsureCreated(); // Erstellt die Datenbank und ALLE Tabellen (Adressen, Gruppen, Dokumente, Foto)
            if (addSampleRecord)
            {
                var sampleAdresse = new Adresse
                {
                    Anrede = "Herrn",
                    Praefix = "Dr. h.c.",
                    Nachname = "Mustermann",
                    Vorname = "Max",
                    Zwischenname = "von und zu",
                    Nickname = "Maxi",
                    Suffix = "Jr. MBA",
                    Strasse = "Langer Weg 144",
                    PLZ = "01234",
                    Ort = "Entenhausen",
                    Grussformel = "Lieber Max",
                    Geburtstag = DateOnly.ParseExact("6.3.1995", "d.M.yyyy", CultureInfo.InvariantCulture),
                    Mail1 = "abc@xyz.com"
                };
                dbContext.Adressen.Add(sampleAdresse);
                dbContext.SaveChanges();
            }
            dbContext.Database.ExecuteSqlRaw($"PRAGMA user_version = {latestSchemaVersion}"); // Schema-Version setzen, wenn Tabellen existieren
        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
    }

    private async Task SaveSQLDatabaseAsync(bool closeDB = false, bool askNever = false, bool isFormClosing = false)
    {
        if (_context == null)
        {
            if (closeDB) { CloseDatabaseConnection(); }
            return;
        }

        // --- SICHERHEITS-CHECK ---
        // Wenn die Liste leer ist (z.B. nach Löschen des letzten Elements),
        // dürfen wir keine Validierung anstoßen, sonst knallt es beim Binding ("Anrede").
        if (addressBindingSource.Count == 0 || addressBindingSource.Current == null)
        {
            if (closeDB) { CloseDatabaseConnection(); }
            return;
        }
        // -------------------------
        // 1. ChangeTracker aktualisieren (fängt letzte Binding-Updates beim Fokusverlust ab)
        _context.ChangeTracker.DetectChanges();

        // 2. Die "echten" Änderungen ermitteln
        // Wir nutzen hier direkt die strenge Filter-Logik
        var realChanges = _context.ChangeTracker.Entries().Where(IsEntryReallyChanged).ToList();

        // 3. WICHTIG: Wenn die Liste leer ist, sofort abbrechen!
        // Selbst wenn EF Core sagt "HasChanges", sind das nur Phantome, die wir ignorieren wollen.
        if (realChanges.Count == 0)
        {
            if (closeDB) { CloseDatabaseConnection(); }
            return;
        }

        // 4. Ab hier wissen wir sicher: Es gibt mind. 1 echte Änderung
        if (!askNever && sAskBeforeSaveSQL)
        {
            var changesCount = realChanges.Count;

            // Dialog anzeigen
            var (isYes, isNo, isCancelled) = Utils.YesNo_TaskDialog(this, appName, "Möchten Sie die Änderungen speichern?",
                changesCount == 1 ? "An einer Adresse wurden Änderungen vorgenommen." : $"Änderungen wurden an {changesCount} Adressen vorgenommen.",
                "&Speichern", "&Ignorieren");
            if (isNo) // Ignorieren gewählt
            {
                foreach (var entry in _context.ChangeTracker.Entries().ToList())
                {
                    switch (entry.State)
                    {
                        case EntityState.Modified:
                        case EntityState.Deleted:
                            await entry.ReloadAsync().ConfigureAwait(false);
                            break;
                        case EntityState.Added:
                            entry.State = EntityState.Detached;
                            break;
                    }
                }
                saveTSButton.Enabled = false;
                if (closeDB) { CloseDatabaseConnection(); }
                return;
            }
            else if (isCancelled) { return; } // Escape-Taste gedrückt
        }
        // 5. Speichern durchführen
        try
        {
            await _context.SaveChangesAsync().ConfigureAwait(false);
            if (!isFormClosing)
            {
                Invoke(() =>
                {
                    saveTSButton.Enabled = false;
                    flexiTSStatusLabel.Text = $"Letztes Speichern: {DateTime.Now:HH:mm:ss}";
                });
            }
            if (sDailyBackup && File.Exists(Utils.CorrectUNC(_databaseFilePath)) && Directory.Exists(sBackupDirectory))
            {
                if (isFormClosing) { Utils.DailyBackup(Utils.CorrectUNC(_databaseFilePath), sBackupDirectory, sBackupSuccess, sSuccessDuration, true); }
                else { await Task.Run(() => { Utils.DailyBackup(Utils.CorrectUNC(_databaseFilePath), sBackupDirectory, sBackupSuccess, sSuccessDuration, false); }); }
            }
        }
        catch (DbUpdateConcurrencyException dbEx)
        {
            Utils.MsgTaskDlg(Handle, "Konflikt beim Speichern", $"Details: {dbEx.Message}\nIhre lokalen Änderungen werden verworfen.");
            foreach (var entry in dbEx.Entries) { await entry.ReloadAsync(); }
            saveTSButton.Enabled = false;
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            saveTSButton.Enabled = false;
        }
        finally
        {
            if (closeDB) { CloseDatabaseConnection(); }
        }
    }

    private static bool IsEntryReallyChanged(Microsoft.EntityFrameworkCore.ChangeTracking.EntityEntry entry)
    {
        if (entry.State == EntityState.Added || entry.State == EntityState.Deleted) { return true; }
        if (entry.State != EntityState.Modified) { return false; }

        foreach (var prop in entry.Properties)
        {
            if (!prop.IsModified) { continue; }

            var current = prop.CurrentValue;
            var original = prop.OriginalValue;

            // 1. Direkter Vergleich (fängt int, date, bool ab)
            if (Equals(original, current)) { continue; }

            // 2. Spezialbehandlung für Strings
            if (prop.Metadata.ClrType == typeof(string))
            {
                var sOriginal = original as string;
                var sCurrent = current as string;

                // Behandle null wie leeren String ("")
                var sOrigClean = sOriginal ?? string.Empty;
                var sCurrClean = sCurrent ?? string.Empty;

                // Optional: .Trim(), falls " " und "" gleich sein sollen
                // sOrigClean = sOrigClean.Trim();
                // sCurrClean = sCurrClean.Trim();

                if (sOrigClean == sCurrClean)
                {
                    continue; // Es war nur ein null vs. "" Unterschied -> Keine echte Änderung
                }
            }

            // Wenn wir hier ankommen, sind die Werte wirklich unterschiedlich
            return true;
        }

        return false;
    }

    //private void CloseDatabaseConnection()
    //{
    //    if (addressBindingSource != null)  // Event aushängen, um Memory Leaks zu vermeiden
    //    {
    //        addressBindingSource.CurrentChanged -= AddressBindingSource_CurrentChanged;
    //        addressBindingSource.DataSource = null; // Wichtig: Erst DataSource nullen
    //    }
    //    if (editControlsDictionary != null)
    //    {
    //        foreach (var control in editControlsDictionary.Keys) { control.DataBindings.Clear(); } // Bindings lösen
    //    }
    //    maskedTextBox?.DataBindings.Clear();
    //    addressDGV?.DataSource = null;
    //    _context?.Dispose();
    //    _context = null;
    //}

    private void CloseDatabaseConnection()
    {
        if (addressBindingSource != null) { addressBindingSource.CurrentChanged -= AddressBindingSource_CurrentChanged; }
        _context?.ChangeTracker.StateChanged -= (s, e) => UpdateSaveButton();
        AutoValidate = AutoValidate.Disable; // Die UI-Controls komplett von den Datenquellen trennen
        if (editControlsDictionary != null)
        {
            foreach (var control in editControlsDictionary.Keys) { control.DataBindings.Clear(); }
        }
        maskedTextBox?.DataBindings.Clear();
        addressBindingSource?.DataSource = null;
        contactBindingSource?.DataSource = null;
        addressDGV?.DataSource = null;
        contactDGV?.DataSource = null;
        _context?.Dispose();
        _context = null;
        Debug.WriteLine("Datenbankverbindung sicher getrennt.");
    }

    private async void OpenToolStripMenuItem_Click(object? sender, EventArgs? e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact lastContact)
        {
            if (string.IsNullOrEmpty(lastContact.ResourceName) && CheckNewContactTidyUp()) { await CreateContactAsync(); }
            else if (ContactChanges_Check()) { await AskSaveContactChangesAsync(); }
        }
        openFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";

        var fullPath = Utils.CorrectUNC(_databaseFilePath);
        var fileName = Path.GetFileName(fullPath) ?? "Adressen.adb";
        var dirName = Path.GetDirectoryName(fullPath);

        openFileDialog.FileName = fileName;
        openFileDialog.InitialDirectory = !string.IsNullOrEmpty(sDatabaseFolder) && Directory.Exists(sDatabaseFolder) ? sDatabaseFolder : dirName ?? string.Empty;

        openFileDialog.Multiselect = false;

        if (openFileDialog.ShowDialog(this) == DialogResult.OK)
        {
            if (addressBindingSource != null && _context != null) { await SaveSQLDatabaseAsync(true); }
            ConnectSQLDatabase(openFileDialog.FileName);
            ignoreSearchChange = true;
            searchTSTextBox.Text = string.Empty;
            ApplyGlobalSearch(string.Empty); // Filter komplett zurücksetzen
            ignoreSearchChange = false;
        }
    }

    private async void ExitToolStripMenuItem_Click(object? sender, EventArgs? e)
    {
        if (addressBindingSource != null) { await SaveSQLDatabaseAsync(true); }
        Close();
    }

    private async void AddressDGV_CellClick(object sender, DataGridViewCellEventArgs e)
    {
        // 1. Validitätsprüfung (Header-Klicks ausschließen)
        if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

        // 2. Prüfung auf Strg-Taste (WinForms-Standard)
        if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
        {
            var colName = addressDGV.Columns[e.ColumnIndex].Name;

            // Zeile im Grid selektieren
            addressDGV.Rows[e.RowIndex].Selected = true;

            // 3. Den Fokus-Diebstahl des Grids durch kurzes Nachgeben verhindern
            await Task.Yield();

            // 4. Das Control finden, das laut Dictionary diesem Spaltennamen zugeordnet ist
            // Wir suchen den Key (Control), dessen Value (string) dem Spaltennamen entspricht.
            var targetEntry = editControlsDictionary.FirstOrDefault(x =>
                string.Equals(x.Value, colName, StringComparison.OrdinalIgnoreCase));

            if (targetEntry.Key is Control targetControl)
            {
                targetControl.Focus();

                // Zusätzlicher Komfort für Textboxen
                if (targetControl is TextBoxBase tb)
                {
                    tb.SelectAll();
                }
                // Für ComboBoxen die Dropdown-Liste öffnen (optional)
                else if (targetControl is ComboBox cb)
                {
                    cb.DroppedDown = true;
                }
            }
        }
    }

    private void AddressBindingSource_ListChanged(object? sender, ListChangedEventArgs e) => UpdateSaveButton();

    private void AddressBindingSource_CurrentChanged(object? sender, EventArgs e)
    {
        if (_isFiltering) { return; } // Konflikte mit Suchfilter vermeiden
        try
        {
            ignoreTextChange = true;
            if (addressBindingSource?.Current is Adresse currentAdresse)
            {
                ShowPhotoInPictureBoxy(currentAdresse);
                cbGrußformel.Items.Clear();
                ErzeugeGrußformeln();
                UpdateMembershipCBox();
                LoadGroupsForCurrentAddress();
                UpdateDocumentListView(currentAdresse);
                if (currentAdresse.Geburtstag.HasValue) { AgeLabel_MaskedTB_Set(currentAdresse.Geburtstag.Value); }
                else { AgeLabel_MaskedTB_Clear(); }
            }
            else
            {
                topAlignZoomPictureBox.Image = Properties.Resources.AddressBild100;
                delPictboxToolStripButton.Enabled = false;
                cbGrußformel.Items.Clear();
                flowLayoutPanel.Controls.Clear(); // Tags entfernen
                dokuListView.Items.Clear();
                AgeLabel_MaskedTB_Clear();
                tabPageDoku.ImageIndex = 3;
            }
            UpdatePlaceholderVis();
            LinkLabel_Enabled();
        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        finally { ignoreTextChange = false; }
    }

    private void UpdateDocumentListView(Adresse adresse) // Wird von AddressBindingSource_CurrentChanged aufgerufen
    {
        dokuListView.Items.Clear();

        if (adresse.Dokumente != null && adresse.Dokumente.Count > 0)
        {
            dokuListView.BeginUpdate(); // Performance bei vielen Dokus
            foreach (var dok in adresse.Dokumente)
            {
                if (!string.IsNullOrWhiteSpace(dok.Dateipfad)) { Add2dokuListView(new FileInfo(dok.Dateipfad), false); }
            }
            dokuListView.ListViewItemSorter = new ListViewItemComparer();
            dokuListView.Sort();
            dokuListView.EndUpdate();
        }
        tabPageDoku.ImageIndex = dokuListView.Items.Count > 0 ? 4 : 3;  // Icon des Tabs aktualisieren (Index 4 = voll, 3 = leer)
    }

    private void LoadGroupsForCurrentAddress()
    {
        curAddressMemberships.Clear();
        if (addressBindingSource.Current is Adresse adresse)
        {
            foreach (var gruppe in adresse.Gruppen) { curAddressMemberships.Add(gruppe.Name); } // EF Core hat die Gruppen (hoffentlich via .Include) geladen
        }
        UpdateMembershipTags(); // UI aktualisieren
    }

    private Image? LadeFotoFuerAddress(Adresse currentAdresse)
    {
        if (currentAdresse == null) { return null; }
        if (_context != null)
        {
            _context.Entry(currentAdresse).Reference(a => a.Foto).Load();
            if (currentAdresse.Foto != null && currentAdresse.Foto.Fotodaten != null)
            {
                var fotoBytes = currentAdresse.Foto.Fotodaten;
                using var ms = new MemoryStream(fotoBytes);
                return Image.FromStream(ms);
            }
        }
        return null; // Gibt null zurück, wenn Adresse oder Foto nicht gefunden wurden
    }

    private void AgeLabel_MaskedTB_Set(DateOnly date)
    {
        maskedTextBox.Text = date.ToString("dd.MM.yyyy");
        DateTime dateAsDateTime = new(date.Year, date.Month, date.Day);
        var todayAsDateTime = DateTime.Today;
        var days = (todayAsDateTime - dateAsDateTime).Days;
        if (Math.Abs(days) <= 31) { ageLabel.Text = Math.Abs(days).Equals(1) ? days.ToString() + " Tag" : days.ToString() + " Tage"; }
        else
        {
            var ddf = Utils.CalcDateDiff(todayAsDateTime, dateAsDateTime);
            ageLabel.Text = (!ddf.years.Equals(0) ? ddf.years.ToString() + (ddf.years.Equals(1) ? " Jahr" : " Jahre") +
                (ddf.months.Equals(0) && ddf.days.Equals(0) ? "" : ", ") : "") + (!ddf.months.Equals(0) ? ddf.months.ToString() +
                (ddf.months.Equals(1) ? " Monat" : " Monate") + (ddf.days.Equals(0) ? "" : ", ") : "") +
                (!ddf.days.Equals(0) ? ddf.days.ToString() + (ddf.days.Equals(1) ? " Tag" : " Tage") : "");

            toolTip.SetToolTip(ageLabel, $"{days} Tage");
        }
    }

    private void AgeLabel_MaskedTB_Clear()
    {
        maskedTextBox.Mask = "";
        maskedTextBox.Text = "";
        ageLabel.Text = string.Empty;
        toolTip.SetToolTip(ageLabel, string.Empty);
    }

    private void AddressDGV_DataSourceChanged(object sender, EventArgs e)
    {
        if (addressDGV.DataSource != null && hideColumnArr.Length == addressDGV.Columns.Count)
        {
            for (var i = 0; i < addressDGV.Columns.Count; i++) { addressDGV.Columns[i].Visible = !hideColumnArr[i]; }
            Text = appName + " – " + (string?)(string.IsNullOrEmpty(_databaseFilePath) ? "unbenannt" : Utils.CorrectUNC(_databaseFilePath));  // Workaround for UNC-Path
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

    private void SearchTSTextBox_TextChanged(object sender, EventArgs e)
    {
        if (!searchTSTextBox.Focused || ignoreSearchChange) { return; } // Nur reagieren, wenn der User tippt
        tsClearLabel.Visible = searchTSTextBox.TextBox.Text.Length > 0;  // "X"-Button Logik
        searchTimer.Stop();  // Laufenden Timer abbrechen
        searchTimer.Start();
    }

    private void ApplyGlobalSearch(string searchText)
    {
        var term = searchText.Trim().ToLower();
        var isSearchEmpty = string.IsNullOrWhiteSpace(term);
        _isFiltering = true;

        // Bestimme aktive BindingSource und Grid
        BindingSource? activeBs = null;
        DataGridView? activeDGV = null;

        if (tabControl.SelectedTab == addressTabPage)
        {
            activeBs = addressBindingSource;
            activeDGV = addressDGV;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            activeBs = contactBindingSource;
            activeDGV = contactDGV;
        }

        if (activeBs == null || activeDGV == null)
        {
            _isFiltering = false;
            return;
        }

        // WICHTIG 1: Laufende Editierung beenden, sonst entstehen Geister-Zeilen
        activeBs.EndEdit();

        // WICHTIG 2: Während gesucht wird, darf der User keine neuen Zeilen anlegen
        // Das verhindert das "Aufploppen" leerer Zeilen beim Backspace
        activeDGV.AllowUserToAddRows = isSearchEmpty;

        var currencyManager = BindingContext?[activeBs] as CurrencyManager;
        currencyManager?.SuspendBinding();

        try
        {
            // --- FALL A: SQL ADRESSEN ---
            if (tabControl.SelectedTab == addressTabPage && _context != null)
            {
                if (isSearchEmpty)  // Reset: Alle lokalen Daten anzeigen
                {
                    addressBindingSource.DataSource = _context.Adressen.Local.ToBindingList();
                    filterRemoveToolStripMenuItem.Visible = false;
                }
                else
                {
                    var filteredList = _context.Adressen.Local.Where(a => a.SearchText.Contains(term)).ToList();
                    addressBindingSource.DataSource = filteredList;
                    filterRemoveToolStripMenuItem.Visible = true;
                }
                UpdateAddressStatusBar();
            }
            // --- FALL B: GOOGLE KONTAKTE ---
            else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
            {
                if (isSearchEmpty)
                {
                    contactBindingSource.DataSource = _allGoogleContacts;
                    filterRemoveToolStripMenuItem.Visible = false;
                }
                else
                {
                    var filteredList = _allGoogleContacts.Where(c => c.SearchText.Contains(term)).ToList();
                    contactBindingSource.DataSource = filteredList;
                    filterRemoveToolStripMenuItem.Visible = true;
                }
                UpdateContactStatusBar();
            }
        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        finally
        {
            currencyManager?.ResumeBinding();
            _isFiltering = false;
            if (tabControl.SelectedTab == contactTabPage && !isSearchEmpty) // Snapshot zurücksetzen nur bei Google Kontakten
            {
                _lastActiveContact = contactBindingSource.Current as Contact;
                _originalContactSnapshot = _lastActiveContact != null ? (Contact)_lastActiveContact.Clone() : null;
            }
        }
    }

    private void UpdateAddressStatusBar()
    {
        if (_context == null) { return; }
        var totalCount = _context.Adressen.Local.Count;
        var visibleCount = addressBindingSource.Count;
        toolStripStatusLabel.Text = visibleCount == totalCount ? $"{totalCount} Adressen" : $"{visibleCount}/{totalCount} Adressen";
        if (visibleCount > 0 && addressDGV.Rows.Count > 0)
        {
            addressDGV.ClearSelection();
            addressDGV.Rows[0].Selected = true;
        }
    }

    private void UpdateContactStatusBar()
    {
        if (_allGoogleContacts == null) { return; }
        var total = _allGoogleContacts.Count;
        var visible = contactBindingSource.Count;
        toolStripStatusLabel.Text = visible == total ? $"{total} Google Kontakte" : $"{visible}/{total} Google Kontakte";
    }


    private async void SaveTSButton_Click(object sender, EventArgs e)
    {
        // --- FALL A: SQLITE ADRESSEN ---
        if (tabControl.SelectedTab == addressTabPage && addressBindingSource?.Current is Adresse aktuelleAdresse)
        {
            await SaveSQLDatabaseAsync(false, true);
            var gespeicherteId = aktuelleAdresse.Id; // Die ID merken, um nach der Sortierung dorthin zurückzukehren
            if (_context != null)
            {
                var sortierteListe = _context.Adressen.Local.OrderBy(a => a.Nachname).ThenBy(a => a.Vorname).ToList();
                addressBindingSource.RaiseListChangedEvents = false;
                addressBindingSource.DataSource = sortierteListe;
                addressBindingSource.RaiseListChangedEvents = true;
                addressBindingSource.ResetBindings(false);
            }

            // 3. Fokus wiederherstellen über die ID
            var newIndex = addressBindingSource.Find("Id", gespeicherteId);
            if (newIndex >= 0)
            {
                addressBindingSource.Position = newIndex;
                addressDGV.FirstDisplayedScrollingRowIndex = Math.Max(0, newIndex);
            }

            saveTSButton.Enabled = false;
        }
        // --- FALL B: GOOGLE KONTAKTE ---
        else if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact googleKontakt)
        {
            if (string.IsNullOrEmpty(googleKontakt.ResourceName)) // Neuer Kontakt
            {
                if (CheckNewContactTidyUp()) { await CreateContactAsync(); }
            }
            else { await AskSaveContactChangesAsync(); } // Bestehender Kontakt 
            saveTSButton.Enabled = false;
        }
        else { Console.Beep(); }
    }

    private async Task AskSaveContactChangesAsync()
    {
        // Sicherheitschecks
        if (_originalContactSnapshot == null || _lastActiveContact == null) { return; }

        // WICHTIG: Erzwingt das Schreiben von pending Changes (z.B. aus der MaskedTextBox) ins Objekt
        contactBindingSource.EndEdit();

        // 2. Änderungen ermitteln
        var changedFields = _lastActiveContact.GetChangedFields(_originalContactSnapshot);

        // Foto ignorieren, da separat behandelt
        changedFields.Remove("photos");

        // Wenn nichts geändert wurde -> Abbruch
        if (changedFields.Count == 0) { return; }

        // --- Vorbereitung TaskDialog (Ihr Code) ---
        var initialButtonYes = new TaskDialogButton("Hochladen");
        var initialButtonNo = TaskDialogButton.Cancel; // Oder eigener Button "Änderungen verwerfen"

        // Text-Aufbereitung für den Dialog
        var fieldList = string.Join("\n", changedFields.Select(f => "• " + char.ToUpper(f[0]) + f[1..]));
        var shortSummary = $"{changedFields.Count} Bereich(e) wurden geändert.\n{fieldList}";

        // Annahme: Utils.GenerateDetailedDiff existiert bei Ihnen
        var detailedDiff = Utils.GenerateDetailedDiff(_lastActiveContact, _originalContactSnapshot);

        var nameParts = new[] { _lastActiveContact.Vorname, _lastActiveContact.Nachname }
            .Where(s => !string.IsNullOrWhiteSpace(s));
        var fullName = string.Join(" ", nameParts);
        var headingText = string.IsNullOrWhiteSpace(fullName)
            ? "Möchten Sie die Änderungen speichern?"
            : $"Möchten Sie die Änderungen an {fullName} speichern?";

        // --- Page Definition ---
        var initialPage = new TaskDialogPage()
        {
            Caption = "Google Kontakte",
            Heading = headingText,
            Text = shortSummary,
            Icon = TaskDialogIcon.ShieldBlueBar,
            AllowCancel = true,
            Buttons = { initialButtonNo, initialButtonYes },

            Expander = new TaskDialogExpander()
            {
                Text = detailedDiff,
                CollapsedButtonText = "Details anzeigen",
                ExpandedButtonText = "Details ausblenden",
                Position = TaskDialogExpanderPosition.AfterText
            }
        };

        var progressPage = new TaskDialogPage()
        {
            Caption = "Google Kontakte",
            Heading = "Bitte warten…",
            Text = "Änderungen werden hochgeladen.",
            Icon = TaskDialogIcon.Information,
            ProgressBar = new TaskDialogProgressBar() { State = TaskDialogProgressBarState.Marquee },
            Buttons = { TaskDialogButton.Close } // Button ist erst disabled
        };
        progressPage.Buttons[0].Enabled = false;

        // --- Navigation & Events ---

        // Verhindern, dass der Dialog beim Klick sofort schließt -> Navigation zur ProgressPage
        initialButtonYes.AllowCloseDialog = false;
        initialButtonYes.Click += (s, e) => initialPage.Navigate(progressPage);

        //// Bei "Abbrechen" oder "Nein"
        //initialButtonNo.Click += (s, e) =>
        //{
        //    // Optional: Änderungen verwerfen?
        //    // _originalContactSnapshot = null; 
        //};

        // Sobald die ProgressPage da ist, startet der Upload
        progressPage.Created += async (s, e) =>
        {
            try
            {
                // Der eigentliche API-Aufruf
                await UpdateGoogleContactAsync(_lastActiveContact, changedFields);

                // Kurz warten für UX (damit man "Fertig" sieht)
                await Task.Delay(500);

                // Erfolgreich -> Schließen erlauben und automatisch klicken oder User klicken lassen
                progressPage.Heading = "Erfolgreich gespeichert.";
                progressPage.ProgressBar.State = TaskDialogProgressBarState.Normal;
                progressPage.ProgressBar.Value = 100;

                progressPage.Buttons[0].Enabled = true;
                progressPage.Buttons[0].PerformClick(); // Dialog automatisch schließen

                // Cache resetten
                _lastActiveContact.ResetSearchCache();

                // Button deaktivieren (da jetzt gespeichert)
                saveTSButton.Enabled = false;
            }
            catch (Exception ex)
            {
                // Fehleranzeige im TaskDialog (oder separater Dialog)
                progressPage.Heading = "Fehler beim Speichern";
                progressPage.Text = ex.Message;
                progressPage.Icon = TaskDialogIcon.Error;
                progressPage.ProgressBar.State = TaskDialogProgressBarState.Error;
                progressPage.Buttons[0].Enabled = true; // User muss manuell schließen

                // Falls Utils.ErrTaskDlg existiert:
                // Utils.ErrTaskDlg(Handle, ex); 
            }
        };

        // Dialog anzeigen (blockiert den Code, bis geschlossen)
        TaskDialog.ShowDialog(Handle, initialPage);
    }

    private bool CheckNewContactTidyUp()
    {
        // Wir prüfen den aktuell aktiven Kontakt (unseren Snapshot-Helfer)
        if (_lastActiveContact == null || !string.IsNullOrEmpty(_lastActiveContact.ResourceName)) { return false; }

        // Prüfen, ob irgendein relevantes Feld ausgefüllt wurde
        var hasData = !string.IsNullOrWhiteSpace(_lastActiveContact.Vorname) ||
                       !string.IsNullOrWhiteSpace(_lastActiveContact.Nachname) ||
                       !string.IsNullOrWhiteSpace(_lastActiveContact.Firma) ||
                       !string.IsNullOrWhiteSpace(_lastActiveContact.Mail1);

        if (hasData)
        {
            return true; // Kontakt ist neu und hat Daten -> Speichern erlaubt
        }
        else
        {
            // Kontakt ist leer -> Aus der Liste entfernen
            if (_allGoogleContacts != null)
            {
                _allGoogleContacts.Remove(_lastActiveContact);
                contactBindingSource.Remove(_lastActiveContact);
            }
            _lastActiveContact = null;
            _originalContactSnapshot = null;
            return false;
        }
    }

    private void TbNotizen_SizeChanged(object sender, EventArgs e) => NativeMethods.ShowScrollBar(tbNotizen.Handle, 1, TextRenderer.MeasureText(tbNotizen.Text, tbNotizen.Font,
        new Size(tbNotizen.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height > tbNotizen.Height);

    private async void NewTSButton_Click(object sender, EventArgs e)
    {
        // Suche leeren, damit der neue Kontakt sichtbar ist
        if (!string.IsNullOrEmpty(searchTSTextBox.Text))
        {
            Clear_SearchTextBox();
        }

        if (tabControl.SelectedTab == contactTabPage)
        {
            // 1. Vorherigen Kontakt prüfen und ggf. speichern
            if (_lastActiveContact != null)
            {
                if (string.IsNullOrEmpty(_lastActiveContact.ResourceName))
                {
                    if (CheckNewContactTidyUp()) { await CreateContactAsync(); }
                }
                else { await AskSaveContactChangesAsync(); }
            }

            // 2. Neues Objekt erstellen
            var newContact = new Contact();

            // 3. In Listen einfügen
            _allGoogleContacts?.Add(newContact);
            var index = contactBindingSource.Add(newContact);

            // 4. Zum neuen Kontakt navigieren
            contactBindingSource.Position = index;

            // 5. UI Fokus setzen
            _lastActiveContact = newContact;
            _originalContactSnapshot = (Contact)newContact.Clone();

            saveTSButton.Enabled = true;
            cbAnrede.Focus();
        }
        else if (tabControl.SelectedTab == addressTabPage && addressBindingSource != null)
        {
            addressBindingSource.AddNew();
            // EF Core 10 erstellt hier automatisch eine neue 'Adresse'-Entity
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
                await AskSaveContactChangesAsync();
            }
            isSelectionChanging = true; // verhindert, dass ContactEditFields() aufgerufen wird 
            contactNewRowIndex = contactDGV.Rows.Add();
            contactDGV.Rows[contactNewRowIndex].Selected = true;
            contactDGV.FirstDisplayedScrollingRowIndex = contactNewRowIndex;
            saveTSButton.Enabled = true;
            cbAnrede.Focus();
            isSelectionChanging = false; // setzt isSelectionChanging zurück, damit ContactEditFields() wieder aufgerufen wird
        }
        else if (tabControl.SelectedTab == addressTabPage && addressBindingSource != null && _context != null)
        {
            if (addressBindingSource.Current is not Adresse originalAdresse)
            {
                Utils.MsgTaskDlg(Handle, "Hinweis", "Bitte wählen Sie zuerst eine Adresse zum Duplizieren aus.", TaskDialogIcon.Information);
                return;
            }
            try
            {
                var duplikat = _context.Adressen.Include(a => a.Foto).AsNoTracking().FirstOrDefault(a => a.Id == originalAdresse.Id);
                if (duplikat == null) { return; }
                duplikat.Id = 0;
                duplikat.Foto?.Id = 0;
                _context.Adressen.Add(duplikat);
                var newIndex = addressBindingSource.IndexOf(duplikat);
                if (newIndex == -1) // Nur wenn es NICHT automatisch erschienen ist, manuell hinzufügen
                {
                    addressBindingSource.Add(duplikat);
                    newIndex = addressBindingSource.IndexOf(duplikat);
                }
                if (newIndex >= 0)
                {
                    addressBindingSource.Position = newIndex;
                    addressDGV.FirstDisplayedScrollingRowIndex = newIndex;
                    addressDGV.Rows[newIndex].Selected = true;
                }
                saveTSButton.Enabled = true;
                cbAnrede.Focus();
            }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
        else { Console.Beep(); }
    }

    private async void CopyToOtherDGVMenuItem_Click(object sender, EventArgs e)
    {
        // FALL 1: Von Google-Kontakten (Contact) zu Lokaler Datenbank (Adresse)
        if (tabControl.SelectedTab == contactTabPage && contactDGV.CurrentRow?.DataBoundItem is Contact selectedGoogleContact)
        {
            // 1. Neue lokale Adresse erstellen
            var newLocalAddress = new Adresse();

            // 2. Felder automatisiert kopieren (via Reflection und dataFields)
            var contactType = typeof(Contact);
            var adresseType = typeof(Adresse);

            foreach (var field in dataFields)
            {
                var val = contactType.GetProperty(field)?.GetValue(selectedGoogleContact);
                adresseType.GetProperty(field)?.SetValue(newLocalAddress, val);
            }

            // 3. Foto übernehmen (falls vorhanden)
            if (!string.IsNullOrEmpty(selectedGoogleContact.PhotoUrl))
            {
                try
                {
                    var bytes = await HttpService.Client.GetByteArrayAsync(selectedGoogleContact.PhotoUrl);
                    newLocalAddress.Foto = new Foto { Fotodaten = bytes };
                }
                catch { }
            }
            addressBindingSource.Add(newLocalAddress);
            addressBindingSource.Position = addressBindingSource.Count - 1; //  triggert automatisch das CurrentChanged-Event und aktualisiert die Controls
            tabControl.SelectedTab = addressTabPage;
            searchTSTextBox.TextBox.Clear();
            addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.RowCount - 1;
            addressDGV.Rows[^1].Selected = true;
            cbAnrede.Focus();
            saveTSButton.Enabled = true;
        }
        // FALL 2: Von Lokaler Datenbank (Adresse) zu Google (Contact)
        else if (tabControl.SelectedTab == addressTabPage && addressDGV.CurrentRow?.DataBoundItem is Adresse selectedLocalAddress)
        {
            var newGoogleContact = new Contact();
            var contactType = typeof(Contact);
            var adresseType = typeof(Adresse);

            foreach (var field in dataFields)
            {
                var val = adresseType.GetProperty(field)?.GetValue(selectedLocalAddress);
                contactType.GetProperty(field)?.SetValue(newGoogleContact, val);
            }
            if (contactDGV.DataSource is BindingSource googleBs)
            {
                googleBs.Add(newGoogleContact);
                googleBs.Position = googleBs.Count - 1;
                tabControl.SelectedTab = contactTabPage;
                searchTSTextBox.TextBox.Clear();
                contactDGV.FirstDisplayedScrollingRowIndex = googleBs.Position;
                contactDGV.Rows[googleBs.Position].Selected = true;
            }

            cbAnrede.Focus();
            saveTSButton.Enabled = true;
        }
        else { Console.Beep(); }
    }

    //private async void DeleteTSButton_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact googleKontakt)
    //    {
    //        var (askBefore, deleteNow) = Utils.AskBeforeDeleteContact(Handle, googleKontakt, sAskBeforeDelete, false); // Tuple Destructuring
    //        sAskBeforeDelete = askBefore;
    //        if (!deleteNow) { return; }
    //        try
    //        {
    //            await DeleteGoogleContactAsync(googleKontakt);
    //            _allGoogleContacts?.Remove(googleKontakt);
    //            contactBindingSource.RemoveCurrent();
    //            UpdateContactStatusBar();
    //        }
    //        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
    //    }
    //    else if (tabControl.SelectedTab == addressTabPage && addressBindingSource.Current is Adresse adresseZumLoeschen && _context != null)
    //    {
    //        if (addressDGV.CurrentRow?.IsNewRow == true) { return; }
    //        bool deleteFinal;
    //        if (sAskBeforeDelete)
    //        {
    //            var (askBefore, deleteNow) = Utils.AskBeforeDeleteAddress(Handle, adresseZumLoeschen, sAskBeforeDelete);
    //            sAskBeforeDelete = askBefore;
    //            deleteFinal = deleteNow;
    //        }
    //        else { deleteFinal = true; }
    //        if (!deleteFinal) { return; }
    //        try
    //        {
    //            var aktuellerIndex = addressBindingSource.Position;
    //            _context.Adressen.Remove(adresseZumLoeschen);
    //            await _context.SaveChangesAsync();
    //            //addressBindingSource.RemoveCurrent();
    //            if (addressBindingSource.Count == 0)
    //            {
    //                ignoreSearchChange = true;
    //                searchTSTextBox.Text = string.Empty;
    //                ApplyGlobalSearch(string.Empty);
    //                ignoreSearchChange = false;
    //            }
    //            else { addressBindingSource.Position = Math.Min(aktuellerIndex, addressBindingSource.Count - 1); }
    //            saveTSButton.Enabled = false;
    //            UpdateAddressStatusBar();
    //        }
    //        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
    //    }
    //    else { Console.Beep(); }
    //}
    private async void DeleteTSButton_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact googleKontakt)
        {
            var (askBefore, deleteNow) = Utils.AskBeforeDeleteContact(Handle, googleKontakt, sAskBeforeDelete, false);
            sAskBeforeDelete = askBefore;
            if (!deleteNow) { return; }
            try
            {
                await DeleteGoogleContactAsync(googleKontakt);
                _allGoogleContacts?.Remove(googleKontakt);
                contactBindingSource.RemoveCurrent();
                UpdateContactStatusBar();
            }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
        else if (tabControl.SelectedTab == addressTabPage && addressBindingSource.Current is Adresse adresseZumLoeschen && _context != null)
        {
            // 1. Edit-Modus sauber beenden, um Seiteneffekte zu vermeiden
            addressDGV.EndEdit();
            addressBindingSource.EndEdit();

            if (addressDGV.CurrentRow?.IsNewRow == true) { return; }

            bool deleteFinal;
            if (sAskBeforeDelete)
            {
                var (askBefore, deleteNow) = Utils.AskBeforeDeleteAddress(Handle, adresseZumLoeschen, sAskBeforeDelete);
                sAskBeforeDelete = askBefore;
                deleteFinal = deleteNow;
            }
            else { deleteFinal = true; }

            if (!deleteFinal) { return; }

            try
            {
                var aktuellerIndex = addressBindingSource.Position;

                // Zugriff auf den EF-Eintrag
                var entry = _context.Entry(adresseZumLoeschen);

                // Prüfen: Ist es ein neuer, ungespeicherter Datensatz?
                var isNewRecord = entry.State == EntityState.Added || adresseZumLoeschen.Id == 0;

                if (isNewRecord)
                {
                    // FALL A: Nur im RAM -> "Vergessen" (Detachen)

                    // 1. Event aushängen! 
                    // Das verhindert, dass CurrentChanged feuert, während wir resetten.
                    // Damit wird der Aufruf von ClearAddressDetails() und der Crash unterbunden.
                    //addressBindingSource.CurrentChanged -= AddressBindingSource_CurrentChanged;
                    _isFiltering = true; // Verhindert AddressBindingSource_CurrentChanged während des Detachens

                    try
                    {
                        if (adresseZumLoeschen.Foto != null)
                        {
                            var fotoEntry = _context.Entry(adresseZumLoeschen.Foto);
                            if (fotoEntry.State == EntityState.Added || adresseZumLoeschen.Foto.Id == 0)
                            {
                                fotoEntry.State = EntityState.Detached;
                            }
                        }

                        entry.State = EntityState.Detached;

                        // Jetzt ist ResetBindings sicher, weil niemand zuhört
                        addressBindingSource.ResetBindings(false);
                    }
                    finally
                    {
                        // 2. Event wieder einhängen
                        //addressBindingSource.CurrentChanged += AddressBindingSource_CurrentChanged;
                        _isFiltering = false;
                    }

                    // 3. UI manuell und sicher aktualisieren
                    if (addressBindingSource.Count > 0)
                    {
                        // Wir springen auf einen gültigen Datensatz, das löst dann sauber CurrentChanged aus
                        addressBindingSource.MoveFirst();
                    }
                    else
                    {
                        // Liste ist leer -> UI leeren
                        // Da Current null ist, müssen wir aufpassen.
                        // Am besten nichts tun, oder falls nötig, ClearAddressDetails aufrufen,
                        // ABER nur wenn es Bindings verträgt.
                        // Oft reicht hier einfach gar nichts zu tun, da die Textboxen eh leer sind.
                    }
                }
                else
                {
                    // FALL B: Existiert in DB -> Löschen + SQL
                    _context.Adressen.Remove(adresseZumLoeschen);
                    await _context.SaveChangesAsync();

                    // Hier manuell aus der UI entfernen (falls nicht automatisch geschehen)
                    if (addressBindingSource.Count > 0)
                    {
                        addressBindingSource.RemoveCurrent();
                    }
                }

                // --- UI Aufräumen ---
                if (addressBindingSource.Count == 0)
                {
                    ignoreSearchChange = true;
                    searchTSTextBox.Text = string.Empty;
                    ApplyGlobalSearch(string.Empty);
                    ignoreSearchChange = false;
                }
                else
                {
                    // Position korrigieren
                    addressBindingSource.Position = Math.Min(aktuellerIndex, addressBindingSource.Count - 1);
                }

                // Button Status prüfen (sollte jetzt korrekt false sein)
                UpdateSaveButton();
                UpdateAddressStatusBar();
            }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
        else { Console.Beep(); }
    }

    private async void FrmAdressen_FormClosing(object sender, FormClosingEventArgs e)
    {
        if (_isClosing) { return; }
        AutoValidate = AutoValidate.Disable;
        e.Cancel = true; // Schließen erst einmal abbrechen, um Zeit für async Aufgaben zu gewinnen
        Enabled = false;  // UI sperren, um Mehrfachklicks zu verhindern
        Cursor = Cursors.WaitCursor;
        try
        {
            CleanupWordObjects();
            SaveConfiguration();
            if (contactBindingSource.Current is Contact currentContact)
            {
                if (CheckNewContactTidyUp()) { await CreateContactAsync(); }
                await AskSaveContactChangesAsync();
            }
            if (_context != null) { await SaveSQLDatabaseAsync(true, false, true); }
            addressBindingSource?.Dispose();
            contactBindingSource?.Dispose();
            searchTimer?.Dispose(); // Den neuen Timer nicht vergessen!
            debounceTimer?.Dispose();
            scrollTimer?.Dispose();
        }
        catch (Exception ex) { Debug.WriteLine($"Kritischer Fehler beim Beenden: {ex.Message}"); }
        finally
        {
            _isClosing = true;
            Cursor = Cursors.Default;
            Close(); // Löst das Event erneut aus, wird aber nun von Zeile 5 durchgelassen
        }
    }

    private void CleanupWordObjects()
    {
        try
        {
            if (wordDoc != null)
            {
                try { wordDoc.Close(false); } catch { }
                Marshal.ReleaseComObject(wordDoc);
                wordDoc = null;
            }

            if (wordApp != null)
            {
                try { wordApp.Quit(false); } catch { }
                Marshal.ReleaseComObject(wordApp);
                wordApp = null;
            }
        }
        finally
        {
            for (var i = 0; i < 2; i++) // Zweimaliges Aufrufen erzwingt das Abräumen von COM-Wrappern
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }

    private void AboutToolStripMenuItem_Click(object sender, EventArgs e) => Utils.HelpMsgTaskDlg(Handle, appLong, Icon);

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

    private async void ImportToolStripMenuItem_Click(object sender, EventArgs e)
    {
        var targetColumns = dataFields.ToList(); // Id, Dokument und Foto werden nicht importiert
        var allowedColumns = new HashSet<string>(targetColumns, StringComparer.OrdinalIgnoreCase);
        var btnCreateCSV = new TaskDialogButton("Beispiel-CSV erstellen");
        var btnImportCSV = new TaskDialogButton("Import starten…");
        var firstPage = new TaskDialogPage()
        {
            Caption = Application.ProductName,
            Heading = "CSV-Import vorbereiten",
            Text = $"Erwartete Spalten: {string.Join(", ", targetColumns)}\n\n" + "Die Spaltenreihenfolge ist beliebig. Gruppen sollten kommagetrennt angegeben werden.",
            Icon = TaskDialogIcon.Information,
            AllowCancel = true,
            Buttons = { btnCreateCSV, btnImportCSV }
        };
        var result = TaskDialog.ShowDialog(this, firstPage);
        if (result == btnCreateCSV)
        {
            CreateExampleCsv(targetColumns);
            return;
        }
        else if (result != btnImportCSV) { return; }
        openFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*";
        openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        if (openFileDialog.ShowDialog() != DialogResult.OK || string.IsNullOrEmpty(openFileDialog.FileName)) { return; }
        if (_context == null)
        {
            try
            {
                _databaseFilePath = Path.ChangeExtension(openFileDialog.FileName, ".adb");
                ConnectSQLDatabase(_databaseFilePath);
            }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); return; }
        }
        var lines = Utils.ReadAsLines(openFileDialog.FileName).ToList();
        if (lines.Count < 2) { return; }
        var headers = lines[0].Split(';');
        var unknownColumns = headers.Where(h => !string.IsNullOrEmpty(h) && !allowedColumns.Contains(h)).ToList();
        if (unknownColumns.Count != 0)
        {
            Utils.MsgTaskDlg(Handle, "Abbruch", $"Unbekannte Spalten in CSV: {string.Join(", ", unknownColumns)}");
            return;
        }
        if (addressBindingSource.Count > 0)
        {
            var (isYes, isNo, isCancelled) = Utils.YesNo_TaskDialog(this, appName, "Daten hinzufügen?", $"Möchten Sie in '{Path.GetFileName(_databaseFilePath)}' importieren?", "Importieren", "Abbrechen");
            if (!isYes) { return; } // Logik: Nur bei 'Yes' weitermachen. Bei 'No' oder 'Escape' (Cancelled) abbrechen.
            var headerMap = headers.Select((h, i) => new { Name = h, Index = i }).Where(x => !string.IsNullOrEmpty(x.Name)).ToDictionary(x => x.Name, x => x.Index);
            var importCount = 0;
            try
            {
                var currencyManager = BindingContext?[addressBindingSource] as CurrencyManager;
                currencyManager?.SuspendBinding();  // UI-Update pausieren
                foreach (var line in lines.Skip(1))
                {
                    var regex = new Regex(";(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                    var fields = regex.Split(line);
                    if (fields.Length < headers.Length) { continue; }
                    var neueAdresse = new Adresse();
                    foreach (var kvp in headerMap)
                    {
                        var val = fields[kvp.Value]?.Trim().Trim('"').Replace("\"\"", "\""); //  Value ist der Index in der CSV-Zeile; Doppelte Anführungszeichen im Text zu einem machen (CSV Standard)
                        if (string.IsNullOrEmpty(val)) { continue; }

                        // Fall A: Gruppen-Relation (M:N)
                        if (kvp.Key == "Gruppen") // Geändert von kvp.Name zu kvp.Key
                        {
                            var gruppenNamen = val.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                            foreach (var gName in gruppenNamen)
                            {
                                var gruppe = _context?.Gruppen.Local.FirstOrDefault(g => g.Name.Equals(gName, StringComparison.OrdinalIgnoreCase))
                                             ?? _context?.Gruppen.FirstOrDefault(g => g.Name.Equals(gName, StringComparison.CurrentCultureIgnoreCase));

                                if (gruppe == null)
                                {
                                    gruppe = new Gruppe { Name = gName };
                                    _context?.Gruppen.Add(gruppe);
                                }
                                neueAdresse.Gruppen.Add(gruppe);
                            }
                        }
                        else if (kvp.Key == "Geburtstag") // Geändert von kvp.Name zu kvp.Key
                        {
                            if (DateTime.TryParse(val, out var dt)) { neueAdresse.Geburtstag = DateOnly.FromDateTime(dt); }
                        }
                        else  // Standard-Textfelder via Reflection
                        {
                            var prop = typeof(Adresse).GetProperty(kvp.Key); // Geändert von kvp.Name zu kvp.Key
                            if (prop != null && prop.CanWrite) { prop.SetValue(neueAdresse, val); }
                        }
                    }
                    _context?.Adressen.Add(neueAdresse);
                    importCount++;
                }
                currencyManager?.ResumeBinding();
                addressBindingSource.ResetBindings(false);  // UI-Update erzwingen
                UpdateSaveButton(); // saveTSButton.Enabled = _context?.ChangeTracker.HasChanges() ?? false;
                if (addressBindingSource.Count > 0) { addressBindingSource.MoveLast(); }
                Utils.MsgTaskDlg(Handle, "Import erfolgreich", $"{importCount} Adressen wurden geladen.\nKlicken Sie auf 'Speichern', um die Änderungen in der Datenbank zu sichern.");
            }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
    }

    private void CreateExampleCsv(List<string> columns)
    {
        var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var filePath = Path.Combine(desktopPath, "adress_vorlage.csv");
        try
        {
            using var writer = new StreamWriter(filePath, false, Encoding.UTF8);
            writer.WriteLine(string.Join(";", columns));
            writer.WriteLine("Herr;;Mustermann;Max;;;;Musterfirma;Musterstraße 1;12345;Musterstadt;Deutschland;;;;12.05.1985;max@muster.de;;030123456;;0170123456;;;Notiztext;Freunde,Wichtig");

            Utils.MsgTaskDlg(Handle, "Vorlage erstellt", $"Die Datei 'adress_vorlage.csv' wurde auf Ihrem Desktop gespeichert.");
        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
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
                }
            }
            e.Handled = e.SuppressKeyPress = true;
        }
    }

    // Diese Methode darf 'async void' sein, da sie wie ein Event-Handler fungiert
    private async void HandleSwitchDatabaseAsync(string currentDbPath)
    {
        foreach (var file in recentFiles)
        {
            if (file == currentDbPath) { continue; }

            if (File.Exists(file))
            {
                // Hier können wir jetzt sauber warten!
                if (addressBindingSource != null) { await SaveSQLDatabaseAsync(true); }

                // Erst wenn das Speichern fertig ist, geht es hier weiter:
                ConnectSQLDatabase(file);

                // UI Updates
                ignoreSearchChange = true;
                searchTSTextBox.TextBox.Clear();
                ignoreSearchChange = false;
            }
            break; // Sobald eine Datei gefunden wurde, brechen wir ab
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
                foreach (var key in (string[])[.. bookmarkTextDictionary.Keys]) { bookmarkTextDictionary[key] = string.Empty; }
                Utils.WordInfoTaskDlg(Handle, [.. bookmarkTextDictionary.Keys], new(Resources.word32));
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
                Utils.StartFile(Handle, @"AdressenKontakte.pdf");
                return true;
            case Keys.I | Keys.Control:
                Utils.HelpMsgTaskDlg(Handle, appLong, Icon);
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
                BirthdayReminder(tabControl.SelectedTab == addressTabPage ? addressDGV : contactDGV);
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
                // Wir rufen die async-Methode auf (Fire & Forget)
                HandleSwitchDatabaseAsync(_databaseFilePath);
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

                if (tabControl.SelectedTab == addressTabPage && addressBindingSource != null) { RejectChangesToolStripMenuItem_Click(null!, null!); }
                else { Console.Beep(); }
                return true;
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
                    Utils.StartDir(Handle, Path.GetDirectoryName(_settingsPath) ?? string.Empty);
                    return true;
                }
            case Keys.F2 | Keys.Control | Keys.Shift:
                {
                    Utils.StartFile(Handle, _settingsPath);
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

    //private void MaskedTextBox_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
    //{
    //    if (e.KeyCode == Keys.OemPeriod)
    //    {
    //        var day = string.Empty;
    //        var month = string.Empty;
    //        var year = string.Empty;
    //        var dateComponents = maskedTextBox.Text.Split('.');
    //        if (dateComponents.Length > 0) { day = dateComponents[0].Trim(); }
    //        if (dateComponents.Length > 1) { month = dateComponents[1].Trim(); }
    //        if (dateComponents.Length > 2) { year = dateComponents[2].Trim(); }
    //        if (day.Length == 1) { day = "0" + day; }
    //        if (month.Length == 1) { month = "0" + month; }
    //        if (year.Length == 2) { year = "20" + year; }
    //        maskedTextBox.Text = day + "." + month + "." + year;
    //    }
    //}

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
        tbNotizen.BackColor = _isDarkMode ? Color.FromArgb(80, 80, 0) : Color.LightYellow;
        tbNotizen.ForeColor = _isDarkMode ? Color.White : Color.Black;
    }

    private void InternetLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo(tbInternet.Text) { UseShellExecute = true }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utils.ErrTaskDlg(Handle, ex); }
    }

    private void Mail1LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "mailto:" + tbMail1.Text }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utils.ErrTaskDlg(Handle, ex); }
    }

    private void Mail2LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "mailto:" + tbMail2.Text }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utils.ErrTaskDlg(Handle, ex); }
    }

    private void Tel1LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "tel:" + Regex.Replace(tbTelefon1.Text, cleanRegex, "") }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utils.ErrTaskDlg(Handle, ex); }
    }

    private void Tel2LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "tel:" + Regex.Replace(tbTelefon2.Text, cleanRegex, "") }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utils.ErrTaskDlg(Handle, ex); }
    }

    private void MobilLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        try { Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = "tel:" + Regex.Replace(tbMobil.Text, cleanRegex, "") }); }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { Utils.ErrTaskDlg(Handle, ex); }
    }

    private void WordTSButton_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count == 0)
        {
            Utils.MsgTaskDlg(Handle, "Es ist keine Adresse gewählt!", "Es gibt keine Daten zu übertragen.");
            return;
        }
        else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count == 0)
        {
            Utils.MsgTaskDlg(Handle, "Es ist kein Kontakt gewählt!", "Es gibt keine Daten zu übertragen.");
            return;
        }

        var isWordInstalled = !(Type.GetTypeFromProgID("Word.Application") == null);
        var isLibreOfficeInstalled = !(Type.GetTypeFromProgID("com.sun.star.ServiceManager") == null); // Utils.IsLibreOfficeInstalled();  
        //MessageBox.Show("isWordInstalled: " + isWordInstalled.ToString() + Environment.NewLine + "isLibreOfficeInstalled: " + isLibreOfficeInstalled.ToString(), "Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        if (sWordProcProg == true && !isWordInstalled)
        {
            Utils.MsgTaskDlg(Handle, "Word wurde nicht gefunden", "Installieren Sie Microsoft Word.");
            return;
        }
        else if (sWordProcProg == false && !isLibreOfficeInstalled)
        {
            Utils.MsgTaskDlg(Handle, "LibreOffice wurde nicht gefunden", "Installieren Sie LibreOffice Writer.");
            return;
        }
        else if (sWordProcProg == true) { WordProcess(); }
        else if (sWordProcProg == false) { LibreProcess(); }
        else if (sWordProcProg == null)
        {
            var result = Utils.AskWordProcessingProgram(Handle);
            if (result == true) { WordProcess(); }
            else if (result == false) { LibreProcess(); }
        }
    }

    private void LibreProcess()
    {
        FillDictionary();
        var helperPath = Path.Combine(Path.GetDirectoryName(appPath) ?? string.Empty, "LibreHelper", "LibreOffice.exe");
        var lastWriterNoDoc = NativeMethods.GetLastVisibleHandleByTitleEnd("LibreOffice"); // Process.GetProcessesByName("soffice.bin") findet immer nur einen Prozess!!
        if (!File.Exists(helperPath)) { Utils.MsgTaskDlg(Handle, @"LibreHelper\LibreOffice.exe nicht gefunden", helperPath, TaskDialogIcon.ShieldErrorRedBar); }
        else if (NativeMethods.GetLastVisibleHandleByTitleEnd("– LibreOffice Writer") != IntPtr.Zero) // geöffnentes Writer-Dokument
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = helperPath,
                Arguments = "\"" + JsonSerializer.Serialize(bookmarkTextDictionary).Replace("\"", "\\\"") + "\"",
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
                    else { Utils.MsgTaskDlg(Handle, "soffice.exe wurde nicht gefunden", exePath); }
                }
                else { Utils.MsgTaskDlg(Handle, "LibreOffice-Installationspfad nicht gefunden.", @"Computer\HKEY_LOCAL_MACHINE\SOFTWARE\LibreOffice\UNO\InstallPath"); }
            }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
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
                Utils.WordInfoTaskDlg(Handle, [.. bookmarkTextDictionary.Keys], new(Resources.word32));
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
                Utils.WordInfoTaskDlg(Handle, [.. bookmarkTextDictionary.Keys], new(Resources.word32));
                return;
            }
            foreach (var entry in bookmarkTextDictionary)
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
            string[] arrayOfAllKeys = [.. bookmarkTextDictionary.Keys];
        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); } //  + Environment.NewLine + ex.StackTrace
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
        bookmarkTextDictionary["Anrede"] = cbAnrede.Text;
        bookmarkTextDictionary["Praefix"] = cbPräfix.Text;
        bookmarkTextDictionary["Vorname"] = tbVorname.Text;
        bookmarkTextDictionary["Zwischenname"] = tbZwischenname.Text;
        bookmarkTextDictionary["Nickname"] = tbNickname.Text;
        bookmarkTextDictionary["Nachname"] = tbNachname.Text;
        bookmarkTextDictionary["Präfix_Zwischenname_Nachname"] = cbPräfix.Text + (cbPräfix.Text.Length > 0 ? " " : "") + tbZwischenname.Text + (tbZwischenname.Text.Length > 0 ? " " : "") + tbNachname.Text;
        bookmarkTextDictionary["Vorname_Zwischenname_Nachname"] = cbPräfix.Text + (cbPräfix.Text.Length > 0 ? " " : "") + tbZwischenname.Text + (tbZwischenname.Text.Length > 0 ? " " : "") + tbNachname.Text;
        bookmarkTextDictionary["Präfix_Vorname_Zwischenname_Nachname"] = cbPräfix.Text + (cbPräfix.Text.Length > 0 ? " " : "") + tbVorname.Text + (tbVorname.Text.Length > 0 ? " " : "") + tbZwischenname.Text + (tbZwischenname.Text.Length > 0 ? " " : "") + tbNachname.Text;
        bookmarkTextDictionary["Anrede_Präfix_Vorname_Zwischenname_Nachname"] = cbAnrede.Text + (cbAnrede.Text.Length > 0 ? " " : "") + cbPräfix.Text + (cbPräfix.Text.Length > 0 ? " " : "") + tbVorname.Text + (tbVorname.Text.Length > 0 ? " " : "") + tbZwischenname.Text + (tbZwischenname.Text.Length > 0 ? " " : "") + tbNachname.Text;
        bookmarkTextDictionary["Suffix"] = tbSuffix.Text;
        bookmarkTextDictionary["Firma"] = tbFirma.Text;
        bookmarkTextDictionary["StrasseNr"] = tbStraße.Text;
        bookmarkTextDictionary["PLZ"] = cbPLZ.Text;
        bookmarkTextDictionary["Ort"] = cbOrt.Text;
        bookmarkTextDictionary["PLZ_Ort"] = cbPLZ.Text + (cbPLZ.Text.Length > 0 ? " " : "") + cbOrt.Text;
        bookmarkTextDictionary["Land"] = cbLand.Text;
        bookmarkTextDictionary["Betreff"] = tbBetreff.Text;
        bookmarkTextDictionary["Grussformel"] = cbGrußformel.Text;
        bookmarkTextDictionary["Schlussformel"] = cbSchlussformel.Text;
        bookmarkTextDictionary["Mail1"] = tbMail1.Text;
        bookmarkTextDictionary["Mail2"] = tbMail2.Text;
        bookmarkTextDictionary["Telefon1"] = tbTelefon1.Text;
        bookmarkTextDictionary["Telefon2"] = tbTelefon2.Text;
        bookmarkTextDictionary["Mobil"] = tbMobil.Text;
        bookmarkTextDictionary["Fax"] = tbFax.Text;
        bookmarkTextDictionary["Internet"] = tbInternet.Text;
    }

    private void WordHelpToolStripMenuItem_Click(object sender, EventArgs e)
    {
        foreach (var key in (string[])[.. bookmarkTextDictionary.Keys]) { bookmarkTextDictionary[key] = string.Empty; }
        Utils.WordInfoTaskDlg(Handle, [.. bookmarkTextDictionary.Keys], new(Resources.word32));
    }

    private void StatusbarToolStripMenuItem_Click(object sender, EventArgs e) => statusStrip.Visible = statusbarToolStripMenuItem.Checked = !statusbarToolStripMenuItem.Checked;
    private void NewToolStripMenuItem_Click(object sender, EventArgs e) => NewTSButton_Click(sender, e);
    private void DuplicateToolStripMenuItem_Click(object sender, EventArgs e) => CopyTSButton_Click(sender, e);
    private void DeleteToolStripMenuItem_Click(object sender, EventArgs e) => DeleteTSButton_Click(sender, e);

    private void SwitchDataBinding(BindingSource targetSource)
    {
        var useNullConversion = (targetSource == addressBindingSource);  // Unterscheidung: Lokale DB (null erlaubt) vs. Google (leerer String bevorzugt)
        foreach (var (control, dataMember) in editControlsDictionary)
        {
            control.DataBindings.Clear();
            var textBinding = new Binding("Text", targetSource, dataMember, true, DataSourceUpdateMode.OnPropertyChanged) { NullValue = string.Empty };
            if (useNullConversion)  // Nur bei EF Core: Leeren String im UI wieder in echten Null-Wert in DB wandeln
            {
                textBinding.Parse += (s, e) => { if (e.Value is string str && string.IsNullOrEmpty(str)) { e.Value = null; } };
            }
            control.DataBindings.Add(textBinding);
        }

        maskedTextBox.DataBindings.Clear(); // Spezialfall: Geburtstag, spezielle Formatierung
        var birthdayBinding = new Binding("Text", targetSource, "Geburtstag", true, DataSourceUpdateMode.OnValidation);
        birthdayBinding.Format += (s, e) =>
        {
            if (e.Value is DateOnly d) { e.Value = d.ToString("dd.MM.yyyy"); }
            else { e.Value = ""; }
        };
        birthdayBinding.Parse += (s, e) =>
        {
            if (e.Value is string str) // DateOnly? kann nur ein gültiges Datum sein (z. B. 01.01.2000) oder null, deshalb kein useNullConversion-Check
            {
                var cleanStr = str.Replace("_", "").Trim();
                if (string.IsNullOrEmpty(cleanStr) || cleanStr == "..") { e.Value = null; }
                else if (DateOnly.TryParseExact(cleanStr, "dd.MM.yyyy", out var result)) { e.Value = result; }
            }
        };
        maskedTextBox.DataBindings.Add(birthdayBinding);
    }

    private async Task UpdateGoogleContactAsync(Contact contact, List<string> changedFields, Action? onClose = null)
    {
        try
        {
            // 1. Service holen
            var service = await Utils.GetPeopleServiceAsync(secretPath, tokenDir);

            // 2. Das Person-Objekt für das Update vorbereiten
            // Wir setzen ResourceName und ETag, damit Google weiß, wen wir meinen und ob die Version stimmt.
            var personToUpdate = new Person
            {
                ResourceName = contact.ResourceName,
                ETag = contact.ETag
            };

            // 3. Felder basierend auf 'changedFields' befüllen
            // Wir bauen nur die Teile des Objekts auf, die sich wirklich geändert haben.

            // --- A: Namen ---
            if (changedFields.Contains("names"))
            {
                personToUpdate.Names =
            [
                new() {
                    HonorificPrefix = contact.Praefix,
                    FamilyName = contact.Nachname,
                    GivenName = contact.Vorname,
                    MiddleName = contact.Zwischenname,
                    HonorificSuffix = contact.Suffix
                }
            ];
            }

            // --- B: Nicknames ---
            if (changedFields.Contains("nicknames"))
            {
                personToUpdate.Nicknames = [new Nickname { Value = contact.Nickname }];
            }

            // --- C: Adressen ---
            if (changedFields.Contains("addresses"))
            {
                personToUpdate.Addresses =
            [
                new() {
                    StreetAddress = contact.Strasse,
                    PostalCode = contact.PLZ,
                    City = contact.Ort,
                    Country = contact.Land
                    // Type = "home" // Optional, falls du Typen unterstützt
                }
            ];
            }

            // --- D: Organisation / Firma ---
            if (changedFields.Contains("organizations"))
            {
                personToUpdate.Organizations = [new Organization { Name = contact.Firma }];
            }

            // --- E: Geburtstag ---
            if (changedFields.Contains("birthdays") && contact.Geburtstag.HasValue)
            {
                personToUpdate.Birthdays =
            [
                new() {
                    Date = new Date
                    {
                        Day = contact.Geburtstag.Value.Day,
                        Month = contact.Geburtstag.Value.Month,
                        Year = contact.Geburtstag.Value.Year
                    }
                }
            ];
            }

            // --- F: Emails (Flattened -> List) ---
            if (changedFields.Contains("emailAddresses"))
            {
                personToUpdate.EmailAddresses = [];
                if (!string.IsNullOrWhiteSpace(contact.Mail1))
                {
                    personToUpdate.EmailAddresses.Add(new EmailAddress { Value = contact.Mail1, Type = "home" });
                }

                if (!string.IsNullOrWhiteSpace(contact.Mail2))
                {
                    personToUpdate.EmailAddresses.Add(new EmailAddress { Value = contact.Mail2, Type = "work" });
                }
            }

            // --- G: Telefone (Flattened -> List) ---
            if (changedFields.Contains("phoneNumbers"))
            {
                personToUpdate.PhoneNumbers = [];
                if (!string.IsNullOrWhiteSpace(contact.Telefon1))
                {
                    personToUpdate.PhoneNumbers.Add(new PhoneNumber { Value = contact.Telefon1, Type = "home" });
                }

                if (!string.IsNullOrWhiteSpace(contact.Telefon2))
                {
                    personToUpdate.PhoneNumbers.Add(new PhoneNumber { Value = contact.Telefon2, Type = "work" });
                }

                if (!string.IsNullOrWhiteSpace(contact.Mobil))
                {
                    personToUpdate.PhoneNumbers.Add(new PhoneNumber { Value = contact.Mobil, Type = "mobile" });
                }

                if (!string.IsNullOrWhiteSpace(contact.Fax))
                {
                    personToUpdate.PhoneNumbers.Add(new PhoneNumber { Value = contact.Fax, Type = "fax" });
                }
            }

            // --- H: URLs ---
            if (changedFields.Contains("urls"))
            {
                personToUpdate.Urls = [new Url { Value = contact.Internet, Type = "homePage" }];
            }

            // --- I: Notizen ---
            if (changedFields.Contains("biographies"))
            {
                personToUpdate.Biographies = [new Biography { Value = contact.Notizen }];
            }

            // --- J: User Defined Fields ---
            if (changedFields.Contains("userDefined"))
            {
                personToUpdate.UserDefined = [];
                if (!string.IsNullOrWhiteSpace(contact.Anrede))
                {
                    personToUpdate.UserDefined.Add(new UserDefined { Key = "Anrede", Value = contact.Anrede });
                }

                if (!string.IsNullOrWhiteSpace(contact.Betreff))
                {
                    personToUpdate.UserDefined.Add(new UserDefined { Key = "Betreff", Value = contact.Betreff });
                }

                if (!string.IsNullOrWhiteSpace(contact.Grussformel))
                {
                    personToUpdate.UserDefined.Add(new UserDefined { Key = "Grussformel", Value = contact.Grussformel });
                }

                if (!string.IsNullOrWhiteSpace(contact.Schlussformel))
                {
                    personToUpdate.UserDefined.Add(new UserDefined { Key = "Schlussformel", Value = contact.Schlussformel });
                }
            }

            // --- K: Gruppen (Memberships) ---
            // Dies ist komplexer, da wir evtl. neue Gruppen anlegen müssen
            HashSet<string> groupsToRemoveToCheck = []; // Für die "Leere Gruppen löschen"-Logik

            if (changedFields.Contains("memberships"))
            {
                personToUpdate.Memberships = [];
                var desiredGroupNames = new HashSet<string>(contact.GroupNames, StringComparer.OrdinalIgnoreCase);

                // Sicherstellen, dass "Starred" korrekt gemappt wird (Google intern vs. UI)
                if (desiredGroupNames.Remove("★"))
                {
                    desiredGroupNames.Add("starred");
                }

                // Systemgruppe "My Contacts" sollte meist erhalten bleiben
                desiredGroupNames.Add("myContacts"); // Oder Logik prüfen, ob der User das entfernt hat

                foreach (var groupName in desiredGroupNames)
                {
                    string resourceName;
                    // 1. Versuchen, ResourceName aus Dictionary zu holen
                    var existingEntry = contactGroupsDict.FirstOrDefault(x => x.Value.Equals(groupName, StringComparison.OrdinalIgnoreCase));

                    if (!string.IsNullOrEmpty(existingEntry.Key))
                    {
                        resourceName = existingEntry.Key;
                    }
                    else if (groupName == "myContacts" || groupName == "starred")
                    {
                        resourceName = "contactGroups/" + groupName;
                    }
                    else
                    {
                        // 2. Gruppe existiert noch nicht -> Neu anlegen
                        var createdResourceName = await CreateContactGroupAsync(service, groupName);
                        if (string.IsNullOrEmpty(createdResourceName)) { continue; } // Fehler beim Erstellen
                        contactGroupsDict[createdResourceName] = groupName;

                        resourceName = createdResourceName;



                        //resourceName = await CreateContactGroupAsync(service, groupName);
                        //if (string.IsNullOrEmpty(resourceName)) continue; // Fehler beim Erstellen

                        //// Dictionary aktualisieren, damit wir es beim nächsten Mal wissen
                        //contactGroupsDict[resourceName] = groupName;
                    }

                    personToUpdate.Memberships.Add(new Membership
                    {
                        ContactGroupMembership = new ContactGroupMembership { ContactGroupResourceName = resourceName }
                    });
                }

                // Bestimmen, welche Gruppen verlassen wurden (für die Lösch-Prüfung später)
                if (_originalContactSnapshot != null)
                {
                    var originalGroups = _originalContactSnapshot.GroupNames
                        .Select(g => g == "★" ? "starred" : g) // Mapping beachten
                        .ToHashSet();

                    // Alle Gruppen, die im Original waren, aber jetzt nicht mehr da sind
                    var removedGroups = originalGroups.Except(desiredGroupNames);

                    foreach (var rem in removedGroups)
                    {
                        var resKey = contactGroupsDict.FirstOrDefault(x => x.Value.Equals(rem, StringComparison.OrdinalIgnoreCase)).Key;
                        if (!string.IsNullOrEmpty(resKey))
                        {
                            groupsToRemoveToCheck.Add(resKey);
                        }
                    }
                }
            }

            // 4. Request senden
            if (changedFields.Count > 0)
            {
                var updateRequest = service.People.UpdateContact(personToUpdate, contact.ResourceName);
                updateRequest.UpdatePersonFields = Utils.BuildMask([.. changedFields]);

                var updatedPerson = await updateRequest.ExecuteAsync();

                // WICHTIG: ETag im lokalen Objekt aktualisieren!
                // Sonst gibt es beim nächsten Speichern einen Concurrency-Fehler.
                contact.ETag = updatedPerson.ETag;
                contact.ResourceName = updatedPerson.ResourceName; // Zur Sicherheit

                // 5. Leere Gruppen prüfen (Legacy Logik)
                if (groupsToRemoveToCheck.Count > 0)
                {
                    onClose?.Invoke(); // Dialog schließen, bevor wir evtl. Fragen stellen
                    await CheckAndDeleteEmptyGroups(service, groupsToRemoveToCheck);
                }
            }

            // UI Abschluss
            onClose?.Invoke();
            saveTSButton.Enabled = false;
        }
        catch (Google.GoogleApiException gEx)
        {
            onClose?.Invoke();
            if (gEx.HttpStatusCode == System.Net.HttpStatusCode.PreconditionFailed)
            {
                Utils.MsgTaskDlg(Handle, "Konflikt beim Speichern",
                    "Der Kontakt wurde zwischenzeitlich an anderer Stelle geändert. Bitte laden Sie die Kontakte neu.");
            }
            else
            {
                Utils.ErrTaskDlg(Handle, gEx);
            }
        }
        catch (Exception ex)
        {
            onClose?.Invoke();
            Utils.ErrTaskDlg(Handle, ex);
        }
    }

    private async Task CheckAndDeleteEmptyGroups(PeopleServiceService service, HashSet<string> groupResourceNames)
    {
        foreach (var resourceName in groupResourceNames)
        {
            try
            {
                if (resourceName.Contains("starred") || resourceName.Contains("myContacts")) { continue; }

                var request = service.ContactGroups.Get(resourceName);
                request.MaxMembers = 1; // Wir wollen nur wissen, ob > 0
                var group = await request.ExecuteAsync();

                if (group.MemberResourceNames == null || group.MemberResourceNames.Count == 0)
                {
                    contactGroupsDict.TryGetValue(resourceName, out var groupName);

                    var (isYes, isNo, isCancelled) = Utils.YesNo_TaskDialog(this, "Google Kontakte",
                        heading: $"Gruppe '{groupName}' löschen?",
                        text: "Die Gruppe ist nun leer und hat keine Mitglieder mehr."); // Dein Icon
                    if (isYes)
                    {
                        await service.ContactGroups.Delete(resourceName).ExecuteAsync();
                        contactGroupsDict.Remove(resourceName); // Auch aus dem lokalen Cache entfernen
                    }
                }
            }
            catch { }
        }
    }


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

                var service = await Utils.GetPeopleServiceAsync(secretPath, tokenDir);

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
                if (!string.IsNullOrEmpty(cbGrußformel.Text.Trim())) { userdefined.Add(new UserDefined() { Key = "Grussformel", Value = cbGrußformel.Text }); }
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
                        var peopleService = await Utils.GetPeopleServiceAsync(secretPath, tokenDir);
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
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
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
            Utils.ErrTaskDlg(Handle, ex);
            return null; // Gibt null zurück, wenn die Erstellung fehlschlägt.
        }
    }

    private async Task DeleteGoogleContactAsync(Contact contact)
    {
        if (contact == null || string.IsNullOrEmpty(contact.ResourceName)) { return; }

        try
        {
            toolStripProgressBar.Style = ProgressBarStyle.Marquee;
            toolStripProgressBar.Visible = true;

            // 1. Service holen
            var service = await Utils.GetPeopleServiceAsync(secretPath, tokenDir);

            // 2. Den Löschbefehl direkt mit dem ResourceName des Objekts ausführen
            // Das ist sicher gegen Sortierung und Filterung!
            var request = service.People.DeleteContact(contact.ResourceName);
            await request.ExecuteAsync();

            // Optional: Cache-Bereinigung, falls nötig
            contact.ResetSearchCache();
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            throw; // Den Fehler weiterwerfen, damit der Aufrufer (DeleteTSButton_Click) ihn bemerkt
        }
        finally
        {
            toolStripProgressBar.Visible = false;
            toolStripProgressBar.Style = ProgressBarStyle.Blocks;
        }
    }

    private async void ShowPhotoInPictureBoxy(object item)
    {
        // 1. Reset
        topAlignZoomPictureBox.Image = tabControl.SelectedTab == contactTabPage ? Resources.ContactBild100 : Resources.AddressBild100;
        delPictboxToolStripButton.Enabled = false;

        if (item is IContactEntity entity)
        {
            try
            {
                // --- NEU: EF Core "Explicit Loading" für SQL-Adressen ---
                // Wenn es eine SQL-Adresse ist und der Context verfügbar ist...
                if (item is Adresse adresse && _context != null)
                {
                    // Prüfen, ob EF Core das Foto für diesen Eintrag schon geladen hat.
                    // IsLoaded ist false, wenn wir das Include beim Start weggelassen haben.
                    var entry = _context.Entry(adresse);
                    if (!entry.Reference(a => a.Foto).IsLoaded)
                    {
                        // Jetzt erst das Foto aus der DB holen (nur für diesen einen Kontakt!)
                        await entry.Reference(a => a.Foto).LoadAsync();
                    }
                }
                // ---------------------------------------------------------

                // 3. Jetzt wie gewohnt das Bild abrufen (Lokal oder Web)
                var image = await entity.GetPhotoAsync();

                // 4. Prüfen, ob der User schon weitergeklickt hat (Race-Condition verhindern)
                var currentBindingSource = tabControl.SelectedTab == addressTabPage
                    ? addressBindingSource
                    : contactBindingSource;

                if (currentBindingSource.Current != item)
                {
                    image?.Dispose();
                    return;
                }

                // 5. Anzeigen
                if (image != null)
                {
                    topAlignZoomPictureBox.Image = image;
                    delPictboxToolStripButton.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Fehler beim Laden des Fotos: " + ex.Message);
            }
        }
    }

    private async Task UpdateContactPhotoAsync(Contact contact, Image imageToUpload, ImageFormat formatToUse, Action onClose)
    {
        try
        {
            var service = await Utils.GetPeopleServiceAsync(secretPath, tokenDir);
            byte[] photoBytes;

            // .NET 10 / C# 14: Sauberer Umgang mit dem Bitmap-Klon
            // Wir nutzen den Klon, um das Original-Image-Objekt nicht zu sperren
            using (var clonedImage = new Bitmap(imageToUpload))
            {
                using var ms = new MemoryStream();
                clonedImage.Save(ms, formatToUse);
                photoBytes = ms.ToArray();
            }

            var base64Photo = Convert.ToBase64String(photoBytes);
            var updatePhotoRequest = new UpdateContactPhotoRequest
            {
                PhotoBytes = base64Photo,
                // Wichtig: PersonFields gibt an, was im Antwort-Objekt enthalten sein soll
                PersonFields = "photos"
            };

            // Wir nutzen contact.ResourceName direkt vom Objekt
            var request = service.People.UpdateContactPhoto(updatePhotoRequest, contact.ResourceName);
            var response = await request.ExecuteAsync();

            // Extraktion der neuen URL aus der API-Antwort
            var newUrl = response?.Person?.Photos?.FirstOrDefault()?.Url;

            if (!string.IsNullOrEmpty(newUrl))
            {
                // 1. Wert im Objekt im RAM aktualisieren
                contact.PhotoUrl = newUrl;

                // 2. Cache zurücksetzen, damit beim nächsten Anzeigen das neue Bild geladen wird
                contact.ResetSearchCache();

                // 3. Grid-Zelle aktualisieren (Sicherer Weg über die BindingSource)
                // Wir suchen die Position des Objekts in der Liste, egal wo es im Grid gerade steht
                var index = contactBindingSource.IndexOf(contact);
                if (index >= 0)
                {
                    // Dies triggert ein automatisches UI-Update der betroffenen Zeile
                    contactBindingSource.ResetItem(index);
                }
            }

            onClose?.Invoke();
        }
        catch (Exception ex)
        {
            onClose?.Invoke();
            Utils.ErrTaskDlg(Handle, ex);
        }
    }

    private async Task DeleteContactPhotoAsync(Contact contact)
    {
        // Sicherheitscheck
        if (contact == null || string.IsNullOrEmpty(contact.ResourceName)) { return; }

        try
        {
            var service = await Utils.GetPeopleServiceAsync(secretPath, tokenDir);

            // Wir nutzen den ResourceName direkt vom Objekt
            var request = service.People.DeleteContactPhoto(contact.ResourceName);
            request.PersonFields = "photos";

            var response = await request.ExecuteAsync();

            // 1. Objekt direkt aktualisieren
            // Wenn Google das Foto löscht, ist die PhotoUrl im Response null oder weg
            var photo = response?.Person?.Photos?.FirstOrDefault();
            contact.PhotoUrl = photo?.Url;

            // 2. Cache zurücksetzen (wichtig für die "Alle mit Bild" Filter-Logik)
            contact.ResetSearchCache();

            // 3. Anzeige aktualisieren
            ShowPhotoInPictureBoxy(contact);
        }
        catch (Google.GoogleApiException gex) when (gex.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
        {
            Utils.MsgTaskDlg(Handle, "Kein Foto vorhanden",
                "Es konnte online kein Foto zum Löschen gefunden werden.", TaskDialogIcon.Information);

            contact.PhotoUrl = null;
            ShowPhotoInPictureBoxy(contact);
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
        }
    }

    private async Task LoadAndDisplayGoogleContactsAsync()
    {
        if (tabControl.SelectedTab == addressTabPage && addressBindingSource != null)
        {
            if (filterRemoveToolStripMenuItem.Visible) { FilterRemoveToolStripMenuItem_Click(null!, EventArgs.Empty); }
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
            await AskSaveContactChangesAsync(); //if (CheckAndSaveContactChangesAsync()) { ShowMultiPageTaskDialog(); }
            lastContactSearch = searchTSTextBox.TextBox.Text;
            ignoreSearchChange = true;
            searchTSTextBox.TextBox.Clear();
            ignoreSearchChange = false;
        }
        if (!Utils.GoogleConnectionCheck(Handle, secretPath)) { return; }  // Bricht die Methode ab, wenn keine Verbindung besteht
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
            BindingList<Contact> contactList = [];
            if (response?.Connections == null || !response.Connections.Any())
            {
                toolStripStatusLabel.Text = "Keine Kontakte gefunden.";
                contactDGV.Rows.Clear();
                contactDGV.Columns.Clear();
                return;
            }
            //ContactEditFields(-1);
            //contactDGV.Rows.Clear();
            //contactDGV.Columns.Clear();
            var people = response.Connections;
            foreach (var person in people)
            {
                var newContact = new Contact
                {
                    ResourceName = person.ResourceName,
                    ETag = person.ETag,

                    // Namen
                    Praefix = person.Names?.FirstOrDefault()?.HonorificPrefix ?? "",
                    Nachname = person.Names?.FirstOrDefault()?.FamilyName ?? "",
                    Vorname = person.Names?.FirstOrDefault()?.GivenName ?? "",
                    Zwischenname = person.Names?.FirstOrDefault()?.MiddleName ?? "",
                    Nickname = person.Nicknames?.FirstOrDefault()?.Value ?? "",
                    Suffix = person.Names?.FirstOrDefault()?.HonorificSuffix ?? "",

                    // Firma
                    Firma = person.Organizations?.FirstOrDefault()?.Name ?? "",

                    // Adresse
                    Strasse = person.Addresses?.FirstOrDefault()?.StreetAddress ?? "",
                    PLZ = person.Addresses?.FirstOrDefault()?.PostalCode ?? "",
                    Ort = person.Addresses?.FirstOrDefault()?.City ?? "",
                    Land = person.Addresses?.FirstOrDefault()?.Country ?? "",

                    // Notizen / Bio
                    Notizen = person.Biographies?.FirstOrDefault()?.Value.ReplaceLineEndings() ?? "",

                    // Internet
                    Internet = person.Urls?.FirstOrDefault()?.Value ?? "",

                    // Emails
                    Mail1 = person.EmailAddresses?.FirstOrDefault()?.Value ?? "",
                    Mail2 = (person.EmailAddresses?.Count > 1) ? person.EmailAddresses[1].Value : "",

                    // Telefone (Hilfsmethode Utils.GetGooglePhoneByType weiter nutzen)
                    Telefon1 = Utils.GetGooglePhoneByType(person, "home") ?? "",
                    Telefon2 = Utils.GetGooglePhoneByType(person, "work") ?? "",
                    Mobil = Utils.GetGooglePhoneByType(person, "mobile") ?? "",
                    Fax = Utils.GetGooglePhoneByType(person, "fax") ?? ""
                };

                // Custom Fields (Anrede, Betreff, etc.)
                if (person.UserDefined != null)
                {
                    foreach (var customField in person.UserDefined)
                    {
                        if (customField.Key == "Anrede") { newContact.Anrede = customField.Value; }
                        else if (customField.Key == "Betreff") { newContact.Betreff = customField.Value; }
                        else if (customField.Key == "Grussformel") { newContact.Grussformel = customField.Value; }
                        else if (customField.Key == "Schlussformel") { newContact.Schlussformel = customField.Value; }
                    }
                }

                // Geburtstag
                if (person.Birthdays != null && person.Birthdays.Count > 0 && person.Birthdays[0].Date != null)
                {
                    var bday = person.Birthdays[0].Date;
                    // Achtung: Jahr kann null sein bei Google Contacts! Hier einfacher Fallback.
                    try
                    {
                        newContact.Geburtstag = new DateOnly(bday.Year ?? 1900, bday.Month ?? 1, bday.Day ?? 1);

                    }
                    catch { /* Ungültiges Datum abfangen */ }
                }

                // Foto URL
                if (person.Photos != null)
                {
                    var photo = person.Photos.FirstOrDefault(p => !string.IsNullOrEmpty(p.Url));
                    // Logik: Wenn nicht Default, dann ist es ein echtes User-Foto
                    if (photo != null && (!photo.Default__ ?? true))
                    {
                        newContact.PhotoUrl = photo.Url;
                    }
                }

                // Gruppen Logik
                var groupNames = new HashSet<string>();
                if (person.Memberships != null)
                {
                    foreach (var membership in person.Memberships)
                    {
                        if (membership.ContactGroupMembership?.ContactGroupResourceName != null &&
                            contactGroupsDict.TryGetValue(membership.ContactGroupMembership.ContactGroupResourceName, out var groupName))
                        {
                            if (!excludedGroups.Contains(groupName))
                            {
                                groupName = groupName.Equals("starred") ? "★" : groupName;
                                groupNames.Add(groupName);
                            }
                        }
                    }
                }
                newContact.GroupNames = [.. groupNames];
                contactList.Add(newContact);
            }
            _allGoogleContacts = contactList;
            allContactMemberships.Add("★");
            toolStripStatusLabel.Text = contactList.Count.ToString() + " Kontakte";
            //    response.Connections.Clear();  // dispose people 
            contactBindingSource.DataSource = contactList;
            contactDGV.DataSource = contactBindingSource;
            SwitchDataBinding(contactBindingSource);

            Utils.ApplyColumnSettings(contactDGV, _settings.ColumnWidths, hideColumnStd);
            tabControl.SelectedIndex = 1;
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
            cbAnrede.Items.AddRange([.. contactList.Select(c => c.Anrede ?? string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()]);
            cbPräfix.Items.AddRange([.. contactList.Select(c => c.Praefix ?? string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()]);
            cbPLZ.Items.AddRange([.. contactList.Select(c => c.PLZ ?? string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()]);
            cbOrt.Items.AddRange([.. contactList.Select(c => c.Ort ?? string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()]);
            cbLand.Items.AddRange([.. contactList.Select(c => c.Land ?? string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()]);
            cbSchlussformel.Items.AddRange([.. contactList.Select(c => c.Schlussformel ?? string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()]);
            contactCbItems_Anrede = [.. cbAnrede.Items.Cast<string>()];
            contactCbItems_Präfix = [.. cbPräfix.Items.Cast<string>()];
            contactCbItems_PLZ = [.. cbPLZ.Items.Cast<string>()];
            contactCbItems_Ort = [.. cbOrt.Items.Cast<string>()];
            contactCbItems_Land = [.. cbLand.Items.Cast<string>()];
            contactCbItems_Schlussformel = [.. cbSchlussformel.Items.Cast<string>()];
            if (contactBirthdayFlag && birthdayContactShow)
            {
                toolStripProgressBar.Visible = false;
                BirthdayReminder(contactDGV);
            }
            contactBirthdayFlag = true;
            Utils.StartSearchCacheWarmup(_allGoogleContacts);
        }
        catch (TokenResponseException)
        {
            contactBirthdayFlag = false;
            Utils.MsgTaskDlg(Handle, "Autorisierung erforderlich",
            "Das Zugriffstoken ist abgelaufen oder ungültig.\nDer Google-OAuth-Dialog wird beim nächsten Versuch erneut im Browser aufgerufen,\ndort können Sie den Zugriff auf Ihre Kontakte erlauben.",
            TaskDialogIcon.Information);
        }
        catch (Google.GoogleApiException ex) { Utils.ErrTaskDlg(Handle, ex); }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
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
        if (isSelectionChanging || _isFiltering) { return; }
        scrollTimer.Start();
        isSelectionChanging = true;
        try
        {
            // Wir arbeiten direkt mit der BindingSource, das ist sicherer als SelectedRows[0]
            if (contactBindingSource.Current is Contact selectedContact)
            {
                // 1. UI-Elemente aktivieren
                btnEditContact.Visible = true;

                // 2. Snapshot für die Änderungsverfolgung erstellen
                // Dies ist der "reine" Zustand, bevor der User tippt.
                _originalContactSnapshot = (Contact)selectedContact.Clone();

                // 3. Den aktuell aktiven Kontakt für andere Methoden (wie Auto-Save) merken
                _lastActiveContact = selectedContact;

                // 4. UI-Details aktualisieren (Foto, Gruppen, etc.)
                // Wir nutzen die Methode, die wir für das Interface vereinheitlicht haben.
                ShowPhotoInPictureBoxy(selectedContact);
                UpdateMembershipTags();

                // 5. Save-Buttons initial deaktivieren (da noch nichts geändert wurde)
                saveTSButton.Enabled = false;
            }
            else
            {
                btnEditContact.Visible = false;
                _originalContactSnapshot = null;
                _lastActiveContact = null;
            }
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
        }
        finally
        {
            isSelectionChanging = false;
        }
    }

    private async void ContactDGV_CellClick(object sender, DataGridViewCellEventArgs e)
    {
        // 1. Validitätsprüfung (keine Header, keine ungültigen Klicks)
        if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

        // 2. Prüfung auf Strg-Taste via WinForms ModifierKeys
        if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
        {
            // Spaltenname aus dem Google-Grid holen
            var colName = contactDGV.Columns[e.ColumnIndex].Name;

            // Zeile markieren
            contactDGV.Rows[e.RowIndex].Selected = true;

            // 3. UI-Thread kurz freigeben, damit der Standard-Zellfokus verarbeitet wird
            await Task.Yield();

            // 4. Reverse Lookup im Dictionary: Suche das Control zum Spaltennamen
            var targetEntry = editControlsDictionary.FirstOrDefault(x =>
                string.Equals(x.Value, colName, StringComparison.OrdinalIgnoreCase));

            if (targetEntry.Key is Control targetControl)
            {
                // Fokus auf das entsprechende Eingabefeld setzen
                targetControl.Focus();

                // Komfort-Funktionen für die Eingabe
                if (targetControl is TextBoxBase tb)
                {
                    tb.SelectAll();
                }
                else if (targetControl is ComboBox cb)
                {
                    cb.DroppedDown = true;
                }
            }
        }
    }

    private async void ContactBindingSource_CurrentChanged(object sender, EventArgs e)
    {
        if (_isFiltering) { return; }
        if (contactBindingSource.Current is not Contact contact)
        {
            _originalContactSnapshot = null;
            _lastActiveContact = null;
            topAlignZoomPictureBox.Image = Resources.ContactBild100;
            delPictboxToolStripButton.Enabled = false;
            AgeLabel_MaskedTB_Clear();
            flowLayoutPanel.Controls.Clear();
            return;
        }

        try
        {
            // --- NEU: Snapshot für den AKTUELLEN Kontakt einrasten ---
            _lastActiveContact = contact;
            _originalContactSnapshot = (Contact)contact.Clone();

            ignoreTextChange = true;
            ShowPhotoInPictureBoxy(contact); // Foto Logik (Vereinheitlicht) 
            cbGrußformel.Items.Clear();
            ErzeugeGrußformeln();

            // --- D: Geburtstag & Alter ---
            if (contact.Geburtstag.HasValue)
            {
                AgeLabel_MaskedTB_Set(contact.Geburtstag.Value);
            }
            else
            {
                AgeLabel_MaskedTB_Clear();
            }

            // --- E: Gruppen / Tags ---
            curContactMemberships = new SortedSet<string>(contact.GroupNames ?? [], StringComparer.OrdinalIgnoreCase);
            if (curContactMemberships.Count > 0)
            {
                allContactMemberships.UnionWith(curContactMemberships);
                UpdateMembershipTags();
            }
            else
            {
                flowLayoutPanel.Controls.Clear();
                UpdatePlaceholderVis();
            }
            UpdateMembershipCBox();
            LinkLabel_Enabled();
            btnEditContact.Visible = true;

        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        finally { ignoreTextChange = false; }
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
        // 1. Google Kontakte laden, wenn der Tab gewechselt wird und leer ist
        if (e.TabPage == contactTabPage && (contactBindingSource.DataSource == null || contactBindingSource.Count == 0))
        {
            var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Google Kontakte", "Keine Kontakte vorhanden", "Möchten Sie Ihre Kontakte jetzt laden?");
            if (isYes) { await LoadAndDisplayGoogleContactsAsync(); }
        }

        // 2. Prüfen auf ungespeicherte Änderungen beim VERLASSEN des Google-Tabs
        // (e.TabPage ist der NEUE Tab, also prüfen wir, ob wir von Google kommen)
        else if (e.TabPage == addressTabPage && contactBindingSource.Current is Contact lastContact)
        {
            // Fall A: Ein neuer Kontakt wurde angefangen (ResourceName ist noch leer)
            if (string.IsNullOrEmpty(lastContact.ResourceName) && CheckNewContactTidyUp())
            {
                // Wir müssen den Tab-Wechsel kurz stoppen, um zu speichern
                e.Cancel = true;
                await CreateContactAsync();
                // Nach dem Speichern manuell den Tab wechseln
                tabControl.SelectedTab = addressTabPage;
                return;
            }

            // Fall B: Bestehender Kontakt wurde geändert
            if (ContactChanges_Check())
            {
                // Da CheckAndSaveContactChangesAsync asynchron ist und Dialoge anzeigt,
                // brechen wir den automatischen Wechsel ab und führen ihn manuell nach dem Dialog aus.
                e.Cancel = true;
                await AskSaveContactChangesAsync();

                // Wenn der User im Dialog gespeichert oder verworfen hat, wechseln wir nun:
                if (!ContactChanges_Check())
                {
                    tabControl.SelectedTab = addressTabPage;
                }
            }
        }

        // 3. Filter beim Tab-Wechsel zurücksetzen
        if (filterRemoveToolStripMenuItem.Visible)
        {
            FilterRemoveToolStripMenuItem_Click(null!, null!);
        }
    }

    private async void TabControl_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage)
        {

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
            // Prüfen, ob die BindingSource Daten enthält (ersetzt _dataTable.Rows.Count)
            await AskSaveContactChangesAsync();
            _originalContactSnapshot = null;
            _lastActiveContact = null;

            if (addressBindingSource?.Count > 0)
            {
                SwitchDataBinding(addressBindingSource);
                // Titelzeile setzen
                Text = appName + " – " + (string.IsNullOrEmpty(_databaseFilePath) ? "unbenannt" : Utils.CorrectUNC(_databaseFilePath));

                btnEditContact.Visible = false;

                // Speichern-Button Status (ersetzt _dataTable.GetChanges)
                // EF Core prüft hier auf Added, Modified und Deleted Entitäten im ChangeTracker
                UpdateSaveButton(); // saveTSButton.Enabled = _context?.ChangeTracker.HasChanges() ?? false;

                // Toolbar-Buttons aktivieren
                newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled =
                deleteToolStripMenuItem.Enabled = deleteTSButton.Enabled =
                newTSButton.Enabled = copyTSButton.Enabled =
                wordTSButton.Enabled = envelopeTSButton.Enabled = true;

                copyToOtherDGVTSMenuItem.Enabled = false;

                // Statuszeile (Anzahl der Adressen)
                // Wir nutzen hier direkt die BindingSource, da diese den aktuellen (ggf. gefilterten) Stand kennt
                var rowCount = _context?.Adressen.Local.Count ?? 0; // Gesamtanzahl geladener Adressen
                var visibleRowCount = addressBindingSource.Count; // Anzahl aktuell angezeigter/gefilterter Adressen

                toolStripStatusLabel.Text = rowCount == visibleRowCount
                    ? $"{visibleRowCount} Adressen"
                    : $"{visibleRowCount}/{rowCount} Adressen";

                // Hinweis: SelectionChanged wird i.d.R. automatisch durch die BindingSource ausgelöst
            }
            Text = !string.IsNullOrWhiteSpace(_databaseFilePath) ? $"Adressen - {_databaseFilePath}" : "Adressen";
        }
        if (tabControl.SelectedTab == contactTabPage && contactDGV.RowCount > 1)
        {
            if (contactBindingSource.Current is Contact current)
            {
                _lastActiveContact = current;
                _originalContactSnapshot = (Contact)current.Clone();
            }

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
            if (contactDGV.RowCount > 0)
            {
                Text = !string.IsNullOrWhiteSpace(userEmail) ? $"Kontakte - {userEmail}" : "Google-Kontakte";
                btnEditContact.Visible = true;
                newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled
                    = duplicateToolStripMenuItem.Enabled = false;
                copyTSButton.Enabled = newTSButton.Enabled = deleteTSButton.Enabled = copyToOtherDGVTSMenuItem.Enabled = wordTSButton.Enabled = envelopeTSButton.Enabled = true;
                var rowCount = contactDGV.Rows.Count;
                var visibleRowCount = contactDGV.Rows.Cast<DataGridViewRow>().Count(static r => r.Visible);
                toolStripStatusLabel.Text = rowCount == visibleRowCount ? $"{visibleRowCount} Kontakte" : $"{visibleRowCount}/{rowCount} Kontakte";
            }
            SwitchDataBinding(contactBindingSource);
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
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
    }

    private void ExtraToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        authentMenuItem.Enabled = Directory.Exists(tokenDir);
        manageGroupsToolStripMenuItem.Enabled = tabControl.SelectedTab == contactTabPage ? contactDGV.Rows.Count > 0 : addressBindingSource != null;
    }

    private void BrowserPeopleMenuItem_Click(object sender, EventArgs e)
    {
        try
        {
            ProcessStartInfo psi = new("https://contacts.google.com/") { UseShellExecute = true };
            Process.Start(psi);
        }
        catch (Exception ex) when (ex is Win32Exception || ex is InvalidOperationException) { Utils.ErrTaskDlg(Handle, ex); }
    }

    private async void GoogleToolStripMenuItem_ClickAsync(object sender, EventArgs e) => await LoadAndDisplayGoogleContactsAsync();

    private void EnvelopeTSButton_Click(object sender, EventArgs e)
    {
        Cursor = Cursors.WaitCursor; // Aktiviert den Wartesymbol
        FillDictionary();
        using var frm = new FrmPrintSetting(sColorScheme, bookmarkTextDictionary,
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
        if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact contact)
        {
            var resourceId = contact.ResourceName.Split('/').LastOrDefault(); // "people/c123456789"
            if (!string.IsNullOrEmpty(resourceId))
            {
                try
                {
                    var url = $"https://contacts.google.com/person/{resourceId}";
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
            else { Utils.MsgTaskDlg(Handle, "Fehler", "Die Google-Ressourcen-ID konnte nicht ermittelt werden."); }
        }
        else { Console.Beep(); }
    }

    private void TsClearLabel_Click(object sender, EventArgs e) => Clear_SearchTextBox();

    private void TsClearLabel_VisibleChanged(object sender, EventArgs e) => searchTSTextBox.Width = 202 + splitContainer.SplitterDistance - 536 - (tsClearLabel.Visible ? tsClearLabel.Width : 0);

    private void TsClearLabel_Paint(object sender, PaintEventArgs e) => BeginInvoke(new Action(() => Graphics.FromHwnd(toolStrip.Handle).DrawRectangle(Pens.Black, tsClearLabel.Bounds.Location.X - 2, tsClearLabel.Bounds.Location.Y + 2, tsClearLabel.Width + 1, tsClearLabel.Height - 4)));
    // private void TsClearLabel_Paint(object sender, PaintEventArgs e) => InvokeAsync(() => Graphics.FromHwnd(toolStrip.Handle).DrawRectangle(Pens.Black, tsClearLabel.Bounds.Location.X - 2, tsClearLabel.Bounds.Location.Y + 2, tsClearLabel.Width + 1, tsClearLabel.Height - 4));

    //private void AddressDGV_KeyDown(object sender, KeyEventArgs e)
    //{
    //    var keyValue = e.KeyValue;
    //    if (e.Control && e.KeyCode == Keys.C)
    //    {
    //        ClipboardTSMenuItem_Click(null!, null!);
    //        e.Handled = true; // Prevent default copy behavior
    //    }
    //    else if (e.Modifiers == Keys.None && (keyValue >= (int)Keys.A && keyValue <= (int)Keys.Z || e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9))
    //    {
    //        searchTSTextBox.Focus();
    //        searchTSTextBox.Text += e.Shift ? ((char)keyValue).ToString() : ((char)(keyValue + 32)).ToString();
    //        searchTSTextBox.SelectionStart = searchTSTextBox.Text.Length;  // Cursor ans Ende stellen
    //    }
    //}

    private void AddressDGV_KeyDown(object sender, KeyEventArgs e)
    {
        var keyValue = e.KeyValue;
        if (e.Control && e.KeyCode == Keys.C)
        {
            ClipboardTSMenuItem_Click(null!, null!);
            e.Handled = true;
            e.SuppressKeyPress = true; // Auch hier sauber unterdrücken
            return;
        }
        else if (e.Modifiers == Keys.None && (keyValue >= (int)Keys.A && keyValue <= (int)Keys.Z || e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9))
        {
            searchTSTextBox.Focus();
            searchTSTextBox.Text += e.Shift ? ((char)keyValue).ToString() : ((char)(keyValue + 32)).ToString();
            searchTSTextBox.SelectionStart = searchTSTextBox.Text.Length;
            e.Handled = true;
            e.SuppressKeyPress = true; // Verhindert, dass das Grid versucht, zu einer Zeile zu springen, die mit dem Buchstaben beginnt
            return;
        }
        if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.PageUp || e.KeyCode == Keys.PageDown)
        {
            if (scrollTimer.Enabled)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }
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
        if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.PageUp || e.KeyCode == Keys.PageDown)
        {
            if (scrollTimer.Enabled)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
    }

    private void SearchTSTextBox_Enter(object sender, EventArgs e)
    {
        // Im Dark Mode nutzen wir ein dunkleres Gelb/Orange, damit weißer Text lesbar bleibt
        // Im Light Mode bleibt es bei deinem gewohnten LightYellow
        searchTSTextBox.BackColor = _isDarkMode ? Color.FromArgb(80, 80, 0) : Color.LightYellow;
        searchTSTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
    }

    private void SearchTSTextBox_Leave(object sender, EventArgs e)
    {
        searchTSTextBox.BackColor = _isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
        searchTSTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
    }

    private void ComboBox_Enter(object sender, EventArgs e)
    {
        if (sender is ComboBox cb)
        {
            cb.BackColor = _isDarkMode ? Color.FromArgb(80, 80, 0) : Color.LightYellow;
            cb.Invalidate(); // bei OwnerDraw ComboBoxen neuzeichnen
        }
    }

    private void ComboBox_Leave(object sender, EventArgs e)
    {
        if (sender is ComboBox cb)
        {
            cb.BackColor = _isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
            cb.Invalidate();
        }
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
            tb.SelectAll();
            // Dark Mode: Dunkles Gold/Gelb | Light Mode: LightYellow
            tb.BackColor = _isDarkMode ? Color.FromArgb(80, 80, 0) : Color.LightYellow;
            tb.ForeColor = _isDarkMode ? Color.White : Color.Black;
        }
        textBoxClicked = false;
    }

    private void TextBox_Leave(object sender, EventArgs e)
    {
        if (sender is TextBox tb)
        {
            tb.BackColor = _isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
            tb.ForeColor = _isDarkMode ? Color.White : Color.Black;
        }
    }

    private void MaskedTextBox_Enter(object sender, EventArgs e)
    {
        ignoreTextChange = true;
        maskedTextBox.Mask = @"00\.00\.0000";
        maskedTextBox.BackColor = _isDarkMode ? Color.FromArgb(80, 80, 0) : Color.LightYellow;
        maskedTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
        //maskedTextBox.BeginInvoke(new Action(() =>
        //{
        if (string.IsNullOrWhiteSpace(maskedTextBox.Text.Replace(".", "").Replace("_", "").Trim())) // falls leer, Cursor ganz links
        {
            maskedTextBox.SelectionStart = 0;
            maskedTextBox.SelectionLength = 0;
        }
        else { maskedTextBox.SelectAll(); } // falls schon was drin steht, alles markieren    
        //}));
        textBoxClicked = false;
        ignoreTextChange = false;
    }

    private void FormatAndSetDate()
    {
        var digits = new string([.. maskedTextBox.Text.Where(char.IsDigit)]); // nur die Ziffern behalten
        if (string.IsNullOrEmpty(digits)) // wenn nix da, alles löschen
        {
            maskedTextBox.Mask = "";
            maskedTextBox.Text = "";
            return;
        }
        string d = "01", m = "01", y = DateTime.Today.Year.ToString();
        switch (digits.Length)
        {
            case <= 2: // Nur Tag eingegeben (z.B. "5") -> 05.AktuellerMonat.AktuellesJahr
                d = digits.PadLeft(2, '0');
                m = DateTime.Today.Month.ToString("00");
                break;
            case 3:
            case 4: // Tag und Monat (z.B. "0512") -> 05.12.AktuellesJahr
                d = digits[..2];
                m = digits[2..].PadLeft(2, '0');
                break;
            case 5:
            case 6: // Tag, Monat, kurzes Jahr (z.B. "051224") -> 05.12.2024
                d = digits[..2];
                m = digits.Substring(2, 2);
                y = digits[4..];
                if (y.Length == 1) { y = "200" + y; }
                else if (y.Length == 2)
                {
                    var yearShort = int.Parse(y); // Logik: Ab 30 nehmen wir 19xx, sonst 20xx
                    y = (yearShort > 30 ? "19" : "20") + y;
                }
                break;
            case 8: // Komplettes Datum (z.B. "05121990")
                d = digits[..2];
                m = digits.Substring(2, 2);
                y = digits.Substring(4, 4);
                break;
        }
        ignoreTextChange = true;
        try
        {
            var finalDate = $"{d}.{m}.{y}";
            if (DateTime.TryParseExact(finalDate, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out _)) { maskedTextBox.Text = finalDate; }
            else
            {
                maskedTextBox.Mask = "";
                maskedTextBox.Text = "";
            }
            maskedTextBox.DataBindings["Text"]?.WriteValue(); // erzwingt BindingSource Update
        }
        finally { ignoreTextChange = false; }
    }

    private void MaskedTextBox_Leave(object sender, EventArgs e)
    {
        ignoreTextChange = true;
        try
        {
            var digits = new string([.. maskedTextBox.Text.Where(char.IsDigit)]); // nur die Ziffern behalten
            if (digits.Length == 0) // nichts eingegeben -> alles löschen
            {
                AgeLabel_MaskedTB_Clear();
            }
            else if (digits.Length > 0 && digits.Length < 8)
            {
                FormatAndSetDate();
                if (DateOnly.TryParseExact(maskedTextBox.Text, "dd.MM.yyyy", out var geburtsdatum)) { AgeLabel_MaskedTB_Set(geburtsdatum); }
                else
                {
                    AgeLabel_MaskedTB_Clear();
                }
            }
        }
        finally { ignoreTextChange = false; }
    }

    private void MaskedTextBox_KeyPress(object sender, KeyPressEventArgs e)
    {
        if (e.KeyChar == '.') // Wir unterdrücken den eigentlichen Punkt, da wir das Feld manuell formatieren
        {
            e.Handled = true;
            FormatAndSetDate();
        }
    }

    private void MaskedTextBox_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left) // !textBoxClicked  &&   
        {
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

    private void BtnResetDate_Click(object sender, EventArgs e)
    {
        ignoreTextChange = true;
        maskedTextBox.Mask = "";
        maskedTextBox.Clear();
        ignoreTextChange = false;
        maskedTextBox.Focus();
        UpdateSaveButton(); // Status aktualisieren, da das TextChanged-Event unterdrückt wurde
    }

    private void EditControlsFromDict_TextChanged(object sender, EventArgs e)
    {
        if (sender is not Control senderControl || !senderControl.Focused || ignoreTextChange || _isFiltering) { return; }
        var isLocal = tabControl.SelectedTab == addressTabPage;
        var isGoogle = tabControl.SelectedTab == contactTabPage;
        if (!isLocal && !isGoogle) { return; }
        if (isLocal) { senderControl.DataBindings["Text"]?.WriteValue(); } // Zwinge das Binding, den Wert SOFORT in das Entity zu schreiben
        if (isGoogle && contactBindingSource.Current is not Contact) { return; }
        if (ReferenceEquals(sender, tbNotizen))
        {
            var textSize = TextRenderer.MeasureText(tbNotizen.Text, tbNotizen.Font, new Size(tbNotizen.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);
            NativeMethods.ShowScrollBar(tbNotizen.Handle, 1, textSize.Height > tbNotizen.Height);
        }
        UpdateSaveButton();
    }


    private void MaskedTextBox_TextChanged(object sender, EventArgs e)
    {

        if (!maskedTextBox.Focused || ignoreTextChange) { return; }  // Guard Clauses

        maskedTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;


        if (!maskedTextBox.MaskFull) // Validierungslogik (Alter berechnen oder Label leeren)
        {
            var cleanText = maskedTextBox.Text.Replace(".", "").Replace("_", "").Trim();
            if (string.IsNullOrWhiteSpace(cleanText)) { AgeLabel_MaskedTB_Clear(); }
        }
        else
        {
            var rawText = maskedTextBox.Text; // Datum parsen und prüfen
            if (DateOnly.TryParseExact(rawText, formats, culture, DateTimeStyles.None, out var geburtsdatum))
            {
                if (geburtsdatum > DateOnly.FromDateTime(DateTime.Today)) { maskedTextBox.ForeColor = Color.Red; }
                else
                {
                    maskedTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
                    var heute = DateOnly.FromDateTime(DateTime.Today);
                    var alter = heute.Year - geburtsdatum.Year;
                    if (geburtsdatum > heute.AddYears(-alter)) { alter--; }
                    ageLabel.Text = $"Alter: {alter} Jahre";
                }
            }
            else // Ungültiges Datum
            {
                maskedTextBox.ForeColor = Color.Red;
                AgeLabel_MaskedTB_Clear();
            }
        }
        if (tabControl.SelectedTab == addressTabPage) { maskedTextBox.DataBindings["Text"]?.WriteValue(); }
        UpdateSaveButton();
    }

    private void OpenCalendar()
    {
        EnsureCalendar();
        if (Utils.TryParseInput(maskedTextBox.Text, out var current)) { monthCalendar!.SetDate(current); }
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

    private async void NewDBToolStripMenuItem_Click(object sender, EventArgs e)
    {
        try
        {
            saveFileDialog.Title = "Neue Datenbank anlegen";
            saveFileDialog.InitialDirectory = string.IsNullOrEmpty(sDatabaseFolder) || !Directory.Exists(sDatabaseFolder) ? null : sDatabaseFolder;
            saveFileDialog.DefaultExt = "adb";
            saveFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (addressBindingSource != null) { await SaveSQLDatabaseAsync(true); }
                _databaseFilePath = saveFileDialog.FileName;
            }
            else { return; }
            CreateNewDatabase(_databaseFilePath, true);
            ConnectSQLDatabase(_databaseFilePath);
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            _databaseFilePath = string.Empty;
        }
    }

    private void ExportToolStripMenuItem_Click(object sender, EventArgs e)
    {
        saveFileDialog.FileName = "Adressen_Export.csv";
        saveFileDialog.DefaultExt = "csv";
        saveFileDialog.Filter = "CSV-Datei (*.csv)|*.csv|Alle Dateien (*.*)|*.*";
        if (saveFileDialog.ShowDialog() != DialogResult.OK) { return; }
        if (addressBindingSource.Count > 0)
        {
            try
            {
                StringBuilder sb = new();
                var exportColumns = dataFields.Where(f => f != "Id").ToList(); // Header-Spaltennamen
                sb.AppendLine(string.Join(";", exportColumns));
                foreach (var item in addressBindingSource)
                {
                    if (item is Adresse adresse)
                    {
                        var fields = exportColumns.Select(columnName =>
                        {
                            object? value;
                            if (columnName == "Gruppen") { value = string.Join(", ", adresse.Gruppen.Select(g => g.Name)); }
                            else if (columnName == "Geburtstag") { value = adresse.Geburtstag?.ToShortDateString(); }
                            else { value = typeof(Adresse).GetProperty(columnName)?.GetValue(adresse); }
                            var fieldString = value?.ToString() ?? string.Empty;
                            return $"\"{fieldString.Replace("\"", "\"\"")}\"";
                        });
                        sb.AppendLine(string.Join(";", fields));
                    }
                }
                File.WriteAllText(saveFileDialog.FileName, sb.ToString(), Encoding.UTF8);
                Utils.MsgTaskDlg(Handle, "Export abgeschlossen", $"{addressBindingSource.Count} Datensätze wurden erfolgreich exportiert.");
            }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
    }

    private void ColumnSelectToolStripMenuItem_Click(object sender, EventArgs e)
    {
        using var frm = new FrmColumns(hideColumnStd);
        var itemCount = frm.GetColumnList().Items.Count;
        if (hideColumnArr.Length < itemCount) { Array.Resize(ref hideColumnArr, itemCount); }
        for (var i = 0; i < itemCount; i++) { frm.GetColumnList().Items[i].Checked = !hideColumnArr[i]; }
        if (frm.ShowDialog() == DialogResult.OK)
        {
            for (var i = 0; i < itemCount; i++)
            {
                var isVisible = frm.GetColumnList().Items[i].Checked;
                if (addressDGV.Columns.Count > i) { addressDGV.Columns[i].Visible = isVisible; }
                if (contactDGV.Columns.Count > i) { contactDGV.Columns[i].Visible = isVisible; }
                hideColumnArr[i] = !isVisible;
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
        using var frm = new FrmCopyScheme(sColorScheme, bookmarkTextDictionary, indexCopyPattern, copyPattern1 ?? [], copyPattern2 ?? [], copyPattern3 ?? [], copyPattern4 ?? [], copyPattern5 ?? [], copyPattern6 ?? []);
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
        for (var i = 0; i < patterns.Length; i++) { result[i] = string.Join(" ", Regex.Matches(patterns[i], @"\b\w+\b").Cast<Match>().Select(m => bookmarkTextDictionary.ContainsKey(m.Value) ? m.Value : string.Empty)).Trim(); }
        return result;
    }

    private void ContextMenu_Opening(object sender, CancelEventArgs e)
    {
        // 1. Grundsätzliche Prüfung: Ist überhaupt etwas ausgewählt?
        var isAddressTab = tabControl.SelectedTab == addressTabPage;
        var isContactTab = tabControl.SelectedTab == contactTabPage;

        // Wir nutzen die BindingSource.Current statt SelectedRows, da dies robuster ist
        if ((isAddressTab && addressBindingSource.Current == null) ||
            (isContactTab && contactBindingSource.Current == null))
        {
            e.Cancel = true;
            return;
        }

        // 2. Sichtbarkeit und Texte anpassen
        if (isAddressTab)
        {
            // Sicherstellen, dass die gewählte Zeile im Sichtfeld ist (UX-Verbesserung)
            if (addressDGV.CurrentRow != null && !Utils.RowIsVisible(addressDGV, addressDGV.CurrentRow))
            {
                addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.CurrentRow.Index;
            }

            copy2OtherDGVMenuItem.Text = "Zu Google-Kontakte hinzufügen";
            // Nur anzeigen, wenn Google-Kontakte grundsätzlich geladen wurden
            copy2OtherDGVMenuItem.Visible = _allGoogleContacts?.Count > 0;
            move2OtherDGVToolStripMenuItem.Visible = false;
        }
        else if (isContactTab)
        {
            if (contactDGV.CurrentRow != null && !Utils.RowIsVisible(contactDGV, contactDGV.CurrentRow))
            {
                contactDGV.FirstDisplayedScrollingRowIndex = contactDGV.CurrentRow.Index;
            }

            copy2OtherDGVMenuItem.Text = "In Lokale Adressen kopieren";
            // Immer möglich, sofern eine Datenbankverbindung besteht
            copy2OtherDGVMenuItem.Visible = _context != null;
            move2OtherDGVToolStripMenuItem.Visible = _context != null;
        }

        // Separator an die Sichtbarkeit des Kopier-Menüs koppeln
        copy2OtherDGVSeparator.Visible = copy2OtherDGVMenuItem.Visible;
    }

    private void NewTSMenuItem_Click(object sender, EventArgs e) => NewTSButton_Click(sender, e);
    private void DupTSMenuItem_Click(object sender, EventArgs e) => CopyTSButton_Click(sender, e);
    private void DelTSMenuItem_Click(object sender, EventArgs e) => DeleteTSButton_Click(sender, e);
    private void ClipTSMenuItem_Click(object sender, EventArgs e) => ClipboardTSMenuItem_Click(sender, e);
    private void Copy2OtherDGVMenuItem_Click(object sender, EventArgs e) => CopyToOtherDGVMenuItem_Click(sender, e);
    private void WordTSMenuItem_Click(object sender, EventArgs e) => WordTSButton_Click(sender, e);
    private void EnvelopeTSMenuItem_Click(object sender, EventArgs e) => EnvelopeTSButton_Click(sender, e);

    private void DGV_CellMouseDown_SelectRow(object sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0)
        {
            if (sender is DataGridView dgv)
            {
                if (!dgv.Rows[e.RowIndex].Selected)
                {
                    dgv.ClearSelection();
                    dgv.Rows[e.RowIndex].Selected = true;
                }
                dgv.CurrentCell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
            }
        }
    }

    private void MainToolStripMenuItem_DropDownOpened(object sender, EventArgs e) => ((ToolStripMenuItem)sender).ForeColor = SystemColors.ControlText;

    private void MainToolStripMenuItem_DropDownClosed(object sender, EventArgs e) => ((ToolStripMenuItem)sender).ForeColor = sColorScheme == "dark" ? SystemColors.HighlightText : SystemColors.ControlText;

    private void AddressDGV_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
    {
        if (e.RowIndex < 0) { return; }

        var dgv = (DataGridView)sender;

        // 1. Schärfere Schrift (Das behalten wir bei, da es sich auf das Graphics-Objekt auswirkt)
        e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

        // 2. Status prüfen
        var isSelected = (e.State & DataGridViewElementStates.Selected) == DataGridViewElementStates.Selected;

        // 3. Farben bestimmen
        Color backColor;
        Color foreColor; // Wichtig: Auch Textfarbe definieren, damit Selection nicht "unsichtbar" wird

        if (isSelected)
        {
            // Wir nehmen die definierten Selection-Farben
            backColor = addressDGV.DefaultCellStyle.SelectionBackColor;
            foreColor = addressDGV.DefaultCellStyle.SelectionForeColor;
        }
        else
        {
            // Deine Zebra-Logik
            var farbeEins = _isDarkMode ? Color.FromArgb(45, 42, 38) : Color.FloralWhite;
            var farbeZwei = _isDarkMode ? Color.FromArgb(32, 30, 28) : Color.White;
            backColor = (e.RowIndex % 2 == 0) ? farbeEins : farbeZwei;
            foreColor = addressDGV.DefaultCellStyle.ForeColor;
        }

        // 4. DER FIX: Wir manipulieren NICHT PaintParts und malen NICHT selbst.
        // Wir weisen dem Grid nur an, welche Farben es gleich selbst benutzen soll.
        // Das verhindert 100% der Ghosting-Effekte, da das Grid seinen internen "Clear"-Prozess sauber durchführt.

        // Zugriff auf die Row-Instanz, um den Style für diesen Paint-Zyklus zu setzen
        dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = backColor;
        dgv.Rows[e.RowIndex].DefaultCellStyle.SelectionBackColor = backColor; // Trick: Damit der blaue Standard-Balken nicht drüber gemalt wird
        dgv.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = foreColor;

        // 5. PaintHeader manuell ist nicht mehr nötig, das macht das System jetzt automatisch korrekt.
        // PaintParts müssen nicht mehr angefasst werden.
    }

    private void ContactDGV_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
    {
        if (e.RowIndex < 0) { return; }
        var dgv = (DataGridView)sender;
        e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        var isSelected = (e.State & DataGridViewElementStates.Selected) == DataGridViewElementStates.Selected;
        Color backColor;
        Color foreColor;
        if (isSelected)
        {
            backColor = contactDGV.DefaultCellStyle.SelectionBackColor;
            foreColor = contactDGV.DefaultCellStyle.SelectionForeColor;
        }
        else
        {
            var farbeEins = _isDarkMode ? Color.FromArgb(35, 38, 45) : Color.AliceBlue;
            var farbeZwei = _isDarkMode ? Color.FromArgb(28, 30, 35) : Color.White;
            backColor = (e.RowIndex % 2 == 0) ? farbeEins : farbeZwei;
            foreColor = contactDGV.DefaultCellStyle.ForeColor;
        }
        dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = backColor;
        dgv.Rows[e.RowIndex].DefaultCellStyle.SelectionBackColor = backColor;
        dgv.Rows[e.RowIndex].DefaultCellStyle.SelectionForeColor = foreColor;
    }

    private void RejectChangesToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (addressBindingSource.Current == null) { return; }
        if (addressBindingSource.Current is Contact currentContact)
        {
            if (_originalContactSnapshot == null) { return; }
            foreach (var propName in editControlsDictionary.Values.Distinct())
            {
                var propInfo = typeof(Contact).GetProperty(propName);
                if (propInfo != null && propInfo.CanWrite) { propInfo.SetValue(currentContact, propInfo.GetValue(_originalContactSnapshot)); }
            }
            currentContact.Geburtstag = _originalContactSnapshot.Geburtstag;
            currentContact.PhotoUrl = _originalContactSnapshot.PhotoUrl;
            currentContact.GroupNames.Clear();
            if (_originalContactSnapshot.GroupNames != null) { currentContact.GroupNames.AddRange(_originalContactSnapshot.GroupNames); } // Listen mit Inhalt ersetzen, nicht nur Referenz tauschen
            currentContact.ResetSearchCache(); // Namen könnten sich geändert haben 
        }
        else if (_context != null) // Lokale Adresse, EF Core Logik
        {
            var changedEntries = _context.ChangeTracker.Entries().Where(e => e.State != EntityState.Unchanged).ToList();
            foreach (var entry in changedEntries)
            {
                switch (entry.State)
                {
                    case EntityState.Added:
                        entry.State = EntityState.Detached;
                        break;
                    case EntityState.Modified:
                        entry.CurrentValues.SetValues(entry.OriginalValues);
                        entry.State = EntityState.Unchanged;
                        break;
                    case EntityState.Deleted:
                        entry.State = EntityState.Unchanged;
                        break;
                }
            }
        }
        addressBindingSource.ResetBindings(false);
        UpdateSaveButton();
    }

    private void EditToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            rejectChangesToolStripMenuItem.Enabled = _context?.ChangeTracker.HasChanges() ?? false;
            copyToOtherDGVTSMenuItem.Text = "Zu Google-&Kontakte hinzufügen";
            copyToOtherDGVTSMenuItem.Enabled = addressDGV.SelectedRows.Count > 0 && contactDGV.Rows.Count > 0;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            rejectChangesToolStripMenuItem.Enabled = ContactChanges_Check();
            copyToOtherDGVTSMenuItem.Text = "Nach Lokale Adressen &kopieren";
            copyToOtherDGVTSMenuItem.Enabled = contactDGV.SelectedRows.Count > 0 && addressDGV.Rows.Count > 0;
        }
    }

    private async void GooglebackupToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (contactDGV.Rows.Count == 0)
        {
            Utils.MsgTaskDlg(Handle, "Keine Daten zum Speichern", "Es sind keine Google-Kontaktdaten vohanden.");
            return;
        }
        saveFileDialog.Filter = "SQLite Database File (*.adb)|*.adb|All files (*.*)|*.*"; // using var sfd = new SaveFileDialog();
        saveFileDialog.Title = "Wählen Sie einen Speicherort";
        saveFileDialog.FileName = "GoogleKontakte.adb";
        saveFileDialog.InitialDirectory = Directory.Exists(sDatabaseFolder) ? sDatabaseFolder : Path.GetDirectoryName(_databaseFilePath);
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
                        await SaveGoogleContactsLocal(backupPath);
                        progressPage.Navigate(readyPage);
                    }
                    catch (Exception ex)
                    {
                        if (progressPage.BoundDialog != null) { progressPage.BoundDialog?.Close(); } // läuft im UI-Thread
                        var displayException = ex;
                        if (ex is AggregateException aggEx && aggEx.InnerExceptions.Count > 0) { displayException = aggEx.InnerExceptions[0]; }
                        Utils.MsgTaskDlg(Handle, displayException.GetType().Name, $"{displayException.Message}\nDer Backupvorgang wird abgebrochen!", TaskDialogIcon.ShieldWarningYellowBar);
                    }
                };
                if (TaskDialog.ShowDialog(Handle, progressPage) == TaskDialogButton.Yes)
                {
                    {
                        if (addressBindingSource != null) { await SaveSQLDatabaseAsync(true); }
                        ConnectSQLDatabase(backupPath);
                        ignoreSearchChange = true;
                        searchTSTextBox.TextBox.Clear();
                        ignoreSearchChange = false;
                        //if (birthdayAddressShow) { BirthdayReminder(); }
                    }
                }
            }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
    }

    private async Task SaveGoogleContactsLocal(string backupPath)
    {
        await Task.Run(() => CreateNewDatabase(backupPath, addSampleRecord: false));
        if (contactDGV.DataSource is not IEnumerable<Contact> googleContacts && contactDGV.DataSource is BindingSource bs && bs.DataSource is IEnumerable<Contact> list)
        {
            googleContacts = list;
        }
        else { return; }
        using var dbContext = new AdressenDbContext(backupPath);
        var groupCache = new Dictionary<string, Gruppe>(StringComparer.OrdinalIgnoreCase);
        var contactType = typeof(Contact);
        var adresseType = typeof(Adresse);
        foreach (var gContact in googleContacts)
        {
            var localAddress = new Adresse();
            foreach (var fieldName in dataFields)
            {
                var sourceProp = contactType.GetProperty(fieldName);
                var destProp = adresseType.GetProperty(fieldName);
                if (sourceProp != null && destProp != null && destProp.CanWrite)
                {
                    var value = sourceProp.GetValue(gContact);
                    destProp.SetValue(localAddress, value);
                }
            }
            if (!string.IsNullOrEmpty(gContact.PhotoUrl)) // GetPhotoAsync() gibt Image zurück, wir brauchen aber die Bytes
            {
                try
                {
                    var bytes = await HttpService.Client.GetByteArrayAsync(gContact.PhotoUrl);
                    if (bytes is { Length: > 0 })
                    {
                        localAddress.Foto = new Foto { Fotodaten = bytes };
                    }
                }
                catch { }
            }
            foreach (var groupName in gContact.GroupNames.Where(n => !string.IsNullOrWhiteSpace(n)))
            {
                if (!groupCache.TryGetValue(groupName, out var existingGroup))
                {
                    existingGroup = new Gruppe { Name = groupName };
                    groupCache[groupName] = existingGroup;
                }
                localAddress.Gruppen.Add(existingGroup);
            }

            dbContext.Adressen.Add(localAddress);
        }
        await dbContext.SaveChangesAsync();
    }

    private void ComboBox_DrawItem(object sender, DrawItemEventArgs e)
    {
        if (sender is not ComboBox comboBox || e.Index < 0) { return; }
        Color bgColor;
        Color textColor;
        if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
        {
            bgColor = _isDarkMode ? Color.FromArgb(176, 125, 71) : SystemColors.Highlight;
            textColor = Color.White;
        }
        else
        {
            bgColor = _isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
            textColor = _isDarkMode ? Color.White : Color.Black;
        }
        e.Graphics.FillRectangle(new SolidBrush(bgColor), e.Bounds); // Hintergrund malen
        e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        var itemText = comboBox.Items[e.Index]?.ToString() ?? string.Empty;
        TextRenderer.DrawText(e.Graphics, itemText, e.Font, e.Bounds, textColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
        if ((e.State & DrawItemState.Focus) == DrawItemState.Focus && !_isDarkMode) { e.DrawFocusRectangle(); } // im Dark Mode oft schöner ohne
    }

    private void BirthdaysToolStripMenuItem_Click(object sender, EventArgs e) => BirthdayReminder(tabControl.SelectedTab == addressTabPage ? addressDGV : contactDGV, true);

    private void BirthdayReminder(DataGridView dgv, bool showIfEmpty = false)
    {
        if (dgv.DataSource is not BindingSource bs) { return; }
        var isLocal = dgv == addressDGV;
        var autoShow = isLocal ? birthdayAddressShow : birthdayContactShow;
        if (isLocal && _context == null) { return; }

        if (!isLocal && _allGoogleContacts == null) { return; }
        IEnumerable<IContactEntity> source = isLocal ? _context!.Adressen.Local : _allGoogleContacts!;
        if (!source.Any() && !showIfEmpty) { return; }
        var heute = DateOnly.FromDateTime(DateTime.Today);

        var bevorstehendeGeburtstage = source.Where(x => x.BirthdayDate.HasValue).Select(x =>
        {
            var g = x.BirthdayDate!.Value;
            var day = g.Day; // Schaltjahr-Korrektur für die Berechnung "dieses Jahr"
            var month = g.Month;
            if (month == 2 && day == 29 && !DateTime.IsLeapYear(heute.Year)) { day = 28; }

            var gebTagDiesesJahr = new DateOnly(heute.Year, month, day);
            var tage = gebTagDiesesJahr.DayNumber - heute.DayNumber;
            if (tage < -birthdayRemindAfter) // Korrektur für Jahreswechsel
            {
                if (month == 2 && g.Day == 29 && !DateTime.IsLeapYear(heute.Year + 1)) { day = 28; } // Nächstes Jahr Schaltjahr Check
                else { day = g.Day; }
                tage = new DateOnly(heute.Year + 1, month, day).DayNumber - heute.DayNumber;
            }
            else if (tage > birthdayRemindLimit)
            {
                if (month == 2 && g.Day == 29 && !DateTime.IsLeapYear(heute.Year - 1)) { day = 28; } // Letztes Jahr Schaltjahr Check
                else { day = g.Day; }
                var tageLetztesJahr = new DateOnly(heute.Year - 1, month, day).DayNumber - heute.DayNumber;
                if (tageLetztesJahr >= -birthdayRemindAfter) { tage = tageLetztesJahr; }
            }
            return new { Entity = x, Tage = tage, OriginalGeb = g };
        })
            .Where(x => x.Tage >= -birthdayRemindAfter && x.Tage <= birthdayRemindLimit).OrderBy(x => x.Tage).Select(x =>
            {
                var gebDatum = x.OriginalGeb; var alter = heute.Year - gebDatum.Year; if (x.Tage > 0) { alter--; } // Noch nicht gehabt dieses Jahr
                return (Datum: gebDatum, Name: x.Entity.DisplayName, Alter: alter, x.Tage, Id: x.Entity.UniqueId);
            })
            .ToList();
        if (bevorstehendeGeburtstage.Count > 0 || showIfEmpty)
        {
            using var frm = new FrmBirthdays(sColorScheme, bevorstehendeGeburtstage, birthdayRemindLimit, birthdayRemindAfter, isLocal) { BirthdayAutoShow = autoShow };

            if (frm.ShowDialog() == DialogResult.OK && frm.SelectionIndex >= 0)
            {
                var selectedId = bevorstehendeGeburtstage[frm.SelectionIndex].Id;
                var item = bs.List.Cast<IContactEntity>().FirstOrDefault(x => x.UniqueId == selectedId);
                if (item != null)
                {
                    bs.Position = bs.IndexOf(item);
                    if (dgv.CurrentRow != null) { dgv.FirstDisplayedScrollingRowIndex = dgv.CurrentRow.Index; }
                }
            }
            birthdayRemindLimit = frm.BirthdayRemindLimit;
            birthdayRemindAfter = frm.BirthdayRemindAfter;
            if (isLocal) { birthdayAddressShow = frm.BirthdayAutoShow; }
            else { birthdayContactShow = frm.BirthdayAutoShow; }
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
                contextDgvMenu.Show(contactDGV, new Point(e.X, e.Y));
            }
        }
    }

    private async void MainDropDown_Opening(object? sender, CancelEventArgs e)
    {
        // Sicherheitscheck: Nur prüfen, wenn wir im Google-Tab sind
        if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact currentContact)
        {
            // 1. Fall: Ein ganz neuer Kontakt ohne ResourceName (ID)
            if (string.IsNullOrEmpty(currentContact.ResourceName))
            {
                if (CheckNewContactTidyUp())
                {
                    // Menü-Öffnen abbrechen, um asynchron zu speichern
                    e.Cancel = true;
                    await CreateContactAsync();

                    // Nach dem Speichern das Menü ggf. manuell wieder öffnen 
                    // oder den User fortfahren lassen
                    return;
                }
            }

            // 2. Fall: Bestehender Kontakt wurde geändert
            if (ContactChanges_Check())
            {
                // Wir brechen das Aufklappen ab, damit der User den TaskDialog sieht
                e.Cancel = true;

                // Startet den asynchronen Vergleich und Dialog
                await AskSaveContactChangesAsync();

                // Wenn nach dem Dialog keine Änderungen mehr da sind (User hat 'Speichern' oder 'Verwerfen' geklickt)
                if (!ContactChanges_Check() && sender is ToolStripDropDownItem item)
                {
                    // Menü jetzt manuell aufklappen
                    item.ShowDropDown();
                }
            }
        }
    }

    private void RecentToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        recentToolStripMenuItem.DropDownItems.Clear();
        var first = true;
        foreach (var file in recentFiles)
        {
            if (file == _databaseFilePath) { continue; }

            var item = new ToolStripMenuItem(file)
            {
                Image = Resources.address_book16,
                ShortcutKeyDisplayString = first ? "F12" : string.Empty
            };

            first = false;

            // WICHTIG: Hier muss "async" vor die Parameter (s, e)
            item.Click += async (s, e) =>
            {
                if (addressBindingSource != null)
                {
                    // Jetzt funktioniert await, weil das Lambda async ist
                    await SaveSQLDatabaseAsync(true);
                }

                // ConnectSQLDatabase wird erst ausgeführt, wenn SaveSQLDatabaseAsync fertig ist
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
        if (dokuListView.SelectedItems.Count > 0)
        {
            var index = dokuListView.SelectedIndices[0];
            foreach (ListViewItem item in dokuListView.SelectedItems) { dokuListView.Items.Remove(item); }
            if (dokuListView.Items.Count > 0) // Neue Selektion setzen, damit der Nutzer nicht den Fokus verliert
            {
                if (index >= dokuListView.Items.Count) { index = dokuListView.Items.Count - 1; }
                dokuListView.Items[index].Selected = true;
                dokuListView.Items[index].EnsureVisible();
            }
            SyncDocumentsToEntity();
        }
    }

    private void DokuShowTSButton_Click(object sender, EventArgs e)
    {
        if (dokuListView.SelectedItems.Count == 1)
        {
            var filePath = dokuListView.SelectedItems[0].Text;
            Utils.StartFile(Handle, filePath);
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
            SyncDocumentsToEntity();
        }
    }

    private void SyncDocumentsToEntity()
    {
        if (addressBindingSource?.Current is not Adresse selectedAddress) { return; }
        selectedAddress.Dokumente.Clear();  // erst alles löschen
        foreach (ListViewItem item in dokuListView.Items)  // und dann neu hinzufügen
        {
            selectedAddress.Dokumente.Add(new Dokument
            {
                Dateipfad = item.Text,
                AdressId = selectedAddress.Id,
                Adresse = selectedAddress
            });
        }
        tabPageDoku.ImageIndex = dokuListView.Items.Count > 0 ? 4 : 3;
        UpdateSaveButton();
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
            Utils.StartFile(Handle, dokuListView.SelectedItems[0].Text);
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

    private void SearchTextBox_TextChanged(object sender, EventArgs e) // Das ist die Search-Funktion für Dokumente (nichte DGV)
    {
        if (!searchTextBox.Focused) { return; }
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
        if (hit.Item != null && hit.SubItem != null && hit.Item.SubItems.IndexOf(hit.SubItem) == 0) { Utils.StartFile(Handle, hit.Item.Text); }
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
            item.SubItems.Add(Utils.FormatBytes(info.Length));
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
                SyncDocumentsToEntity();
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
                    SyncDocumentsToEntity();
                    tabulation.SelectedTab = tabPageDoku;
                }
                else if (tabControl.SelectedTab == contactTabPage)
                {
                    Utils.MsgTaskDlg(Handle, "Funktion nicht verfügbar", "Google-Kontakte haben beschränkte Feldgrößen", TaskDialogIcon.Information);
                }
            }
        }

        else if (result == copyButton)
        {
            try { Clipboard.SetText(text); }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
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

    private void UpdateSaveButton()
    {
        if (InvokeRequired) // fallls der Aufruf nicht im UI-Thread erfolgt
        {
            BeginInvoke(UpdateSaveButton);
            return;
        }
        if (tabControl.SelectedTab == addressTabPage) { saveTSButton.Enabled = HasRealEFChanges(); } // _context?.ChangeTracker.HasChanges() ?? false; }
        else if (tabControl.SelectedTab == contactTabPage) { saveTSButton.Enabled = ContactChanges_Check() || contactNewRowIndex >= 0; }
    }

    //private bool HasRealEFChanges()
    //{
    //    if (_context == null) { return false; }
    //    _context.ChangeTracker.DetectChanges();  // Zwingt EF Core, den aktuellen Zustand der Objekte mit den Snapshots zu vergleichen
    //    foreach (var entry in _context.ChangeTracker.Entries())
    //    {
    //        if (entry.State == EntityState.Added || entry.State == EntityState.Deleted) { return true; }
    //        if (entry.State == EntityState.Modified)
    //        {
    //            foreach (var prop in entry.Properties)
    //            {
    //                if (prop.IsModified)
    //                {
    //                    var original = prop.OriginalValue;
    //                    var current = prop.CurrentValue;
    //                    if (Equals(original, current)) { continue; }  // direkter Vergleich (für int, DateOnly, etc.)


    //                    if (prop.Metadata.ClrType == typeof(string)) // Spezialbehandlung: String "" wie null behandeln!
    //                    {
    //                        var sOriginal = original as string;
    //                        var sCurrent = current as string;

    //                        if (string.IsNullOrEmpty(sOriginal) && string.IsNullOrEmpty(sCurrent)) { continue; } // Phantom-Änderung
    //                    }
    //                    return true; // Wenn wir hier ankommen, ist es eine echte Änderung
    //                }
    //            }
    //        }
    //    }
    //    return false;
    //}

    private bool HasRealEFChanges()
    {
        if (_context == null) { return false; }

        _context.ChangeTracker.DetectChanges(); // Abgleich Snapshots vs. aktuelle Werte

        foreach (var entry in _context.ChangeTracker.Entries())
        {
            // Neu hinzugefügte oder gelöschte sind immer echte Änderungen
            if (entry.State == EntityState.Added || entry.State == EntityState.Deleted) { return true; }

            if (entry.State == EntityState.Modified)
            {
                foreach (var prop in entry.Properties)
                {
                    if (prop.IsModified)
                    {
                        var original = prop.OriginalValue;
                        var current = prop.CurrentValue;

                        // 1. Direkter Vergleich (für int, DateOnly, bool, etc.)
                        if (Equals(original, current)) { continue; }

                        // 2. Spezialbehandlung für Strings (null == "")
                        if (prop.Metadata.ClrType == typeof(string))
                        {
                            var sOriginal = original as string;
                            var sCurrent = current as string;

                            // Behandle null wie leeren String ("")
                            // Das löst das Problem, dass der SaveButton beim Reinklicken anspringt
                            var sOrigClean = sOriginal ?? string.Empty;
                            var sCurrClean = sCurrent ?? string.Empty;

                            if (sOrigClean == sCurrClean)
                            {
                                continue; // Phantom-Änderung (null vs empty) -> Ignorieren
                            }
                        }

                        // Wenn wir hier ankommen, ist es eine echte Änderung
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private bool ContactChanges_Check() // für saveTSButton.Enabled oder NotEnabled
    {
        if (_originalContactSnapshot is null || contactBindingSource.Current is not Contact) { return false; }
        foreach (var (control, propName) in editControlsDictionary) // alle TextBoxen und ComboBoxen via Dictionary
        {
            var propInfo = typeof(Contact).GetProperty(propName);
            if (propInfo is null) { continue; }
            var originalVal = propInfo.GetValue(_originalContactSnapshot)?.ToString() ?? string.Empty;
            if (!string.Equals(originalVal, control.Text, StringComparison.Ordinal)) { return true; }
        }
        var uiDateClean = maskedTextBox.Text.Replace(".", "").Replace("_", "").Replace(" ", "").Trim();
        var originalDateClean = _originalContactSnapshot.Geburtstag.HasValue ? _originalContactSnapshot.Geburtstag.Value.ToString("ddMMyyyy") : string.Empty;
        if (!string.Equals(uiDateClean, originalDateClean, StringComparison.Ordinal)) { return true; }
        return false; // alle Felder identisch mit dem Snapshot; Foto-Änderungen werden hier nicht geprüft
    }

    private void Clear_SearchTextBox()
    {
        // 1. Das aktuell ausgewählte Objekt merken (Sicherer als der Index!)
        object? selectedItem = null;
        BindingSource? activeBs = null;
        DataGridView? activeDgv = null;

        if (tabControl.SelectedTab == addressTabPage)
        {
            activeBs = addressBindingSource;
            activeDgv = addressDGV;
            selectedItem = addressBindingSource.Current;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            activeBs = contactBindingSource;
            activeDgv = contactDGV;
            selectedItem = contactBindingSource.Current;
        }

        // 2. Suche leeren (Löst Filter-Reset aus)
        ignoreSearchChange = true; // Verhindert unnötige Zwischen-Events
        searchTSTextBox.Text = string.Empty;

        // Wichtig: Den Filter auch wirklich anwenden/aufheben
        ApplyGlobalSearch(string.Empty);
        ignoreSearchChange = false;

        // 3. Den Fokus auf das gemerkte Objekt zurücksetzen
        if (activeBs != null && activeDgv != null && selectedItem != null)
        {
            // Wir suchen das Objekt in der nun ungefilterten Liste
            var newIndex = activeBs.IndexOf(selectedItem);
            if (newIndex >= 0)
            {
                activeBs.Position = newIndex;

                // Scroll-Position korrigieren, damit die Zeile sichtbar ist
                if (activeDgv.Rows.Count > newIndex)
                {
                    activeDgv.FirstDisplayedScrollingRowIndex = newIndex;
                }
            }
        }

        searchTSTextBox.Focus();
    }

    public static bool? GetGender(string name) // true für weiblich, false für männlich, null wenn unbekannt
    {
        if (string.IsNullOrWhiteSpace(name)) { return null; }
        return nameGenderMap.TryGetValue(name.Trim(), out var isFemale) ? isFemale : null;
    }

    private void WeiblicheVornamenToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartFile(Handle, girlPath);

    private void MännlicheVornamenToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartFile(Handle, boysPath);

    private void WebsiteToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartLink(Handle, @"https://www.netradio.info/address");

    private void GithubToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartLink(Handle, @"https://github.com/ophthalmos/Adressen");

    private void HelpdokuTSMenuItem_Click(object sender, EventArgs e) => Utils.StartFile(Handle, "AdressenKontakte.pdf");

    private void SortNamesToolStripMenuItem_Click(object sender, EventArgs e)
    {
        SortNameFiles(girlPath);
        SortNameFiles(boysPath);
    }

    private void SortNameFiles(string path)
    {
        if (!File.Exists(path)) { Utils.MsgTaskDlg(Handle, "Datei existiert nicht", boysPath); }
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
                Utils.MsgTaskDlg(Handle, message, path, TaskDialogIcon.Information);
            }
            catch (Exception ex) { Utils.MsgTaskDlg(Handle, "Fehler beim Sortieren der Vornamen", ex.Message); }
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
                Utils.MsgTaskDlg(Handle, "Namen in beiden Dateien gefunden", string.Join(Environment.NewLine, namesInBothFiles), TaskDialogIcon.Information);
            }
            else
            {
                Utils.MsgTaskDlg(Handle, "Keine Duplikate", "Es wurden keine Namen gefunden, die in beiden Dateien vorkommen.", TaskDialogIcon.Information);
            }
        }
        else { Utils.MsgTaskDlg(Handle, "Dateien nicht gefunden", girlPath + Environment.NewLine + boysPath); }
    }

    private void TermsofuseToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartLink(Handle, "https://www.netradio.info/adressen-terms-of-use/");
    private void PrivacypolicyToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartLink(Handle, "https://www.netradio.info/adressen-privacy-policy/");
    private void LicenseTxtToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartFile(Handle, "Lizenzvereinbarung.txt");

    //private void AdressenMitBriefToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == addressTabPage) { ApplyAddressPredicate(a => a.Dokumente.Count != 0, "… mit Briefverweis"); }
    //}

    //private void PhotoPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == addressTabPage) { ApplyAddressPredicate(a => a.Foto != null, "… mit Bild"); }
    //    else if (tabControl.SelectedTab == contactTabPage) { ApplyContactPredicate(c => !string.IsNullOrWhiteSpace(c.PhotoUrl), "… mit Bild"); }
    //}

    //private void PhotoMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == addressTabPage) { ApplyAddressPredicate(a => a.Foto == null, "… ohne Bild"); }
    //    else if (tabControl.SelectedTab == contactTabPage) { ApplyContactPredicate(c => string.IsNullOrWhiteSpace(c.PhotoUrl), "… ohne Bild"); }
    //}

    //private void MailPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == addressTabPage) { ApplyAddressPredicate(a => !string.IsNullOrWhiteSpace(a.Mail1) || !string.IsNullOrWhiteSpace(a.Mail2), "… mit E-Mail"); }
    //    else if (tabControl.SelectedTab == contactTabPage) { ApplyContactPredicate(a => !string.IsNullOrWhiteSpace(a.Mail1) || !string.IsNullOrWhiteSpace(a.Mail2), "… mit E-Mail"); }
    //}

    //private void MailMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == addressTabPage) { ApplyAddressPredicate(a => string.IsNullOrWhiteSpace(a.Mail1) && string.IsNullOrWhiteSpace(a.Mail2), "… ohne E-Mail"); }
    //    else if (tabControl.SelectedTab == contactTabPage) { ApplyContactPredicate(a => string.IsNullOrWhiteSpace(a.Mail1) && string.IsNullOrWhiteSpace(a.Mail2), "… ohne E-Mail"); }
    //}

    //private void TelephonePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == addressTabPage) { ApplyAddressPredicate(a => !string.IsNullOrWhiteSpace(a.Telefon1) || !string.IsNullOrWhiteSpace(a.Telefon2), "… mit Telefonnummer"); }
    //    else if (tabControl.SelectedTab == contactTabPage) { ApplyContactPredicate(a => !string.IsNullOrWhiteSpace(a.Mail1) || !string.IsNullOrWhiteSpace(a.Mail2), "… mit Telefonnummer"); }
    //}

    //private void TelephoneMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == addressTabPage) { ApplyAddressPredicate(a => string.IsNullOrWhiteSpace(a.Telefon1) && string.IsNullOrWhiteSpace(a.Telefon2), "… ohne Telefonnummer"); }
    //    else if (tabControl.SelectedTab == contactTabPage) { ApplyContactPredicate(a => string.IsNullOrWhiteSpace(a.Mail1) && string.IsNullOrWhiteSpace(a.Mail2), "… ohne Telefonnummer"); }
    //}

    //private void MobilePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == addressTabPage) { ApplyAddressPredicate(a => !string.IsNullOrWhiteSpace(a.Mobil), "… mit Mobilnummer"); }
    //    else if (tabControl.SelectedTab == contactTabPage) { ApplyContactPredicate(a => !string.IsNullOrWhiteSpace(a.Mobil), "… mit Mobilnummer"); }
    //}

    //private void MobileMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    if (tabControl.SelectedTab == addressTabPage) { ApplyAddressPredicate(a => string.IsNullOrWhiteSpace(a.Mobil), "… ohne Mobilnummer"); }
    //    else if (tabControl.SelectedTab == contactTabPage) { ApplyContactPredicate(a => string.IsNullOrWhiteSpace(a.Mobil), "… ohne Mobilnummer"); }
    //}

    // --- Spezialfall: Dokumente gibt es nur bei Adressen ---
    private void AdressenMitBriefToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => a.Dokumente.Count != 0, "… mit Briefverweis", "Adressen");
        }
    }

    // --- Foto Filter ---
    private void PhotoPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // 1. Wir fragen die DB: Welche IDs haben ein Foto?
            // Das generiert ein effizientes SQL ("SELECT Id FROM Adressen WHERE FotoId IS NOT NULL")
            var idsWithPhoto = _context.Adressen
                .Where(a => a.Foto != null)
                .Select(a => a.Id)
                .ToHashSet(); // HashSet für extrem schnelle Suche

            // 2. Wir filtern die lokale Liste anhand dieser IDs
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => idsWithPhoto.Contains(a.Id), "… mit Bild", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => !string.IsNullOrWhiteSpace(c.PhotoUrl), "… mit Bild", "Google Kontakte");
        }
    }

    private void PhotoMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // 1. Gleiches Spiel: IDs holen
            var idsWithPhoto = _context.Adressen
                .Where(a => a.Foto != null)
                .Select(a => a.Id)
                .ToHashSet();

            // 2. Filter umdrehen: Zeige alle, deren ID NICHT in der Liste ist
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => !idsWithPhoto.Contains(a.Id), "… ohne Bild", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => string.IsNullOrWhiteSpace(c.PhotoUrl), "… ohne Bild", "Google Kontakte");
        }
    }

    // --- E-Mail Filter ---
    private void MailPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => !string.IsNullOrWhiteSpace(a.Mail1) || !string.IsNullOrWhiteSpace(a.Mail2), "… mit E-Mail", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => !string.IsNullOrWhiteSpace(c.Mail1) || !string.IsNullOrWhiteSpace(c.Mail2), "… mit E-Mail", "Google Kontakte");
        }
    }

    private void MailMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => string.IsNullOrWhiteSpace(a.Mail1) && string.IsNullOrWhiteSpace(a.Mail2), "… ohne E-Mail", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => string.IsNullOrWhiteSpace(c.Mail1) && string.IsNullOrWhiteSpace(c.Mail2), "… ohne E-Mail", "Google Kontakte");
        }
    }

    // --- Telefon Filter (Hier waren deine Copy-Paste Fehler!) ---
    private void TelephonePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => !string.IsNullOrWhiteSpace(a.Telefon1) || !string.IsNullOrWhiteSpace(a.Telefon2), "… mit Telefonnummer", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            // KORRIGIERT: c.Telefon statt c.Mail
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => !string.IsNullOrWhiteSpace(c.Telefon1) || !string.IsNullOrWhiteSpace(c.Telefon2), "… mit Telefonnummer", "Google Kontakte");
        }
    }

    private void TelephoneMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => string.IsNullOrWhiteSpace(a.Telefon1) && string.IsNullOrWhiteSpace(a.Telefon2), "… ohne Telefonnummer", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            // KORRIGIERT: c.Telefon statt c.Mail
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => string.IsNullOrWhiteSpace(c.Telefon1) && string.IsNullOrWhiteSpace(c.Telefon2), "… ohne Telefonnummer", "Google Kontakte");
        }
    }

    // --- Mobil Filter ---
    private void MobilePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => !string.IsNullOrWhiteSpace(a.Mobil), "… mit Mobilnummer", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => !string.IsNullOrWhiteSpace(c.Mobil), "… mit Mobilnummer", "Google Kontakte");
        }
    }

    private void MobileMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV, a => string.IsNullOrWhiteSpace(a.Mobil), "… ohne Mobilnummer", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV, c => string.IsNullOrWhiteSpace(c.Mobil), "… ohne Mobilnummer", "Google Kontakte");
        }
    }

    private void DatePlusFilterMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // Wir prüfen direkt das Nullable DateOnly Feld "Geburtstag"
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => a.Geburtstag.HasValue, "… mit Geburtsdatum", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            // Auch für Google Kontakte (vorausgesetzt, das Feld heißt dort ähnlich)
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => c.Geburtstag.HasValue, "… mit Geburtsdatum", "Google Kontakte");
        }
    }

    private void DateMinusFilterMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // Wir filtern auf alle Adressen, deren Geburtstag NICHT gesetzt ist (null)
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => !a.Geburtstag.HasValue, "… ohne Geburtsdatum", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            // Dieselbe Logik für die Google Kontakte
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => !c.Geburtstag.HasValue, "… ohne Geburtsdatum", "Google Kontakte");
        }
    }

    // --- Menü-Status aktualisieren ---
    private void FilterlToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        var isAddressTab = tabControl.SelectedTab == addressTabPage && addressDGV.Rows.Count > 0;
        var isContactTab = tabControl.SelectedTab == contactTabPage && contactDGV.Rows.Count > 0;

        adressenMitBriefToolStripMenuItem.Enabled = isAddressTab;

        // Gemeinsame Filter aktivieren, wenn Daten vorhanden sind
        var enableCommonFilters = isAddressTab || isContactTab;

        photoPlusFilterToolStripMenuItem.Enabled = photoMinusFilterToolStripMenuItem.Enabled =
        mailPlusFilterToolStripMenuItem.Enabled = mailMinusFilterToolStripMenuItem.Enabled =
        telephonePlusFilterToolStripMenuItem.Enabled = telephoneMinusFilterToolStripMenuItem.Enabled =
        mobilePlusFilterToolStripMenuItem.Enabled = mobileMinusFilterToolStripMenuItem.Enabled = enableCommonFilters;
    }

    private void ExecuteFilter<T>(IEnumerable<T> sourceList, BindingSource bs, DataGridView dgv, Func<T, bool> predicate, string statusText, string entityName)
    {
        if (sourceList == null || bs == null) { return; }
        var currencyManager = BindingContext?[bs] as CurrencyManager;
        try
        {
            currencyManager?.SuspendBinding();
            dgv.CurrentCell = null; // Verhindert Index-Fehler beim Wechsel der DataSource

            // Hier passiert die Magie: Filtern der Liste
            var filteredList = sourceList.Where(predicate).ToList();
            bs.DataSource = filteredList;

            // UI Updates
            filterRemoveToolStripMenuItem.Visible = true;
            flexiTSStatusLabel.Text = statusText;

            // Statusbar aktualisieren
            var totalCount = sourceList.Count();
            var visibleCount = filteredList.Count;

            toolStripStatusLabel.Text = visibleCount == totalCount ? $"{totalCount} {entityName}" : $"{visibleCount}/{totalCount} {entityName}";
            // Erste Zeile markieren, falls vorhanden
            if (visibleCount > 0 && dgv.Rows.Count > 0) { dgv.Rows[0].Selected = true; }
        }
        catch (Exception ex) { Debug.WriteLine(ex.Message); }
        finally { currencyManager?.ResumeBinding(); }
    }

    private async void TopAlignZoomPictureBox_DoubleClick(object sender, EventArgs e)
    {
        // Sicherheitschecks
        if ((tabControl.SelectedTab == addressTabPage && addressBindingSource.Current == null) ||
            (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count == 0))
        {
            return;
        }

        openFileDialog.Title = "Foto auswählen";
        openFileDialog.Filter = $"Bilddateien|{string.Join(";", pictureBoxExtensions.Select(ext => "*" + ext))}|Alle Dateien|*.*";
        openFileDialog.Multiselect = false;
        openFileDialog.FileName = string.Empty;
        openFileDialog.CheckFileExists = true;

        if (openFileDialog.ShowDialog(this) != DialogResult.OK) { return; }

        // ---------------------------------------------------------
        // FALL 1: Lokale Datenbank (EF Core)
        // ---------------------------------------------------------
        if (tabControl.SelectedTab == addressTabPage)
        {
            // KORREKTUR 1: Wir holen das Objekt direkt aus der BindingSource und casten auf 'Adresse'
            if (addressBindingSource.Current is Adresse adresse)
            {
                var bildDaten = await File.ReadAllBytesAsync(openFileDialog.FileName);
                if (bildDaten.Length == 0)
                {
                    Utils.MsgTaskDlg(Handle, "Fehler", "Die Datei ist leer.", TaskDialogIcon.ShieldErrorRedBar);
                    return;
                }

                Image? loadedImage = null;
                Image? scaledImage = null;

                try
                {
                    // Alte Anzeige bereinigen
                    topAlignZoomPictureBox.Image?.Dispose();
                    topAlignZoomPictureBox.Image = null;

                    using var ms = new MemoryStream(bildDaten);
                    loadedImage = Image.FromStream(ms);
                    var originalFormat = loadedImage.RawFormat;
                    Utils.WendeExifOrientierungAn(loadedImage);

                    Image finalImage;

                    // Skalierung (Logik aus Ihrem Code übernommen)
                    if (loadedImage.Width > 100)
                    {
                        scaledImage = Utils.SkaliereBildDaten(loadedImage, 100);
                        finalImage = scaledImage;
                    }
                    else { finalImage = loadedImage; }

                    // Anzeige aktualisieren
                    topAlignZoomPictureBox.Image = finalImage; // PictureBox übernimmt Referenz (nicht disposen!)
                    delPictboxToolStripButton.Enabled = true;

                    // Bilddaten für DB vorbereiten
                    byte[] datenZumSpeichern;
                    using (var outputMs = new MemoryStream())
                    {
                        var saveFormat = originalFormat.Equals(ImageFormat.Png) ? ImageFormat.Png : ImageFormat.Jpeg;
                        finalImage.Save(outputMs, saveFormat);
                        datenZumSpeichern = outputMs.ToArray();
                    }

                    // KORREKTUR 2: Speichern über EF Core Beziehung statt SQL-Helper
                    adresse.Foto ??= new Foto(); // Neue Foto-Entity anlegen, falls noch keine existiert

                    // Blob aktualisieren
                    adresse.Foto.Fotodaten = datenZumSpeichern;

                    // Änderungen speichern
                    // Falls Sie einen zentralen "Speichern"-Button haben, reicht auch:
                    // Wenn das Foto sofort in die DB soll:
                    _context?.SaveChanges();
                    addressBindingSource.ResetCurrentItem(); // Wichtig: BindingSource aktualisieren, damit UI den Status kennt
                    saveTSButton.Enabled = false; // (passiert via BindingSource-Event meist automatisch)

                    // Aufräumen der lokalen Referenzen (nicht das Bild in der PB!)
                    loadedImage = null;
                    scaledImage = null;
                }
                catch (Exception ex)
                {
                    loadedImage?.Dispose();
                    scaledImage?.Dispose();
                    Utils.ErrTaskDlg(Handle, ex);
                }
            }
        }
        // ---------------------------------------------------------
        // FALL 2: Google Kontakte (Ihr Code, weitgehend unverändert)
        // ---------------------------------------------------------
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
                    Utils.WendeExifOrientierungAn(originalImage);
                    if (fs.Length > 1024 * 1024)
                    {
                        Utils.MsgTaskDlg(Handle, "Automatische Größenreduzierung", $"Die Dateigröße ist größer als 1 MB ({Utils.FormatBytes(fs.Length)}).\nEs erfolgt eine Skalierung auf 250 Pixel Breite.", TaskDialogIcon.ShieldWarningYellowBar);
                        workingImage = Utils.SkaliereBildDaten(originalImage, 250);
                    }
                    else { workingImage = (Image)originalImage.Clone(); }
                }
                //var ressource = contactDGV.Rows[contactDGV.SelectedRows[0].Index].Cells["Ressource"]?.Value?.ToString() ?? string.Empty;
                if (contactBindingSource.Current is not Contact currentContact) { return; }
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
                        workingImage = Utils.BeschneideZuQuadrat(workingImage, null);
                        finalImageToUpload = workingImage;
                        finalImageForDisplay = (Image)workingImage.Clone();
                    }
                    else if (centerRadio?.Checked == true)
                    {
                        intermediateImageToDispose = workingImage; // Das alte 'workingImage' zum Dispose vormerken
                        workingImage = Utils.BeschneideZuQuadrat(workingImage, false); // 'workingImage' ist jetzt das *neue* beschnittene
                        finalImageToUpload = workingImage; // Hochladen
                        finalImageForDisplay = (Image)workingImage.Clone(); // Anzeigen
                    }
                    else if (downRadio?.Checked == true)
                    {
                        intermediateImageToDispose = workingImage;
                        workingImage = Utils.BeschneideZuQuadrat(workingImage, true);
                        finalImageToUpload = workingImage;
                        finalImageForDisplay = (Image)workingImage.Clone();
                    }
                    else if (skipRadio?.Checked == true)
                    {
                        finalImageToUpload = workingImage; // 'workingImage' wird *nicht* ersetzt
                        finalImageForDisplay = Utils.ReduziereWieGoogle(workingImage, 100);
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
                    try { await UpdateContactPhotoAsync(currentContact, finalImageToUpload!, origImgFormat, () => progressPage.Buttons.First().PerformClick()); }
                    finally { workingImage?.Dispose(); }  // finalImageForDisplay wird von PictureBox verwaltet, darf hier nicht disposed werden    
                };
                TaskDialog.ShowDialog(Handle, initialPage);
                delPictboxToolStripButton.Enabled = true;
            }
            catch (Exception ex)
            {
                Utils.MsgTaskDlg(Handle, $"Fehler beim Laden: {ex.GetType()}", $"Bild konnte nicht geladen werden: {ex.Message}", TaskDialogIcon.Error);
                workingImage?.Dispose();
                finalImageForDisplay?.Dispose();
            }
        }
    }

    private void FilterRemoveToolStripMenuItem_Click(object sender, EventArgs e)
    {
        ignoreSearchChange = true;
        searchTSTextBox.TextBox.Clear();
        tsClearLabel.Visible = false;
        ignoreSearchChange = false;
        if (tabControl.SelectedTab == addressTabPage)
        {
            if (_context == null) { return; }
            var currencyManager = BindingContext?[addressBindingSource] as CurrencyManager;
            currencyManager?.SuspendBinding(); //  SuspendBinding nur aufrufen, wenn der Manager existiert
            try
            {
                addressDGV.CurrentCell = null;  // Wichtig: CurrentCell auf null setzen, BEVOR die DataSource getauscht wird
                addressBindingSource.DataSource = _context.Adressen.Local.ToBindingList();
            }
            finally { currencyManager?.ResumeBinding(); }
            UpdateAddressStatusBar();
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            if (_allGoogleContacts != null && contactBindingSource != null) { contactBindingSource.DataSource = _allGoogleContacts; }
            UpdateContactStatusBar();
        }
        filterRemoveToolStripMenuItem.Visible = false;
        flexiTSStatusLabel.Text = string.Empty;
    }

    private void AddPictboxToolStripButton_Click(object sender, EventArgs e) => TopAlignZoomPictureBox_DoubleClick(addPictboxToolStripButton, EventArgs.Empty);


    private async void DelPictboxToolStripButton_Click(object sender, EventArgs e)
    {
        // --- FALL A: SQL ADRESSEN ---
        if (tabControl.SelectedTab == addressTabPage && addressBindingSource.Current is Adresse adresse)
        {
            var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Adressen", "Möchten Sie das Bild wirklich löschen?",
                    "Es wird unwiderruflich aus der Datenbank entfernt.", "&Löschen", "&Belassen", false);
            if (isYes)
            {
                try
                {
                    if (adresse.Foto != null)
                    {
                        _context?.Fotos.Remove(adresse.Foto);
                        // EF Core 10 Tipp: Wir setzen die Referenz explizit auf null
                        adresse.Foto = null;

                        await _context!.SaveChangesAsync();

                        topAlignZoomPictureBox.Image = Properties.Resources.AddressBild100;
                        delPictboxToolStripButton.Enabled = false;

                        addressBindingSource.ResetCurrentItem();
                    }
                }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
        }
        // --- FALL B: GOOGLE KONTAKTE ---
        else if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact googleKontakt)
        {
            var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Google Kontakte", "Möchten Sie das Bild wirklich löschen?",
                    "Das Foto wird bei Google unwiderruflich entfernt.", "&Löschen", "&Belassen", false);
            if (isYes)
            {
                try
                {
                    // WICHTIG: Wir übergeben das OBJEKT googleKontakt
                    await DeleteContactPhotoAsync(googleKontakt);

                    // UI-Update
                    topAlignZoomPictureBox.Image = Properties.Resources.ContactBild100; // Spezielles Kontakt-Icon
                    delPictboxToolStripButton.Enabled = false;

                    // Da das Foto weg ist, muss die Spalte im Grid ("alle mit Bild") aktualisiert werden
                    contactBindingSource.ResetCurrentItem();
                }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
        }
    }

    private async void Move2OtherDGVToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact googleKontakt)
        {
            var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Google Kontakte",
                "Möchten Sie den Kontakt in die lokale Datenbank verschieben?",
                "Der Google-Kontakt wird dabei unwiderruflich gelöscht!",
                "&Verschieben", "&Abbrechen", false);
            if (isYes)
            {
                try
                {
                    CopyToOtherDGVMenuItem_Click(move2OtherDGVToolStripMenuItem, EventArgs.Empty);
                    await DeleteGoogleContactAsync(googleKontakt);
                    _allGoogleContacts?.Remove(googleKontakt);
                    contactBindingSource.RemoveCurrent();
                    tabControl.SelectedTab = addressTabPage; // Wechsle zum Adress-Tab um neue Adresse zu sehen
                    if (addressBindingSource.Count > 0) { addressBindingSource.MoveLast(); }
                    UpdateContactStatusBar();
                }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
        }
    }

    private void UpdateMembershipTags()
    {
        var isContactTab = tabControl.SelectedTab == contactTabPage;
        var groupsList = isContactTab ? curContactMemberships : curAddressMemberships;
        flowLayoutPanel.Controls.Clear();
        foreach (var membership in groupsList)
        {
            var tagControl = new TagControl
            {
                Text = membership,
                Membership = membership
            };

            tagControl.DeleteClick += (sender, e) =>
            {
                var ctrl = sender as TagControl;
                var membershipToRemove = ctrl?.Membership;
                if (string.IsNullOrEmpty(membershipToRemove)) { return; }

                if (isContactTab) // --- Google Kontakte Logic ---
                {
                    curContactMemberships.Remove(membershipToRemove);
                    UpdateMembershipTags();
                    UpdateMembershipJson();
                }
                else
                {
                    if (addressBindingSource.Current is Adresse adresse)
                    {
                        var gruppeToDelete = adresse.Gruppen.FirstOrDefault(g => g.Name.Equals(membershipToRemove, StringComparison.OrdinalIgnoreCase));
                        if (gruppeToDelete != null)
                        {
                            adresse.Gruppen.Remove(gruppeToDelete); // 2. Beziehung entfernen (NICHT die Gruppe selbst löschen, nur die Verknüpfung!)
                            curAddressMemberships.Remove(membershipToRemove);
                            UpdateMembershipTags();
                            UpdateMembershipCBox();
                            UpdatePlaceholderVis();
                            addressBindingSource.ResetCurrentItem();
                        }
                    }
                }
            };
            flowLayoutPanel.Controls.Add(tagControl);
        }
        UpdatePlaceholderVis();
    }

    //private void TagPanel_MouseDeactivation(object? sender, EventArgs e)
    //{
    //    var currentControl = sender as Control;
    //    var currentPanel = (currentControl as Panel) ?? (currentControl?.Parent as Panel);
    //    if (currentPanel == null) { return; }
    //    var clientPoint = currentPanel.PointToClient(Cursor.Position);
    //    if (currentPanel.ClientRectangle.Contains(clientPoint)) { return; }
    //    var currentButton = currentPanel.Controls.OfType<Button>().FirstOrDefault();
    //    currentButton?.Enabled = false;
    //}

    private void TagButton_Click(object sender, EventArgs e)
    {
        var newMembershipName = tagComboBox.Text.Trim();
        if (string.IsNullOrEmpty(newMembershipName)) { return; }
        if (newMembershipName == "*") { newMembershipName = "★"; }

        // --- FALL 1: Google Kontakte (bleibt wie es war) ---
        if (tabControl.SelectedTab == contactTabPage)
        {
            if (curContactMemberships.Contains(newMembershipName)) { return; }
            curContactMemberships.Add(newMembershipName);
            allContactMemberships.Add(newMembershipName);

            UpdateMembershipTags();
            UpdateMembershipCBox();
            UpdateMembershipJson(); // Google nutzt weiterhin JSON/Strings
        }
        // --- FALL 2: Lokale EF Core Adressen ---
        else if (tabControl.SelectedTab == addressTabPage)
        {
            if (addressBindingSource.Current is Adresse adresse)
            {
                // 1. Prüfen, ob die Adresse die Gruppe schon hat
                if (adresse.Gruppen.Any(g => g.Name.Equals(newMembershipName, StringComparison.OrdinalIgnoreCase)))
                {
                    tagComboBox.SelectAll();
                    tagComboBox.Focus();
                    return;
                }

                // 2. Gruppe in der DB suchen oder neu erstellen
                // Wir schauen erst im ChangeTracker (.Local), dann in der DB
                var lowerMembershipName = newMembershipName.ToLower();

                // 1. Erst im lokalen Speicher schauen (hier funktioniert StringComparison, da C#)
                var gruppe = _context?.Gruppen.Local.FirstOrDefault(g => g.Name.Equals(newMembershipName, StringComparison.OrdinalIgnoreCase));

                // 2. Wenn nicht lokal gefunden, in der Datenbank suchen (hier ToLower() für SQL nutzen)
                gruppe ??= _context?.Gruppen.FirstOrDefault(g => g.Name.Equals(lowerMembershipName, StringComparison.CurrentCultureIgnoreCase));
                if (gruppe == null)
                {
                    // Neue Gruppe anlegen
                    gruppe = new Gruppe { Name = newMembershipName };
                    _context?.Gruppen.Add(gruppe); // Dem Context bekannt machen

                    // Auch zur globalen Liste für die ComboBox hinzufügen
                    allAddressMemberships.Add(newMembershipName);
                }

                // 3. Verknüpfung herstellen (M:N)
                adresse.Gruppen.Add(gruppe);

                // 4. UI aktualisieren (lokale String-Liste und Anzeige)
                curAddressMemberships.Add(newMembershipName);
                UpdateMembershipTags();
                UpdateMembershipCBox();

                // 5. Save-Button aktivieren (passiert meist automatisch via BindingSource Event, sonst:)
                // saveTSButton.Enabled = true; 
                //addressBindingSource.ResetCurrentItem(); // UI Refresh
            }
        }
    }

    private void UpdateMembershipJson()
    {
        if (tabControl.SelectedTab == contactTabPage)
        {
            if (contactBindingSource.Current is Contact contact) { contact.GroupNames = [.. curContactMemberships]; }
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
        // 1. Gruppenliste für den Dialog vorbereiten
        SortedSet<string> dialogGroups;

        if (tabControl.SelectedTab == addressTabPage)
        {
            if (_context == null) { return; }
            // SQL-Gruppen laden
            dialogGroups = new SortedSet<string>(
                _context.Gruppen.Local.Select(g => g.Name),
                StringComparer.OrdinalIgnoreCase
            );
        }
        else
        {
            // Google-Gruppen (existieren bereits als Set)
            dialogGroups = allContactMemberships;
        }

        // 2. Dialog anzeigen
        using var frm = new FrmGroupFilter(dialogGroups);
        if (frm.ShowDialog(this) != DialogResult.OK) { return; }

        var included = frm.IncludedGroups;
        var excluded = frm.ExcludedGroups;

        // Wenn gar nichts ausgewählt wurde -> Filter entfernen
        if (included.Count == 0 && excluded.Count == 0)
        {
            FilterRemoveToolStripMenuItem_Click(sender, e);
            return;
        }

        // 3. Lokale Hilfsfunktion: Die Filterlogik an EINER Stelle
        // Prüft für eine Liste von Gruppennamen, ob sie den Kriterien entspricht
        bool MatchesFilter(IEnumerable<string> itemGroups)
        {
            // Muss EINE der "Included" Gruppen enthalten (oder Include ist leer)
            var matchesInclude = included.Count == 0 || itemGroups.Any(g => included.Contains(g));

            // Darf KEINE der "Excluded" Gruppen enthalten (oder Exclude ist leer)
            var matchesExclude = excluded.Count == 0 || !itemGroups.Any(g => excluded.Contains(g));

            return matchesInclude && matchesExclude;
        }

        // 4. Generischen Filter ausführen
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(
                _context.Adressen.Local,
                addressBindingSource,
                addressDGV,
                // Bei Adressen müssen wir erst die Namen aus den Objekten holen
                a => MatchesFilter(a.Gruppen.Select(g => g.Name)),
                "… mit Gruppenfilter",
                "Adressen"
            );
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(
                _allGoogleContacts,
                contactBindingSource,
                contactDGV,
                // Bei Kontakten haben wir schon Strings
                c => MatchesFilter(c.GroupNames),
                "… mit Gruppenfilter",
                "Google Kontakte"
            );
        }
    }

    private async void ManageGroupsToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            if (_context == null) { return; }
            var groupDict = _context.Gruppen.Local.ToDictionary(g => g.Name, g => g.Adressen.Count);
            using var frm = new FrmGroupsEdit(groupDict);
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                var changes = frm.groupNameMap.Where(kvp => kvp.Key != kvp.Value || string.IsNullOrEmpty(kvp.Value)).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                if (changes.Count == 0) { return; }
                var needsSave = false;
                foreach (var (oldName, newName) in changes)
                {
                    if (oldName == "★") { continue; }  // Favoriten schützen
                    var groupEntity = _context.Gruppen.Local.FirstOrDefault(g => g.Name.Equals(oldName, StringComparison.OrdinalIgnoreCase));
                    if (groupEntity == null) { continue; }
                    if (string.IsNullOrWhiteSpace(newName))
                    {
                        _context.Gruppen.Remove(groupEntity);
                        allAddressMemberships.Remove(oldName);
                        needsSave = true;
                    }
                    else
                    {
                        groupEntity.Name = newName;
                        allAddressMemberships.Remove(oldName);
                        allAddressMemberships.Add(newName);
                        needsSave = true;
                    }
                }
                if (needsSave)
                {
                    await SaveSQLDatabaseAsync();
                    addressBindingSource.ResetBindings(false);
                    if (addressBindingSource.Current != null) { LoadGroupsForCurrentAddress(); }
                }
            }
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            var groupDict = new Dictionary<string, int>();
            if (_allGoogleContacts != null)
            {
                foreach (var contact in _allGoogleContacts)
                {
                    foreach (var gName in contact.GroupNames)
                    {
                        if (groupDict.TryGetValue(gName, out var count)) { groupDict[gName] = count + 1; }
                        else { groupDict[gName] = 1; }
                    }
                }
            }
            using var frm = new FrmGroupsEdit(groupDict);
            if (frm.ShowDialog(this) == DialogResult.OK) { ProcessGoogleGroupChanges(frm.groupNameMap); }
        }
    }

    private static void ProcessGoogleGroupChanges(Dictionary<string, string> groupChanges)
    {
        List<string> deleteChanges = [];
        List<string> renameChanges = [];
        var realChanges = groupChanges.Where(kvp => kvp.Key != kvp.Value || string.IsNullOrEmpty(kvp.Value)).ToDictionary(k => k.Key, k => k.Value);
        foreach (var kvp in realChanges)
        {
            if (!string.IsNullOrEmpty(kvp.Value)) { renameChanges.Add(kvp.Value); }
            else { deleteChanges.Add(kvp.Key); }
        }
        if (deleteChanges.Count == 0 && renameChanges.Count == 0) { return; }
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
            try { Utils.SetClipboardText(strValue); }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
    }

    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);
        if (m.Msg == NativeMethods.WM_SETTINGCHANGE)
        {
            var area = Marshal.PtrToStringUni(m.LParam);
            if (string.IsNullOrEmpty(area) || area == "ImmersiveColorSet")
            {
                Application.SetColorMode(SystemColorMode.System); // Zwingt die App, den Modus neu zu evaluieren
                UpdateAppearanceStatus(); // spezifische Farben für Controls und Grids anpassen
                Refresh(); // auch für Child Controls
                ToolStripManager.VisualStylesEnabled = true;  // ToolStrips brauchen manchmal einen extra Schubs für ihren Renderer
            }
        }
    }

    private void UpdateAppearanceStatus()
    {
        _isDarkMode = Application.SystemColorMode == SystemColorMode.Dark;
        if (Application.SystemColorMode == SystemColorMode.System) { _isDarkMode = Control.DefaultBackColor.R < 128; } //falls die Automatik hakt
        ConfigureDgvAppearance(addressDGV, Color.FromArgb(176, 125, 71)); // Dein Braun
        ConfigureDgvAppearance(contactDGV, Color.FromArgb(0, 102, 204));  // Blau (z.B. Windows Default Blue)
        foreach (var c in Utils.GetAllControls(this))
        {
            if (c is TextBox || c is MaskedTextBox || c is ComboBox)
            {
                c.BackColor = _isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
                c.ForeColor = _isDarkMode ? Color.White : Color.Black;
                c.Invalidate(); // ungültig machen  
                c.Update(); // sofortiges Neuzeichnen
            }
        }
        PerformLayout();
    }

    private void ConfigureDgvAppearance(DataGridView dgv, Color selectionColor)
    {
        dgv.SuspendLayout();
        dgv.RowsDefaultCellStyle.BackColor = Color.Empty;
        dgv.RowsDefaultCellStyle.ForeColor = Color.Empty;
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.Empty;
        dgv.AlternatingRowsDefaultCellStyle.ForeColor = Color.Empty;
        dgv.BackgroundColor = _isDarkMode ? Color.FromArgb(30, 30, 30) : SystemColors.AppWorkspace;
        dgv.GridColor = _isDarkMode ? Color.FromArgb(60, 60, 60) : SystemColors.ControlLight;
        dgv.EnableHeadersVisualStyles = false; // Muss false bleiben, damit Dark Mode Farben greifen
        if (_isDarkMode)
        {
            var darkHeader = Color.FromArgb(50, 50, 50);
            dgv.ColumnHeadersDefaultCellStyle.BackColor = darkHeader;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.RowHeadersDefaultCellStyle.BackColor = darkHeader;
            dgv.RowHeadersDefaultCellStyle.ForeColor = Color.White;
        }
        else
        {
            dgv.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.ControlLight;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.WindowText;
            dgv.RowHeadersDefaultCellStyle.BackColor = SystemColors.MenuBar;
            dgv.RowHeadersDefaultCellStyle.ForeColor = SystemColors.WindowText;
        }
        dgv.DefaultCellStyle.SelectionBackColor = selectionColor;
        dgv.DefaultCellStyle.SelectionForeColor = Color.White;
        dgv.RowsDefaultCellStyle.SelectionBackColor = selectionColor;
        dgv.RowsDefaultCellStyle.SelectionForeColor = Color.White;
        dgv.ResumeLayout();
    }

    private void AddressDGV_DataError(object sender, DataGridViewDataErrorEventArgs e)
    {
        if (e.Exception is IndexOutOfRangeException || e.Exception is ArgumentException)
        {
            e.Cancel = true;
            e.ThrowException = false;
        }
    }

    private void ContactBindingSource_ListChanged(object sender, ListChangedEventArgs e) => UpdateSaveButton();

    private void SearchTimer_Tick(object? sender, EventArgs e)
    {
        searchTimer.Stop();
        ApplyGlobalSearch(searchTSTextBox.TextBox.Text); // Da wir im UI-Thread sind, können wir direkt auf die TextBox zugreifen.
    }

    private async void ContactDGV_RowValidating(object sender, DataGridViewCellCancelEventArgs e) => await AskSaveContactChangesAsync();

    private void AddressDGV_SelectionChanged(object sender, EventArgs e) => scrollTimer.Start();

    private void ScrollTimer_Tick(object sender, EventArgs e) => scrollTimer.Stop();
}
