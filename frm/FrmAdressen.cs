using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Adressen.cls;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;

namespace Adressen;

public partial class FrmAdressen : Form
{
    private static readonly string appPath = Application.ExecutablePath; // EXE-Pfad
    private SQLiteConnection? liteConnection;
    private DataTable? dataTable;
    private string databaseFilePath = string.Empty; // Path.ChangeExtension(appPath, ".adb");
    private bool sAskBeforeSaveSQL = false; // Änderungen automatisch speichern
    private readonly string xmlPath;
    private readonly string tokenDir;
    private readonly string secretPath;
    private readonly string boysPath;
    private readonly string girlPath;
    private SQLiteDataAdapter dataAdapter = new();
    private readonly string cleanRegex = @"[^\+0-9]";
    private readonly string appName = Application.ProductName ?? "Adressen";
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
    private string[]? pSenderLines1;
    private string[]? pSenderLines2;
    private string[]? pSenderLines3;
    private string[]? pSenderLines4;
    private string[]? pSenderLines5;
    private string[]? pSenderLines6;
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
    private string[]? copyPattern1 = ["Anrede", "Präfix_Vorname_Zwischenname_Nachname", "StraßeNr", "PLZ_Ort"];
    private string[]? copyPattern2 = ["Telefon1", "Telefon2", "Mobil", "Fax"];
    private string[]? copyPattern3 = ["Mail1", "Mail2", "Internet"];
    private string[]? copyPattern4 = [];
    private string[]? copyPattern5 = [];
    private string[]? copyPattern6 = [];
    private readonly bool[] hideColumnStd = [true, true, false, false, true, true, true, false, false, false, false, false, true, true, true, false, false, false, false, false, false, false, false, true, true];
    private bool[] hideColumnArr = new bool[25];
    private int splitterPosition;
    private string windowPosition = string.Empty;
    private bool windowMaximized = false;
    private readonly bool argsPath = false;
    private string columnWidths = "100,100,200,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100";
    private SQLiteCommandBuilder? builder;
    //private delegate void MyDelegate(ComboBox myControl);
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
    private bool birthdayAutoShow = false;
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
    private bool isFilteredAddress = false;
    private bool isFilteredContact = false;
    private readonly string[] dataFields = ["Anrede", "Präfix", "Nachname", "Vorname", "Zwischenname", "Nickname",
        "Suffix", "Firma", "Straße", "PLZ", "Ort", "Land", "Betreff", "Grußformel", "Schlussformel", "Geburtstag",
        "Mail1", "Mail2", "Telefon1", "Telefon2", "Mobil", "Fax", "Internet", "Notizen", "Dokumente"];

    public FrmAdressen(string[] args)
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

        typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, addressDGV, [true]);
        typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, contactDGV, [true]);

        imageList.Images.Add(Properties.Resources.address_book);
        imageList.Images.Add(Properties.Resources.address_book_blue);
        imageList.Images.Add(Properties.Resources.universal24);
        imageList.Images.Add(Properties.Resources.inbox24);
        imageList.Images.Add(Properties.Resources.inboxdoc24);
        tabControl.ImageList = imageList; // Bilder zur Laufzeit aus Projekt-Ressourcen laden, vermeidet BinaryFormatter
        tabControl.TabPages[0].ImageIndex = 0;
        tabControl.TabPages[1].ImageIndex = 1;
        tabulation.TabPages[0].ImageIndex = 2;
        tabulation.TabPages[1].ImageIndex = 3;

        if (Utilities.IsInnoSetupValid(Path.GetDirectoryName(appPath)!))
        {
            xmlPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName, appName + ".xml");
            tokenDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName, "token.json");
            secretPath = Path.Combine(Path.GetDirectoryName(appPath) ?? string.Empty, "client_secret.json");
        }
        else
        {
            xmlPath = Path.ChangeExtension(appPath, ".xml");
            tokenDir = Path.Combine(AppContext.BaseDirectory, "token.json"); ;
            secretPath = Path.Combine(AppContext.BaseDirectory, "client_secret.json");
        }
        boysPath = Path.Combine(Path.GetDirectoryName(xmlPath) ?? string.Empty, "MännlicheVornamen.txt");
        girlPath = Path.Combine(Path.GetDirectoryName(xmlPath) ?? string.Empty, "WeiblicheVornamen.txt");
        if (File.Exists(girlPath))
        {
            foreach (var name in File.ReadLines(girlPath)) // Erste Zeile überspringen)
            {
                var trimmedName = name.Trim();
                if (!string.IsNullOrEmpty(trimmedName)) { nameGenderMap[trimmedName] = true; }
            }
        }
        if (File.Exists(boysPath))
        {
            foreach (var name in File.ReadLines(boysPath))
            {
                var trimmedName = name.Trim();
                if (!string.IsNullOrEmpty(trimmedName)) { nameGenderMap[trimmedName] = false; }
            }
        }

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

    private void FrmAdressen_Load(object sender, EventArgs e)
    {
        var hideColumns = string.Empty;
        if (File.Exists(xmlPath))
        {
            try
            {
                var xDocument = XDocument.Load(xmlPath);
                if (xDocument != null && xDocument.Root != null)
                {
                    foreach (var element in xDocument.Root.Descendants("Configuration"))
                    {
                        pDevice = element.Element("PrintDevice")?.Value ?? string.Empty;
                        pSource = element.Element("PrintSource")?.Value ?? string.Empty;
                        pLandscape = bool.TryParse(element.Element("PrintLandscape")?.Value, out var ls) ? ls : pLandscape;
                        pFormat = element.Element("PrintFormat")?.Value ?? string.Empty;
                        pFont = element.Element("PrintFont")?.Value ?? pFont;
                        pSenderSize = int.TryParse(element.Element("SenderFontsize")?.Value, out var ss) ? ss : pSenderSize;
                        pRecipSize = int.TryParse(element.Element("RecipientFontsize")?.Value, out var rs) ? rs : pRecipSize;
                        pSenderIndex = int.TryParse(element.Element("SenderIndex")?.Value, out var si) ? si : pSenderIndex;
                        pSenderLines1 = element.Element("SenderLines1")?.Value.Split('|');
                        pSenderLines2 = element.Element("SenderLines2")?.Value.Split('|');
                        pSenderLines3 = element.Element("SenderLines3")?.Value.Split('|');
                        pSenderLines4 = element.Element("SenderLines4")?.Value.Split('|');
                        pSenderLines5 = element.Element("SenderLines5")?.Value.Split('|');
                        pSenderLines6 = element.Element("SenderLines6")?.Value.Split('|');
                        pSenderPrint = bool.TryParse(element.Element("PrintSender")?.Value, out var sp) ? sp : pSenderPrint;
                        pRecipX = decimal.TryParse(element.Element("RecipientOffsetX")?.Value, out var rx) ? rx : pRecipX;
                        pRecipY = decimal.TryParse(element.Element("RecipientOffsetY")?.Value, out var ry) ? ry : pRecipY;
                        pSendX = decimal.TryParse(element.Element("SenderOffsetX")?.Value, out var sx) ? sx : pSendX;
                        pSendY = decimal.TryParse(element.Element("SenderOffsetY")?.Value, out var sy) ? sy : pSendY;
                        pRecipBold = bool.TryParse(element.Element("PrintRecipientBold")?.Value, out var rb) ? rb : pRecipBold;
                        pSendBold = bool.TryParse(element.Element("PrintSenderBold")?.Value, out var sb) ? sb : pSendBold;
                        pSalutation = bool.TryParse(element.Element("PrintRecipientSalutation")?.Value, out var st) ? st : pSalutation;
                        pCountry = bool.TryParse(element.Element("PrintRecipientCountry")?.Value, out var pc) ? pc : pCountry;
                        sAskBeforeDelete = bool.TryParse(element.Element("AskBeforeDelete")?.Value, out var ad) ? ad : sAskBeforeDelete;
                        sColorScheme = element.Element("ColorScheme")?.Value ?? "blue";
                        sContactsAutoload = bool.TryParse(element.Element("ContactsAutoload")?.Value, out var ca) ? ca : sContactsAutoload;
                        sAskBeforeSaveSQL = bool.TryParse(element.Element("AskBeforeSaveSQL")?.Value, out var ab) ? ab : sAskBeforeSaveSQL;
                        sReloadRecent = bool.TryParse(element.Element("ReloadRecent")?.Value, out var rr) ? rr : sReloadRecent;
                        sNoAutoload = bool.TryParse(element.Element("NoAutoload")?.Value, out var nf) ? nf : sNoAutoload;
                        sStandardFile = element.Element("StandardFile")?.Value ?? string.Empty;
                        sDailyBackup = bool.TryParse(element.Element("DailyBackup")?.Value, out var db) ? db : sDailyBackup;
                        sWatchFolder = bool.TryParse(element.Element("WatchFolder")?.Value, out var wf) ? wf : sWatchFolder;
                        sBackupSuccess = bool.TryParse(element.Element("BackupSuccess")?.Value, out var bs) ? bs : sBackupSuccess;
                        sSuccessDuration = decimal.TryParse(element.Element("SuccessDuration")?.Value, out var su) ? su : sSuccessDuration;
                        sBackupDirectory = element.Element("BackupDirectory")?.Value ?? string.Empty;
                        sLetterDirectory = element.Element("DocumentFolder")?.Value ?? string.Empty;
                        sDatabaseFolder = element.Element("DatabaseFolder")?.Value ?? string.Empty;
                        indexCopyPattern = int.TryParse(element.Element("CopyPatternIndex")?.Value, out var cp) ? si : indexCopyPattern;
                        copyPattern1 = string.IsNullOrEmpty(element.Element("CopyPattern1")?.Value) ? copyPattern1 : element.Element("CopyPattern1")?.Value.Split('|');
                        copyPattern2 = string.IsNullOrEmpty(element.Element("CopyPattern2")?.Value) ? copyPattern2 : element.Element("CopyPattern2")?.Value.Split('|');
                        copyPattern3 = string.IsNullOrEmpty(element.Element("CopyPattern3")?.Value) ? copyPattern3 : element.Element("CopyPattern3")?.Value.Split('|');
                        copyPattern4 = string.IsNullOrEmpty(element.Element("CopyPattern4")?.Value) ? copyPattern4 : element.Element("CopyPattern4")?.Value.Split('|');
                        copyPattern5 = string.IsNullOrEmpty(element.Element("CopyPattern5")?.Value) ? copyPattern5 : element.Element("CopyPattern5")?.Value.Split('|');
                        copyPattern6 = string.IsNullOrEmpty(element.Element("CopyPattern6")?.Value) ? copyPattern6 : element.Element("CopyPattern6")?.Value.Split('|');
                        hideColumns = element.Element("HideColumns")?.Value ?? string.Empty;
                        splitterPosition = int.TryParse(element.Element("SplitterPosition")?.Value, out var sr) ? sr : splitterPosition;
                        windowMaximized = bool.TryParse(element.Element("WindowMaximized")?.Value, out var wm) ? wm : windowMaximized;
                        windowPosition = element.Element("WindowPosition")?.Value ?? string.Empty;
                        var newWidths = element.Element("ColumnWidths")?.Value;
                        if (!string.IsNullOrWhiteSpace(newWidths)) { columnWidths = newWidths; } // ansonsten den Standardwert behalten 
                        birthdayRemindLimit = int.TryParse(element.Element("BirthdayRemindLimit")?.Value, out var br) ? br : birthdayRemindLimit;
                        birthdayAutoShow = bool.TryParse(element.Element("BirthdayAutoShow")?.Value, out var ba) ? ba : birthdayAutoShow;
                        recentFiles = [.. (element.Element("RecentFiles")?.Value ?? "").Split('|', StringSplitOptions.RemoveEmptyEntries)];
                        databaseFilePath = argsPath ? databaseFilePath : recentFiles.Count > 0 ? recentFiles[0] : string.Empty;
                        sWordProcProg = bool.TryParse(element.Element("WordProcessorProgram")?.Value, out var wp) ? wp : null;
                    }

                }
            }
            catch (XmlException ex)
            {
                Utilities.StartFile(Handle, xmlPath);
                Utilities.ErrorMsgTaskDlg(Handle, "FrmAdressen_Load: " + ex.GetType().ToString(), ex.Message);
            }
        }
        else { Directory.CreateDirectory(Path.GetDirectoryName(xmlPath)!); } // If the folder exists already, the line will be ignored.     

        hideColumnArr = string.IsNullOrEmpty(hideColumns) ? hideColumnStd : Utilities.FromBase64String(hideColumns) ?? hideColumnStd;

        if (windowMaximized) { WindowState = FormWindowState.Maximized; }
        else if (!string.IsNullOrEmpty(windowPosition))
        {
            var coords = windowPosition.Split(',');
            var width = int.Parse(coords[2]);
            var height = int.Parse(coords[3]);
            var xPos = int.Parse(coords[0]);
            var yPos = int.Parse(coords[1]);
            var screen = Screen.PrimaryScreen!.WorkingArea;                         // Arbeitsbereich ohne Taskleisten, angedockte Fenster und angedockte Symbolleisten
            xPos = xPos > screen.Width ? screen.Width - width : xPos;               // xPos für den Fall korrigieren, dass der rechte Rand nicht mehr zu sehen ist
            xPos = xPos < 0 ? 0 : xPos;                                             // Minuswerte korrieren
            yPos = yPos + height > screen.Height ? screen.Height - height : yPos;            // yPos für den Fall korrigieren, dass der untere Rand nicht mehr zu sehen ist             
            yPos = yPos < 0 ? 0 : yPos;                                             // Minuswerte korrigieren
            width = xPos + width > screen.Width ? screen.Width - xPos : width;      // Breite für den Fall korrigieren, dass sie zu groß geworden sind
            height = yPos + height > screen.Height ? screen.Height - yPos : height; // Höhe für den Fall korrigieren, dass sie zu groß geworden sind
            Location = new Point(xPos, yPos);
            Size = new Size(width, height);
        }
        NativeMethods.SendMessage(searchTSTextBox.TextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 4 << 16);
        NativeMethods.SendMessage(searchTSTextBox.TextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 4);
        NativeMethods.SendMessage(tbNotizen.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        NativeMethods.SendMessage(tbNotizen.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);
        NativeMethods.SendMessage(maskedTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        NativeMethods.SendMessage(maskedTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);
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
    }

    private void SaveConfiguration()
    {
        var point = Location; // Fensterposition speichern
        var size = Size;
        if (WindowState != FormWindowState.Normal)
        {
            point = RestoreBounds.Location;
            size = RestoreBounds.Size;
        }
        XElement element = new("Configuration");
        element.Add(new XElement("PrintDevice", pDevice));
        element.Add(new XElement("PrintSource", pSource));
        element.Add(new XElement("PrintLandscape", pLandscape.ToString()));
        element.Add(new XElement("PrintFormat", pFormat));
        element.Add(new XElement("PrintFont", pFont));
        element.Add(new XElement("SenderFontsize", pSenderSize));
        element.Add(new XElement("RecipientFontsize", pRecipSize));
        element.Add(new XElement("SenderIndex", pSenderIndex));
        element.Add(new XElement("SenderLines1", pSenderLines1 != null ? string.Join('|', pSenderLines1) : string.Empty));
        element.Add(new XElement("SenderLines2", pSenderLines2 != null ? string.Join('|', pSenderLines2) : string.Empty));
        element.Add(new XElement("SenderLines3", pSenderLines3 != null ? string.Join('|', pSenderLines3) : string.Empty));
        element.Add(new XElement("SenderLines4", pSenderLines4 != null ? string.Join('|', pSenderLines4) : string.Empty));
        element.Add(new XElement("SenderLines5", pSenderLines5 != null ? string.Join('|', pSenderLines5) : string.Empty));
        element.Add(new XElement("SenderLines6", pSenderLines6 != null ? string.Join('|', pSenderLines6) : string.Empty));
        element.Add(new XElement("PrintSender", pSenderPrint.ToString()));
        element.Add(new XElement("RecipientOffsetX", pRecipX));
        element.Add(new XElement("RecipientOffsetY", pRecipY));
        element.Add(new XElement("SenderOffsetX", pSendX));
        element.Add(new XElement("SenderOffsetY", pSendY));
        element.Add(new XElement("PrintRecipientBold", pRecipBold.ToString()));
        element.Add(new XElement("PrintSenderBold", pSendBold.ToString()));
        element.Add(new XElement("PrintRecipientSalutation", pSalutation.ToString()));
        element.Add(new XElement("PrintRecipientCountry", pCountry.ToString()));
        element.Add(new XElement("AskBeforeDelete", sAskBeforeDelete.ToString()));
        element.Add(new XElement("ColorScheme", sColorScheme));
        element.Add(new XElement("ContactsAutoload", sContactsAutoload.ToString()));
        element.Add(new XElement("AskBeforeSaveSQL", sAskBeforeSaveSQL.ToString()));
        element.Add(new XElement("RecentFiles", string.Join('|', recentFiles)));
        element.Add(new XElement("ReloadRecent", sReloadRecent.ToString()));
        element.Add(new XElement("NoAutoload", sNoAutoload.ToString()));
        element.Add(new XElement("StandardFile", sStandardFile));
        element.Add(new XElement("DailyBackup", sDailyBackup.ToString()));
        element.Add(new XElement("WatchFolder", sWatchFolder.ToString()));
        element.Add(new XElement("BackupSuccess", sBackupSuccess.ToString()));
        element.Add(new XElement("SuccessDuration", sSuccessDuration));
        element.Add(new XElement("BackupDirectory", sBackupDirectory));
        element.Add(new XElement("DocumentFolder", sLetterDirectory));
        element.Add(new XElement("DatabaseFolder", sDatabaseFolder));
        element.Add(new XElement("CopyPatternIndex", indexCopyPattern));
        element.Add(new XElement("CopyPattern1", copyPattern1 != null ? string.Join('|', copyPattern1) : string.Empty));
        element.Add(new XElement("CopyPattern2", copyPattern2 != null ? string.Join('|', copyPattern2) : string.Empty));
        element.Add(new XElement("CopyPattern3", copyPattern3 != null ? string.Join('|', copyPattern3) : string.Empty));
        element.Add(new XElement("CopyPattern4", copyPattern4 != null ? string.Join('|', copyPattern4) : string.Empty));
        element.Add(new XElement("CopyPattern5", copyPattern5 != null ? string.Join('|', copyPattern5) : string.Empty));
        element.Add(new XElement("CopyPattern6", copyPattern6 != null ? string.Join('|', copyPattern6) : string.Empty));
        element.Add(new XElement("DefaultSplitter", splitterAutomaticToolStripMenuItem.Checked.ToString()));
        element.Add(new XElement("SplitterPosition", splitContainer.SplitterDistance));
        element.Add(new XElement("HideColumns", Utilities.BoolArray2Base64String(hideColumnArr)));
        element.Add(new XElement("WindowMaximized", WindowState == FormWindowState.Maximized));
        element.Add(new XElement("WindowPosition", string.Join(",", point.X, point.Y, size.Width, size.Height)));
        element.Add(new XElement("BirthdayRemindLimit", birthdayRemindLimit));
        element.Add(new XElement("BirthdayAutoShow", birthdayAutoShow.ToString()));
        if (tabControl.SelectedTab == contactTabPage) { element.Add(new XElement("ColumnWidths", Utilities.GetColumnWidths(contactDGV))); }
        else if (tabControl.SelectedTab == addressTabPage) { element.Add(new XElement("ColumnWidths", Utilities.GetColumnWidths(addressDGV))); }
        element.Add(new XElement("WordProcessorProgram", sWordProcProg.HasValue ? sWordProcProg.ToString() : string.Empty));
        XDocument xDocument = new(new XElement(appName, element));
        xDocument.Save(xmlPath);
    }

    private void FrmAdressen_Shown(object sender, EventArgs e)
    {
        if (birthdayAutoShow) { BirthdaysToolStripMenuItem_Click(birthdayAutoShow, EventArgs.Empty); }

        if (sContactsAutoload) { backgroundWorker.RunWorkerAsync(); }

        splitContainer.SplitterDistance = splitterPosition;
        flexiTSStatusLabel.Width = 244 + splitContainer.SplitterDistance - 536;
        if (addressDGV.DataSource != null) { searchTSTextBox.TextBox.Focus(); }
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
            liteConnection = new SQLiteConnection((string?)$"Data Source={databaseFilePath};FailIfMissing=True"); //Create a SqliteConnection object called Connection
            liteConnection.Open(); //open a connection to the database
            dataAdapter = new SQLiteDataAdapter("SELECT * FROM Adressen", liteConnection);  // Create a SQLiteDataAdapter to retrieve data
            dataTable = new DataTable(); // Create a DataTable to hold the data
            dataAdapter.Fill(dataTable);           // Use the Fill method to retrieve data into the DataTable
            var sortedRows = from row in dataTable.AsEnumerable() orderby row.Field<string>("Vorname") ascending orderby row.Field<string>("Nachname") ascending select row; // alphabetisch 
            //var sortedRows = from row in dataTable.AsEnumerable() orderby row.Field<long>("Id") ascending select row; // alphabetisch 
            using (var sortedDT = dataTable.Clone())
            {
                foreach (var row in sortedRows) { sortedDT.ImportRow(row); }
                dataTable = sortedDT;
            }
            addressDGV.SuspendLayout();
            addressDGV.DataSource = dataTable;
            if (!string.IsNullOrEmpty(columnWidths)) { Utilities.SetColumnWidths(columnWidths, addressDGV); }
            foreach (DataGridViewColumn column in addressDGV.Columns) { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
            if (addressDGV.Rows.Count > 0)
            {
                //dataTable = dataTable.AsEnumerable().Where(row => row.ItemArray.SkipLast(1).Any(field => field is not DBNull)).CopyToDataTable(); 
                var emptyRows = from DataGridViewRow row in addressDGV.Rows.Cast<DataGridViewRow>()
                                       .Where(row => row.Cells.Cast<DataGridViewCell>().SkipLast(2).All(cell => cell.Value == null || string.IsNullOrEmpty(cell.Value.ToString())))
                                select row;
                foreach (var emptyRow in emptyRows) // //if (emptyRows.Any())
                {
                    //MessageBox.Show(emptyRow.Cells[^1].Value.ToString());
                    if (emptyRow.DataBoundItem is DataRowView drv) { drv.Row.Delete(); }
                    using (var changes = dataTable.GetChanges(DataRowState.Deleted))
                    {
                        if (changes != null)
                        {
                            builder = new SQLiteCommandBuilder(dataAdapter); // Automatically generates single-table commands used to reconcile changes
                            dataAdapter.Update(changes);
                        }
                    }

                    dataTable.AcceptChanges();
                }
            }
            liteConnection.Close();
            addressDGV.ResumeLayout(false);
            cbAnrede.Items.Clear();
            cbPräfix.Items.Clear();
            cbPLZ.Items.Clear();
            cbOrt.Items.Clear();
            cbLand.Items.Clear();
            cbSchlussformel.Items.Clear();
            cbAnrede.Items.AddRange([.. dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Anrede")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbPräfix.Items.AddRange([.. dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Präfix")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbPLZ.Items.AddRange([.. dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("PLZ")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbOrt.Items.AddRange([.. dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Ort")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbLand.Items.AddRange([.. dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Land")!).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct()]);
            cbSchlussformel.Items.AddRange([.. dataTable.Rows.Cast<DataRow>().Select(row => row.Field<string>("Schlussformel")!).Where(value => !string.IsNullOrEmpty(value)).Distinct()]);
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
            AdressEditFields(dataTable.Rows.Count > 0 ? 0 : -1);
            searchTSTextBox.Focus(); //MessageBox.Show(string.Join(Environment.NewLine, [.. dataTable.Columns.Cast<DataColumn>().Select(x => x.ColumnName)]));
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "ConnectSQLDatabase: " + ex.GetType().ToString(), ex.Message); }
    }

    //private void DataTable_RowEvents(object sender, DataRowChangeEventArgs e)
    //{
    //    //Console.Beep();
    //    saveTSButton.Enabled = true;
    //}

    //private void DataTable_TableNewRow(object sender, DataTableNewRowEventArgs e) => saveTSButton.Enabled = true;
    private List<string> FillComboBoxFromAddressDataTable()
    {
        HashSet<string> uniqueValues = []; // HashSet, um Duplikate zu vermeiden
        foreach (DataRow row in dataTable!.Rows)
        {
            var foo = row.Field<string>("Grußformel");
            if (!string.IsNullOrEmpty(foo))
            {
                var bar = row.Field<string>("Nachname");
                if (!string.IsNullOrEmpty(bar)) { foo = foo.Replace(bar, "#nachname"); }
                var baz = row.Field<string>("Zwischenname");
                if (!string.IsNullOrEmpty(baz)) { foo = foo.Replace(baz, "#zwischenname"); }
                var qux = row.Field<string>("Vorname");
                if (!string.IsNullOrEmpty(qux))
                {
                    var array = qux.Split(' ');
                    for (var i = 0; i < array.Length; i++)
                    {
                        var wort = array[i];
                        if (string.IsNullOrEmpty(wort)) { continue; }
                        foo = foo.Replace(wort, "#vorname" + i);
                    }
                }
                uniqueValues.Add(foo);
                //uniqueValues.Add(Regex.Replace(foo, @"#\w+", "")); // restliche Ersetzungen (Wörter die mit # beginnen) entfernen
            }
        }
        //List<string> sortedList = [.. uniqueValues];
        //sortedList.Sort();

        //MessageBox.Show(string.Join(Environment.NewLine, uniqueValues));
        return [.. uniqueValues];
    }

    private List<string> FillComboBoxFromContactDGV()
    {
        HashSet<string> uniqueValues = []; // Duplikate vermeiden
        foreach (DataGridViewRow row in contactDGV.Rows)
        {
            if (row.IsNewRow) { continue; }
            var foo = row.Cells["Grußformel"].Value?.ToString();
            if (!string.IsNullOrEmpty(foo))
            {
                var bar = row.Cells["Nachname"].Value?.ToString();
                if (!string.IsNullOrEmpty(bar)) { foo = foo.Replace(bar, "#nachname"); }
                var baz = row.Cells["Zwischenname"].Value?.ToString();
                if (!string.IsNullOrEmpty(baz)) { foo = foo.Replace(baz, "#zwischenname"); }
                var qux = row.Cells["Vorname"].Value?.ToString();
                if (!string.IsNullOrEmpty(qux))
                {
                    var array = qux.Split(' ');
                    for (var i = 0; i < array.Length; i++)
                    {
                        var wort = array[i];
                        if (string.IsNullOrEmpty(wort)) { continue; }
                        foo = foo.Replace(wort, "#vorname" + i);
                    }
                }
                uniqueValues.Add(Regex.Replace(foo, @"#\w+", ""));
            }
        }
        return [.. uniqueValues];
    }

    private void SaveSQLDatabase(bool closeDB = false, bool askNever = false) // askNever: wenn auch Speichern-Icon geklickt wird, soll in jedem Fall keine Abfrage erfolgen
    {
        if (dataTable == null || dataTable.Rows.Count <= 0) { return; } // Verhindert nebenbei, dass alle Rows gelöscht werden bzw. wenn, dann wird es nicht gespeichert.

        //dataTable.RowChanged -= DataTable_RowEvents;
        //dataTable.RowDeleted -= DataTable_RowEvents;
        //dataTable.TableNewRow -= DataTable_TableNewRow;

        addressDGV.EndEdit();  // Schritt 1: Beenden Sie alle laufenden Bearbeitungen im DataGridView
        if (addressDGV.DataSource != null && BindingContext != null) { BindingContext[addressDGV.DataSource].EndCurrentEdit(); } // Zusätzlicher Schutz, um Änderungen zu übernehmen
        using var changes = dataTable?.GetChanges(DataRowState.Added | DataRowState.Modified | DataRowState.Deleted);
        if (dataTable != null && changes != null && dataAdapter != null)  // //MessageBox.Show(changes.Rows.Count.ToString());
        {
            if (tabControl.SelectedTab != addressTabPage) { tabControl.SelectTab(addressTabPage); }
            if (!askNever && sAskBeforeSaveSQL && !Utilities.YesNo_TaskDialog(Handle, appName, "Möchten Sie die Änderungen speichern?",
                changes.Rows.Count == 1 ? "An einer Adresse wurden Änderungen vorgenommen." : $"Änderungen wurden an {changes.Rows.Count} Adressen vorgenommen.", TaskDialogIcon.ShieldGrayBar)) { return; }
            {
                try
                {
                    liteConnection = new SQLiteConnection((string?)$"Data Source={databaseFilePath};FailIfMissing=True"); //Create a SqliteConnection object called Connection
                    liteConnection.Open(); //open a connection to the database
                    builder = new SQLiteCommandBuilder(dataAdapter); // Automatically generates single-table commands used to reconcile changes
                    dataAdapter.Update(changes);  // Automatically generates single-table commands used to reconcile changes
                    dataTable.AcceptChanges(); // Änderungen markieren als übernommen
                    saveTSButton.Enabled = false; // Speichern-Button deaktivieren   addressChanges = 
                    flexiTSStatusLabel.Text = $"Letztes Speichern: {DateTime.Now:HH:mm:ss}";
                }
                catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "SaveSQLDatabase: " + ex.GetType().ToString(), ex.Message); }
                finally { liteConnection?.Close(); }
            }
            if (sDailyBackup && File.Exists(Utilities.CorrectUNC(databaseFilePath)) && Directory.Exists(sBackupDirectory)) { Utilities.DailyBackup(Utilities.CorrectUNC(databaseFilePath), sBackupDirectory, sBackupSuccess, sSuccessDuration, Handle); }
        }
        if (closeDB)
        {
            dataAdapter?.Dispose();
            if (dataTable != null)
            {
                dataTable.Dispose();
                dataTable = null;
            }
            if (liteConnection != null)
            {
                try { liteConnection.Close(); } catch { }
                liteConnection.Dispose();
                liteConnection = null;
            }
            addressDGV.DataSource = null;
            addressDGV.Rows.Clear();
            AdressEditFields(-1);
            duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled
                = deleteTSButton.Enabled = newToolStripMenuItem.Enabled = newTSButton.Enabled = duplicateToolStripMenuItem.Enabled = copyTSButton.Enabled = wordTSButton.Enabled
                = envelopeTSButton.Enabled = false;
            copyToOtherDGVTSMenuItem.Enabled = false;
            flexiTSStatusLabel.Text = string.Empty;
            searchTSTextBox.TextBox.Clear();
            tsClearLabel.Visible = false;
        }
        else
        {
            foreach (var entry in changedAddressData) { originalAddressData[entry.Key] = entry.Value; } // Überschreibe den Wert, wenn Schlüssel existiert oder füge einen neuen hinzu
            changedAddressData.Clear(); // Leeren des Dictionaries nach der Aktualisierung
        }
    }

    private async void OpenToolStripMenuItem_Click(object? sender, EventArgs? e)
    { //openFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp()) { await CreateContactAsync(); }
            //var ressource = CheckContactDataChange() ?? string.Empty;
            //if (!string.IsNullOrEmpty(ressource)) { ShowMultiPageTaskDialog(ressource); }
            if (CheckContactDataChange()) { ShowMultiPageTaskDialog(); }
        }
        var fileName = Path.GetFileName(Utilities.CorrectUNC(databaseFilePath)) ?? string.Empty;
        var dirName = Path.GetDirectoryName(Utilities.CorrectUNC(databaseFilePath)) ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(fileName)) { openFileDialog.FileName = fileName; }
        openFileDialog.InitialDirectory = !string.IsNullOrEmpty(sDatabaseFolder) && Directory.Exists(sDatabaseFolder) ? sDatabaseFolder : !string.IsNullOrWhiteSpace(dirName) ? dirName : null;
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            if (dataAdapter != null) { SaveSQLDatabase(true); }
            ConnectSQLDatabase(openFileDialog.FileName);
            ignoreSearchChange = true;
            searchTSTextBox.TextBox.Clear();
            ignoreSearchChange = false;
        }
    }

    private void ExitToolStripMenuItem_Click(object? sender, EventArgs? e)
    {
        if (dataAdapter != null) { SaveSQLDatabase(true); }
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
                AdressEditFields(prevSelectedAddressRowIndex);
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

    private void AdressEditFields(int rowIndex) // rowIndex = -1 => ClearFields
    {
        tabulation.SelectedTab = tabPageDetail;
        try
        {
            ignoreTextChange = true; // verhindert, dass TextChanged
            foreach (var (ctrl, colText) in dictEditField) { ctrl.Text = rowIndex < 0 ? "" : addressDGV.Rows[rowIndex].Cells[colText]?.Value?.ToString() ?? ""; }

            cbGrußformel.Items.Clear();
            if (rowIndex >= 0) { ErzeugeGrußformeln(); }
            if (rowIndex >= 0 && DateTime.TryParse(addressDGV.Rows[rowIndex].Cells["Geburtstag"]?.Value?.ToString(), out var date))
            {
                maskedTextBox.Text = date.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("de-DE"));
                AgeLabel_SetText(date);
            }
            else
            {
                AgeLabel_DeleteText();
                maskedTextBox.Text = string.Empty;
            }

            tbNotizen.Text = rowIndex < 0 ? "" : addressDGV.Rows[rowIndex].Cells["Notizen"]?.Value?.ToString();

            dokuListView.Items.Clear();
            if (rowIndex >= 0 && addressDGV.Rows[rowIndex].DataBoundItem is DataRowView rowView)
            {
                var json = rowView.Row["Dokumente"].ToString();
                if (!string.IsNullOrEmpty(json))
                {
                    var dateipfade = JsonSerializer.Deserialize<List<string>>(json);
                    if (dateipfade != null)
                    {
                        foreach (var pfad in dateipfade)
                        {
                            Add2dokuListView(pfad, false);
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
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "AdressEditFields: " + ex.GetType().ToString(), ex.Message); }
        finally { ignoreTextChange = false; } // TextChanged wieder aktivieren
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
        if (addressDGV.DataSource != null)
        {
            for (var i = 0; i < addressDGV.Columns.Count; i++) { addressDGV.Columns[i].Visible = !hideColumnArr[i]; }
            Text = appName + " – " + (string?)(string.IsNullOrEmpty(databaseFilePath) ? "unbenannt" : Utilities.CorrectUNC(databaseFilePath));  // Workaround for UNC-Path
        }
        else { Text = appName; }
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
                    isFilteredAddress = false;
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
                if (string.IsNullOrWhiteSpace(normalizedSearchTerm))
                {
                    FilterContactDGV(row => true);
                    isFilteredContact = false;
                }
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
        if (tabControl.SelectedTab == addressTabPage && dataTable != null)
        {
            if (dataAdapter != null)
            {
                var rowIndex = addressDGV.SelectedRows.Count > 0 ? addressDGV.SelectedRows[0].Index : -1;
                SaveSQLDatabase(false, true);
                if (rowIndex >= 0 && addressDGV.Rows[rowIndex] != null)
                {
                    addressDGV.Rows[rowIndex].Selected = true;
                    addressDGV.FirstDisplayedScrollingRowIndex = rowIndex;
                }
            }
            //if (addressDGV.SelectedRows.Count > 0) { AdressEditFields(addressDGV.SelectedRows[0].Index); } // z.B. nach Neuanlegen einer Adresse
        }
        else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0)
        {
            if (contactNewRowIndex >= 0 && contactDGV.SelectedRows[0].Index == contactNewRowIndex && CheckNewContactTidyUp()) { await CreateContactAsync(); }
            //var ressource = CheckContactDataChange() ?? string.Empty;
            //if (!string.IsNullOrEmpty(ressource)) { ShowMultiPageTaskDialog(ressource); }
            if (CheckContactDataChange()) { ShowMultiPageTaskDialog(); }
        }
        else { Console.Beep(); }
    }

    private bool CheckContactDataChange()  //originalContactData wird in ContactEditFields() gesetzt 
    {
        if (originalContactData == null || originalContactData.Count == 0 || contactDGV == null || prevSelectedContactRowIndex < 0) { return false; }
        changedContactData.Clear();
        foreach (var cell in contactDGV.Rows[prevSelectedContactRowIndex].Cells.Cast<DataGridViewCell>().SkipLast(1).Where(cell => !Equals(originalContactData[cell.OwningColumn.Name], cell.Value)))
        {
            changedContactData[cell.OwningColumn.Name] = cell.Value?.ToString() ?? string.Empty;
        }
        if (changedContactData.Count > 0) { return true; }
        return false;
    }

    //private bool CheckAddressDataChange()
    //{
    //    if (originalAddressData == null || originalAddressData.Count == 0 || addressDGV == null || prevSelectedAddressRowIndex < 0) { return false; }
    //    changedAddressData.Clear();
    //    foreach (var cell in addressDGV.Rows[prevSelectedAddressRowIndex].Cells.Cast<DataGridViewCell>().SkipLast(1).Where(cell => !Equals(originalAddressData[cell.OwningColumn.Name], cell.Value)))
    //    {
    //        changedAddressData[cell.OwningColumn.Name] = cell.Value?.ToString() ?? string.Empty;
    //    }
    //    if (changedAddressData.Count > 0) { return true; }
    //    else if (addressDGV.Rows[prevSelectedAddressRowIndex].DataBoundItem is DataRowView dataBoundItem)
    //    {
    //        dataBoundItem.Row.AcceptChanges();  // keine Änderung, RowState auf Unchanged setzen 
    //        return false;
    //    }
    //    return true; // false wäre vielleicht auch richtig
    //}

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
        var initialButtonYes = TaskDialogButton.Yes;
        var initialButtonNo = TaskDialogButton.No;
        initialButtonYes.AllowCloseDialog = false; // don't close the dialog when this button is clicked
        var initialPage = new TaskDialogPage()
        {
            Caption = "Google Kontakte",
            Heading = "Möchten Sie die Änderungen speichern?",
            Text = message,
            Icon = TaskDialogIcon.ShieldBlueBar,
            AllowCancel = true,
            SizeToContent = true,
            //Expander = new TaskDialogExpander()
            //{
            //    Text = "Beim Speichern werden vorhandene Daten überschrieben.\nUnter bestimmten Umständen droht Datenverlust!",
            //    Position = TaskDialogExpanderPosition.AfterFootnote
            //},
            Buttons = { initialButtonNo, initialButtonYes },
            DefaultButton = initialButtonNo
        };

        var inProgressCloseButton = TaskDialogButton.Close;
        inProgressCloseButton.Enabled = false;
        var progressPage = new TaskDialogPage()
        {
            Caption = appName,
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
        NativeMethods.ShowScrollBar(tbNotizen.Handle, 1, TextRenderer.MeasureText(tbNotizen.Text, tbNotizen.Font, new System.Drawing.Size(tbNotizen.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue),
        TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height > tbNotizen.Height);
        if (ignoreTextChange) { return; } // verhindert, dass TextChanged bei AdressEditFields aufgerufen wird  
        if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0 && addressDGV.SelectedRows[0].Cells["Notizen"].Value.ToString() != tbNotizen.Text.Trim())
        {
            addressDGV.SelectedRows[0].Cells["Notizen"].Value = tbNotizen.Text;
        }
        else if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0) { contactDGV.SelectedRows[0].Cells["Notizen"].Value = tbNotizen.Text; }
        CheckSaveButton();
    }

    private void TbNotizen_SizeChanged(object sender, EventArgs e) => NativeMethods.ShowScrollBar(tbNotizen.Handle, 1, TextRenderer.MeasureText(tbNotizen.Text, tbNotizen.Font, new System.Drawing.Size(tbNotizen.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height > tbNotizen.Height);

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
            if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0) { addressDGV.SelectedRows[0].Cells[colName].Value = ctrl.Text.Trim(); }
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
                //var ressource = CheckContactDataChange() ?? string.Empty;
                //if (!string.IsNullOrEmpty(ressource)) { ShowMultiPageTaskDialog(ressource); }
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
        else if (tabControl.SelectedTab == addressTabPage && dataTable != null)
        {
            var row = dataTable.NewRow();
            row["Id"] = dataTable.Rows.Count > 0 ? dataTable.AsEnumerable().Max(r => r.Field<long>("Id")) + 1 : 1;
            dataTable.Rows.Add(row);
            if (addressDGV.RowCount > 0)
            {
                addressDGV.Rows[^1].Selected = true;
                //addressNewRowIndex = addressDGV.Rows[^1].Index;
                addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.Rows[^1].Index;
                AdressEditFields(addressDGV.Rows[^1].Index);
                cbAnrede.Focus();
            }
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
        else if (tabControl.SelectedTab == addressTabPage && dataTable != null && addressDGV.SelectedRows[0] != null)
        {
            var newRow = dataTable.NewRow();
            newRow["Id"] = dataTable.Rows.Count > 0 ? dataTable.AsEnumerable().Where(r => r.RowState != DataRowState.Deleted).Max(r => r.Field<long>("Id")) + 1 : 1;
            if (addressDGV.SelectedRows[0].DataBoundItem is DataRowView dataBoundItem) { newRow.ItemArray = dataBoundItem.Row.ItemArray; }
            else { return; }
            dataTable.Rows.Add(newRow);
            addressDGV.Rows[^1].Selected = true;
            addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.Rows[^1].Index;
            cbAnrede.Focus();
        }
        else { Console.Beep(); }
    }

    private void CopyToOtherDGVMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0 && dataTable != null)
        {
            using var selectedRow = contactDGV.SelectedRows[0];
            var newRow = dataTable.NewRow();
            //newRow["Id"] = dataTable.Rows.Count > 0 ? dataTable.AsEnumerable().Max(r => r.Field<long>("Id")) + 1 : 1;

            newRow["Id"] = dataTable.Rows.Count > 0 ? dataTable.AsEnumerable().Where(r => r.RowState != DataRowState.Deleted).Max(r => r.Field<long>("Id")) + 1 : 1;

            selectedRow.Cells.Cast<DataGridViewCell>().SkipLast(1).ToList().ForEach(cell => newRow[cell.ColumnIndex] = cell.Value);
            tabControl.SelectedTab = addressTabPage;
            searchTSTextBox.TextBox.Clear();
            dataTable.Rows.Add(newRow);
            if (addressDGV.RowCount > 0)
            {
                addressDGV.Rows[^1].Selected = true;
                addressDGV.FirstDisplayedScrollingRowIndex = addressDGV.Rows[^1].Index;
                AdressEditFields(addressDGV.Rows[^1].Index);
                cbAnrede.Focus();
            }
        }
        else if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0 && contactDGV != null)
        {
            using var selectedRow = addressDGV.SelectedRows[0];
            contactNewRowIndex = contactDGV.Rows.Add();
            selectedRow.Cells.Cast<DataGridViewCell>().SkipLast(1).ToList().ForEach(cell => contactDGV.Rows[contactNewRowIndex].Cells[cell.ColumnIndex].Value = cell.Value);
            tabControl.SelectedTab = contactTabPage;
            searchTSTextBox.TextBox.Clear();
            contactDGV.Rows[contactNewRowIndex].Selected = true;
            contactDGV.FirstDisplayedScrollingRowIndex = contactNewRowIndex;
            ContactEditFields(contactNewRowIndex);
            cbAnrede.Focus();
        }
        else { Console.Beep(); } // für Tastenkombination Strg+K
    }

    private void DeleteTSButton_Click(object sender, EventArgs e)
    {
        var delete = false;

        if (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count > 0 && contactDGV.SelectedRows[0] != null)
        {
            var row = contactDGV.SelectedRows[0];
            (sAskBeforeDelete, delete) = Utilities.AskBeforeDeleteTaskDlg(Handle, row, sAskBeforeDelete, false); // false = keine Verification (ask before delete)  
            if (delete && row != null) { DeleteGoogleContact(row.Index); }
        }

        else if (addressDGV.SelectedRows.Count > 0 && !addressDGV.SelectedRows[0].IsNewRow && dataTable != null) // && !addressDGV.SelectedRows[0].IsNewRow
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
                if (addressDGV.Rows.Count == 0) // Wenn keine Zeilen mehr vorhanden sind, Auswahl löschen und Felder leeren.
                {
                    ignoreSearchChange = true;
                    searchTextBox.Clear();
                    ignoreSearchChange = false; 
                    AdressEditFields(-1);
                    return;
                }
                if (searchTSTextBox.TextLength > 0) { SearchTSTextBox_TextChanged(null!, null!); } // Schritt 3: Den Filter neu anwenden, wenn ein Suchtext vorhanden ist.
                var nextSelectedIndex = -1; // Schritt 4: Die vorherige sichtbare Zeile finden und auswählen.
                for (var i = indexToDelete; i >= 0; i--)  // Wir starten die Suche beim Index der gelöschten Zeile und gehen rückwärts.
                {
                    if (addressDGV.Rows[i].Visible)
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
        if (dataTable != null) { SaveSQLDatabase(true); }

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

    private void AboutToolStripMenuItem_Click(object sender, EventArgs e) => Utilities.HelpMsgTaskDlg(Handle, appName, Icon);

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


    //private void ErzeugeGrußformeln()
    //{
    //    cbGrußformel.Items.Clear();

    //    (string, string Text)[] pt = [("#vorname", tbVorname.Text), ("#nickname", tbNickname.Text), ("#nachname", tbNachname.Text)]; // Platzhalter ohne #titel
    //    var gender = GetGender(tbVorname.Text);
    //    cbGrußformel.Items.AddRange(
    //        [
    //            .. (gender == false ? männlichGruß
    //            : gender == true ? weiblichGruß
    //            : weiblichGruß.Concat(männlichGruß))
    //        .SelectMany(s =>
    //            //new[] { s }
    //            //.Concat(
    //                pt
    //                    .Where(p => !string.IsNullOrWhiteSpace(p.Text))
    //                    .Select(p => s.Replace(p.Item1, p.Text))
    //            //)
    //        )
    //        .Select(text =>
    //            !string.IsNullOrWhiteSpace(cbPräfix.Text)
    //                ? text.Replace("#titel", cbPräfix.Text)
    //                : text
    //        )
    //        .Where(text => !text.Contains('#'))
    //        .Distinct()
    //        ]
    //    );
    //}

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

        if (openFileDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(openFileDialog.FileName) && dataTable != null)
        {
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
                var row = dataTable.NewRow();
                foreach (var mapping in columnIndexMap)
                {
                    var csvIndex = mapping.Key;
                    var columnName = mapping.Value;
                    var value = splitArray[csvIndex]; // Der Index wird verwendet, um den Wert aus dem splitArray zu lesen
                    row[columnName] = string.IsNullOrEmpty(value) ? DBNull.Value : value; // Der Name wird verwendet, um die richtige Spalte in der DataRow zu finden
                }
                dataTable.Rows.Add(row);
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
                Utilities.WordInfoTaskDlg(Handle, [.. addBookDict.Keys], new(Properties.Resources.word32));
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
                Utilities.StartFile(Handle, @"Adressen.pdf");
                return true;
            case Keys.I | Keys.Control:
                Utilities.HelpMsgTaskDlg(Handle, appName, Icon);
                return true;
            case Keys.F9:
                if (tabControl.SelectedTab == addressTabPage)
                {
                    AdressenSelResetToolStripMenuItem_Click(null!, null!);
                    return true;
                }
                else { return false; }
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
                BirthdaysToolStripMenuItem_Click(null!, null!);
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
                        if (dataAdapter != null) { SaveSQLDatabase(true); }
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
                if (tabControl.SelectedTab == addressTabPage && dataTable?.GetChanges() != null) { RejectChangesToolStripMenuItem_Click(null!, null!); }
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
                    Utilities.StartDir(Handle, Path.GetDirectoryName(xmlPath) ?? string.Empty);
                    return true;
                }
            case Keys.F2 | Keys.Control | Keys.Shift:
                {
                    Utilities.StartFile(Handle, xmlPath);
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
            tbNotizen.Focus();
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
                Utilities.WordInfoTaskDlg(Handle, [.. addBookDict.Keys], new(Properties.Resources.word32));
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
                Utilities.WordInfoTaskDlg(Handle, [.. addBookDict.Keys], new(Properties.Resources.word32));
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
        Utilities.WordInfoTaskDlg(Handle, [.. addBookDict.Keys], new(Properties.Resources.word32));
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
                List<string> dateipfade = [];
                foreach (ListViewItem item in dokuListView.Items) { dateipfade.Add(item.Text); }
                if (dateipfade.Count > 0) { userdefined.Add(new UserDefined() { Key = "Dokumente", Value = string.Join("\n", dateipfade) }); }

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
                var response = await service.People.CreateContact(person).ExecuteAsync(); // Person
                if (!string.IsNullOrEmpty(response.ResourceName))
                {
                    contactNewRowIndex = -1;
                    saveTSButton.Enabled = false;
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

            //person.UserDefined ??= [];
            //var dokuexist = person.UserDefined.FirstOrDefault(ud => ud.Key == "Dokumente");
            //if (dokuexist != null)
            //{
            //    if (changedContactDocuments.Count > 0) { dokuexist.Value = string.Join("\n", changedContactDocuments); }
            //    else { person.UserDefined.Remove(dokuexist); }
            //}
            //personFields.Add("userDefined");

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
            if (personFields.Count > 0)
            {
                var updateRequest = service.People.UpdateContact(person, ressource);
                updateRequest.UpdatePersonFields = Utilities.BuildMask([.. personFields]); // Specify the fields to update
                var result = await updateRequest.ExecuteAsync();
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

    private async void DeleteGoogleContact(int rowIndex)
    {
        try
        {
            toolStripProgressBar.Style = ProgressBarStyle.Marquee;
            toolStripProgressBar.Visible = true;
            string[] scopes = [PeopleServiceService.Scope.Contacts]; // für OAuth2-Freigabe, mehrere Eingaben mit Komma gerennt (PeopleServiceService.Scope.ContactsOtherReadonly)
            UserCredential credential;
            using (FileStream stream = new(secretPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets, scopes, "user", CancellationToken.None, new FileDataStore(tokenDir, true)).Result;
            }
            var service = new PeopleServiceService(new BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = appName, });
            await DeleteContactAsync(service, rowIndex);
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "DeleteGoogleContact: " + ex.GetType().ToString(), ex.Message); }
        finally
        {
            toolStripProgressBar.Visible = false;
            toolStripProgressBar.Style = ProgressBarStyle.Blocks;
        }
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

    private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
    {
        Invoke(() =>
        {
            if (tabControl.SelectedTab == addressTabPage && dataTable != null)
            {
                if (searchTSTextBox.TextBox.TextLength > 0)
                {
                    lastAddressSearch = searchTSTextBox.TextBox.Text;
                    ignoreSearchChange = true;
                    searchTSTextBox.TextBox.Clear();
                    ignoreSearchChange = false;
                }
            }
            tabControl.SelectedIndex = 1;
            if (!Utilities.GoogleConnectionCheck(Handle, secretPath))
            {
                e.Cancel = true;
                return;
            }
            toolStripStatusLabel.Text = string.Empty;
            toolStripProgressBar.Style = ProgressBarStyle.Marquee;
            toolStripProgressBar.Visible = true;
        });
        try  // https://console.cloud.google.com/flows/enableapi?apiid=people.googleapis.com
        {
            if (!File.Exists(Path.Combine(tokenDir, "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user"))) { birthdayShow = false; }
            string[] scopes = [PeopleServiceService.Scope.Contacts]; // für OAuth2-Freigabe, mehrere Eingaben mit Komma gerennt (PeopleServiceService.Scope.ContactsOtherReadonly)
            UserCredential credential;
            using (FileStream stream = new(secretPath, FileMode.Open, FileAccess.Read)) // "..\\..\\..\\client_secret.json"
            {
                if (!Directory.Exists(tokenDir) || !Directory.EnumerateFiles(tokenDir, "*", SearchOption.AllDirectories).Any())
                {
                    Invoke(() =>
                    {
                        toolStripProgressBar.Visible = false;
                        toolStripProgressBar.Style = ProgressBarStyle.Blocks;
                    });
                }
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets, scopes, "user", CancellationToken.None, new FileDataStore(tokenDir, true)).Result;
            }
            using PeopleServiceService service = new(new BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = appName });

            //var groupResponse = service.ContactGroups.List().Execute();
            //if (groupResponse != null) { contactGroupsDict = groupResponse.ContactGroups.ToDictionary(g => g.ResourceName, g => g.Name); }

            var peopleRequest = service.People.Connections.List("people/me");
            peopleRequest.PersonFields = "names,memberships,nicknames,addresses,phoneNumbers,emailAddresses,biographies,birthdays,urls,organizations,userDefined"; // nickNames
            peopleRequest.SortOrder = (PeopleResource.ConnectionsResource.ListRequest.SortOrderEnum)3;
            peopleRequest.PageSize = 2000; // The number of connections to include in the response. Valid values are between 1 and 2000, inclusive. Defaults to 100 if not set or set to 0.
            e.Result = peopleRequest.Execute();
        }
        //catch (TokenResponseException ex)
        //{
        //    birthdayAutoShow = false;
        //    MessageBox.Show(ex.Message);
        //}
        catch (Google.GoogleApiException ex)
        {
            Invoke(() =>
            {
                toolStripProgressBar.Visible = false;
                toolStripProgressBar.Style = ProgressBarStyle.Blocks;
                Application.DoEvents();
            });
            Utilities.ErrorMsgTaskDlg(Handle, "BackgroundWorker_DoWork: " + ex.GetType().ToString(), ex.Message);
        }
    }

    private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        //var birthdayShow = true; // keine Geburtstagsanzeige wenn Token abgelaufen und Browser gestartet wird

        Invoke(() =>
        {
            toolStripProgressBar.Visible = false;
            toolStripProgressBar.Style = ProgressBarStyle.Blocks;
            toolStripStatusLabel.Visible = true;
        });
        if (e.Cancelled) { return; }
        if (e.Error != null)
        {
            if (e.Error is TokenResponseException)
            {
                Utilities.ErrorMsgTaskDlg(Handle, "Das Oauth-Zugriffstoken ist abgelaufen.", "Der Google-OAuth-Dialog wird im Browser aufgerufen,\nDort könne Sie den Zugriff auf Ihre Kontakte erneut erlauben.", TaskDialogIcon.Information);
                timer.Start();
            }
            else { Utilities.ErrorMsgTaskDlg(Handle, e.Error.GetType().ToString(), e.Error.Message); }
            return;
        }
        var response = (ListConnectionsResponse?)e.Result;
        if (response != null)
        {
            ContactEditFields(-1);
            contactDGV.Rows.Clear();
            contactDGV.Columns.Clear();
            //allDocumentsByRessourcename.Clear();
            var people = response.Connections;
            if (people != null && people.Count > 0)
            {
                contactDGV.Columns.Add("Anrede", "Anrede"); // ColumnName, HeaderText   
                contactDGV.Columns.Add("Präfix", "Präfix");
                contactDGV.Columns.Add("Nachname", "Nachname");
                contactDGV.Columns.Add("Vorname", "Vorname");
                contactDGV.Columns.Add("Zwischenname", "Zwischenname");
                contactDGV.Columns.Add("Nickname", "Nickname");
                contactDGV.Columns.Add("Suffix", "Suffix");
                contactDGV.Columns.Add("Firma", "Firma");
                contactDGV.Columns.Add("Straße", "Straße");
                contactDGV.Columns.Add("PLZ", "PLZ");
                contactDGV.Columns.Add("Ort", "Ort");
                contactDGV.Columns.Add("Land", "Land");
                contactDGV.Columns.Add("Betreff", "Betreff");
                contactDGV.Columns.Add("Grußformel", "Grußformel");
                contactDGV.Columns.Add("Schlussformel", "Schlussformel");
                contactDGV.Columns.Add("Geburtstag", "Geburtstag");
                contactDGV.Columns.Add("Mail1", "Mail1");
                contactDGV.Columns.Add("Mail2", "Mail2");
                contactDGV.Columns.Add("Telefon1", "Telefon1");
                contactDGV.Columns.Add("Telefon2", "Telefon2");
                contactDGV.Columns.Add("Mobil", "Mobil");
                contactDGV.Columns.Add("Fax", "Fax");
                contactDGV.Columns.Add("Internet", "Internet");
                contactDGV.Columns.Add("Notizen", "Notizen");
                contactDGV.Columns.Add("Ressource", "Ressource");
                for (var i = 0; i < contactDGV.Columns.Count - 1; i++) { contactDGV.Columns[i].Visible = !hideColumnArr[i]; }
                foreach (var person in people)
                {
                    var anrede = string.Empty;
                    var betreff = string.Empty;
                    var grußformel = string.Empty;
                    var schlussformel = string.Empty;
                    var dokumente = string.Empty;
                    if (person.UserDefined != null && person.UserDefined.Count > 0)
                    {
                        foreach (var customField in person.UserDefined)
                        {
                            if (customField.Key == "Anrede") { anrede = customField.Value ?? string.Empty; }
                            else if (customField.Key == "Betreff") { betreff = customField.Value ?? string.Empty; }
                            else if (customField.Key == "Grußformel") { grußformel = customField.Value ?? string.Empty; }
                            else if (customField.Key == "Schlussformel") { schlussformel = customField.Value ?? string.Empty; }
                            else if (customField.Key == "Dokumente") { dokumente = customField.Value ?? string.Empty; }
                        }
                    }

                    //if (person.Memberships != null && person.Memberships.Any())
                    //{
                    //    foreach (var membership in person.Memberships)
                    //    {
                    //        if (contactGroupsDict != null && membership.ContactGroupMembership != null)
                    //        {
                    //            var test = membership.ContactGroupMembership.ContactGroupResourceName;
                    //            if (test != null && test != "contactGroups/myContacts" && test != "contactGroups/starred")
                    //            {
                    //                var cg = membership.ContactGroupMembership;
                    //                if (cg != null && contactGroupsDict.TryGetValue(cg.ContactGroupResourceName, out var groupName))
                    //                {
                    //                    MessageBox.Show($"Kontakt ist in Gruppe: {groupName}");
                    //                }
                    //            }
                    //        }
                    //    }
                    //}
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
                        person.ResourceName ?? string.Empty
                     );
                    //if (!string.IsNullOrEmpty(dokumente) && !string.IsNullOrEmpty(person.ResourceName)) { allDocumentsByRessourcename.Add(person.ResourceName, dokumente); }
                }
                toolStripStatusLabel.Text = people.Count.ToString() + " Kontakte";
                response.Connections.Clear();  // dispose people 
                foreach (DataGridViewColumn column in contactDGV.Columns) { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
                if (!string.IsNullOrEmpty(columnWidths)) { Utilities.SetColumnWidths(columnWidths, contactDGV); }

                //contactDGV.ResumeLayout(false);
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
                //Invoke(() =>
                //{
                if (birthdayShow && birthdayAutoShow) { BirthdaysToolStripMenuItem_Click(birthdaysToolStripMenuItem, EventArgs.Empty); }
                //});
                birthdayShow = true;
            }
        }
    }

    private void Timer_Tick(object sender, EventArgs e)
    {
        if (!backgroundWorker.IsBusy)
        {
            birthdayShow = false;
            backgroundWorker.RunWorkerAsync();
            timer.Stop();
        }
    }

    private void GoogleTSButton_Click(object sender, EventArgs e)
    {
        if (!backgroundWorker.IsBusy) { backgroundWorker.RunWorkerAsync(); }
    }

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

    private void ContactEditFields(int rowIndex) // rowIndex = -1 => ClearFields
    {
        try
        {
            ignoreTextChange = true; // verhindert, dass TextChanged
            foreach (var (ctrl, colText) in dictEditField) { ctrl.Text = rowIndex < 0 ? "" : contactDGV.Rows[rowIndex].Cells[colText]?.Value?.ToString() ?? ""; }

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
        if (e.TabPage == contactTabPage && contactDGV.Rows.Count == 0 && !backgroundWorker.IsBusy) // GoogleTSButton_Click startet den BackgroundWorker
        {
            if (Utilities.YesNo_TaskDialog(Handle, "Google Kontakte", heading: "Keine Kontakte vorhanden", text: "Möchten Sie Ihre Kontakte laden?", TaskDialogIcon.ShieldBlueBar))
            {
                GoogleTSButton_Click(googleTSButton, EventArgs.Empty);
            }
            //else { e.Cancel = true; }
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
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                Text = appName + " – " + (string?)(string.IsNullOrEmpty(databaseFilePath) ? "unbenannt" : Utilities.CorrectUNC(databaseFilePath));  // Workaround for UNC-Path
                btnEditContact.Visible = false;
                saveTSButton.Enabled = dataTable?.GetChanges(DataRowState.Added | DataRowState.Modified | DataRowState.Deleted) != null;
                newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled
                    = deleteTSButton.Enabled = newToolStripMenuItem.Enabled = newTSButton.Enabled = duplicateToolStripMenuItem.Enabled = copyTSButton.Enabled = wordTSButton.Enabled
                    = envelopeTSButton.Enabled = true;
                copyToOtherDGVTSMenuItem.Enabled = false;
                var rowCount = addressDGV.Rows.Count;
                var visibleRowCount = addressDGV.Rows.Cast<DataGridViewRow>().Count(static r => r.Visible);
                toolStripStatusLabel.Text = rowCount == visibleRowCount ? $"{visibleRowCount} Adressen" : $"{visibleRowCount}/{rowCount} Adressen";
                if (addressDGV.SelectedRows.Count > 0) { AdressEditFields(addressDGV.SelectedRows[0].Index); }
            }
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
            Text = appName + " – Google-Kontakte"; // + GetMailAdress();
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
        using TaskDialogIcon questionDialogIcon = new(Properties.Resources.question32);
        TaskDialogPage page = new()
        {
            Caption = appName,
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

    private void ExtraToolStripMenuItem_DropDownOpening(object sender, EventArgs e) => authentMenuItem.Enabled = Directory.Exists(tokenDir);

    private void BrowserPeopleMenuItem_Click(object sender, EventArgs e)
    {
        try
        {
            ProcessStartInfo psi = new("https://contacts.google.com/") { UseShellExecute = true };
            Process.Start(psi);
        }
        catch (Exception ex) when (ex is Win32Exception || ex is InvalidOperationException) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
    }

    private void GoogleToolStripMenuItem_Click(object sender, EventArgs e) => GoogleTSButton_Click(sender, e);

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
        frm.SuccessDuration = sSuccessDuration;
        frm.BirthdayAutoShow = birthdayAutoShow;
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
            sSuccessDuration = frm.SuccessDuration;
            birthdayAutoShow = frm.BirthdayAutoShow;
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
    private void ComboBox_Enter(object sender, EventArgs e)
    {
        ((ComboBox)sender).BackColor = Color.LightYellow;
    }

    private void ComboBox_Leave(object sender, EventArgs e)
    {
        ((ComboBox)sender).BackColor = Color.White;
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
                addressDGV.SelectedRows[0].Cells["Geburtstag"].Value = geburtsdatum.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("de-DE"));
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
                addressDGV.SelectedRows[0].Cells["Geburtstag"].Value = string.Empty;
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
                if (dataAdapter != null) { SaveSQLDatabase(true); }
                databaseFilePath = saveFileDialog.FileName;
                if (File.Exists(databaseFilePath))
                { //Overwrite=true im SaveFileDialog bewirkt MessageBox "Datei existiert bereits. Überschreiben?"   
                    File.Delete(databaseFilePath); // File.Delete method will throw an exception in case of failure
                    Thread.Sleep(100); // Windows muss mitbekommen, dass die Datei gelöscht wurde
                }
            }
            else { return; }
        }
        catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, "NewDBToolStripMenuItem_Click: " + ex.GetType().ToString(), ex.Message); }
        try
        {   //If the database file doesn't exist, the default behaviour is to create a new database file. Use 'FailIfMissing=True' to raise an error instead.
            using (var connection = new SQLiteConnection((string?)$"Data Source={databaseFilePath};FailIfMissing=False;"))
            {
                connection.Open();
                var columnDefinitions = string.Join(", ", dataFields.Select(field => $"{field} TEXT"));
                var createTableQuery = $@"CREATE TABLE Adressen ({columnDefinitions}, Id INTEGER PRIMARY KEY AUTOINCREMENT)";
                using var command1 = new SQLiteCommand(createTableQuery, connection);
                command1.ExecuteNonQuery();

                using (var command2 = new SQLiteCommand("INSERT INTO Adressen (Anrede, Präfix, Nachname, Vorname, Zwischenname,  Nickname, Suffix, Straße, PLZ, Ort, Grußformel, Geburtstag, Mail1) " +
                    "VALUES (@Anrede, @Präfix, @Nachname, @Vorname, @Zwischenname, @Nickname, @Suffix, @Straße, @Plz, @Ort, @Grußformel, @Geburtstag, @Mail1)", connection))
                {
                    command2.Parameters.AddWithValue("@Anrede", "Herrn");
                    command2.Parameters.AddWithValue("@Präfix", "Dr. h.c.");
                    command2.Parameters.AddWithValue("@Nachname", "Mustermann");
                    command2.Parameters.AddWithValue("@Vorname", "Max");
                    command2.Parameters.AddWithValue("@Zwischenname", "von und zu");
                    command2.Parameters.AddWithValue("@Nickname", "Maxi");
                    command2.Parameters.AddWithValue("@Suffix", "Jr. MBA");  // Master of Business Administration
                    command2.Parameters.AddWithValue("@Straße", "Langer Weg 144");
                    command2.Parameters.AddWithValue("@Plz", "01234");
                    command2.Parameters.AddWithValue("@Ort", "Entenhausen");
                    command2.Parameters.AddWithValue("@Grußformel", "Lieber Max");
                    command2.Parameters.AddWithValue("@Geburtstag", "6.3.1995");
                    command2.Parameters.AddWithValue("@Mail1", "abc@xyz.com");
                    command2.ExecuteNonQuery();
                }
                connection.Close(); // not really necessary, because of using   
            }
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
            if (dataTable?.Rows.Count > 0)
            {
                try
                {
                    StringBuilder sb = new();
                    sb.AppendLine(string.Join(";", dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName)));
                    foreach (DataRow row in dataTable.Rows) { sb.AppendLine(string.Join(";", row.ItemArray.Select(field => string.Concat("\"", field?.ToString()?.Replace("\"", "\"\""), "\"")))); }
                    File.WriteAllText(saveFileDialog.FileName, sb.ToString(), Encoding.UTF8);
                }
                catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message); }
            }
        }

    }

    private void ColumnSelectToolStripMenuItem_Click(object sender, EventArgs e)
    {
        using var frm = new FrmColumns(hideColumnStd);
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
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            if (contactDGV.SelectedRows.Count <= 0) { e.Cancel = true; return; }
            else if (!Utilities.RowIsVisible(contactDGV, contactDGV.SelectedRows[0])) { contactDGV.FirstDisplayedScrollingRowIndex = contactDGV.SelectedRows[0].Index; }
            copy2OtherDGVMenuItem.Text = "Nach Lokale Adressen kopieren";
            copy2OtherDGVMenuItem.Visible = addressDGV.Rows.Count > 0;
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
                AdressEditFields(rowSelected);
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
        dataTable?.RejectChanges();
        if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0) { AdressEditFields(addressDGV.SelectedRows[0].Index); }
    }

    private void EditToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        rejectChangesToolStripMenuItem.Enabled = tabControl.SelectedTab == addressTabPage && dataTable?.GetChanges() != null;
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
            try
            {
                if (File.Exists(saveFileDialog.FileName)) { File.Delete(saveFileDialog.FileName); }
                using (var dt = new DataTable())
                {
                    foreach (DataGridViewColumn column in contactDGV.Columns) { dt.Columns.Add(column.Name); }
                    foreach (DataGridViewRow row in contactDGV.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            var dr = dt.NewRow();
                            for (var i = 0; i < contactDGV.Columns.Count; i++) { dr[i] = row.Cells[i].Value ?? string.Empty; }
                            dt.Rows.Add(dr);
                        }
                    }
                }
                var dbPath = saveFileDialog.FileName;
                SQLiteConnection.CreateFile(dbPath);
                using var connection = new SQLiteConnection((string?)$"Data Source={dbPath};Version=3;");
                connection.Open();
                var columnDefinitions = string.Join(", ", dataFields.Select(field => $"{field} TEXT"));
                var createTableSql = $@"CREATE TABLE IF NOT EXISTS Adressen ({columnDefinitions}, Id INTEGER PRIMARY KEY AUTOINCREMENT)";
                new SQLiteCommand(createTableSql, connection).ExecuteNonQuery();
                using var transaction = connection.BeginTransaction();
                var columnNames = string.Join(", ", dataFields);
                var columnParams = string.Join(", ", dataFields.Select(field => "@" + field));
                var insertCommand = new SQLiteCommand(connection) { CommandText = $@"INSERT INTO Adressen ({columnNames}) VALUES ({columnParams})" };
                foreach (DataGridViewRow row in contactDGV.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        insertCommand.Parameters.Clear();
                        for (var i = 0; i < 23; i++) { insertCommand.Parameters.AddWithValue($"@{contactDGV.Columns[i].Name}", row.Cells[i].Value ?? string.Empty); } // Felder außer Id
                        insertCommand.ExecuteNonQuery();
                    }
                }
                transaction.Commit();
                connection.Close();
                Utilities.ErrorMsgTaskDlg(Handle, "Backup erfolgreich", $"Die Google-Kontakte wurden erfolgreich in {saveFileDialog.FileName} gespeichert.", TaskDialogIcon.ShieldSuccessGreenBar);
            }
            catch (Exception ex) { Utilities.ErrorMsgTaskDlg(Handle, ex.GetType().ToString(), ex.Message, TaskDialogIcon.ShieldErrorRedBar); }
        }
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

    private void BirthdaysToolStripMenuItem_Click(object sender, EventArgs e)
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
                if (!DateTime.TryParse(row.Cells["Geburtstag"].Value.ToString(), out var geburtsdatum)) { continue; }
                var naechsterGeburtstag = new DateTime(heute.Year, geburtsdatum.Month, geburtsdatum.Day);
                if (naechsterGeburtstag < heute) { naechsterGeburtstag = naechsterGeburtstag.AddYears(1); }
                var bisGeburtstag = naechsterGeburtstag - heute; // TimeSpan
                if (bisGeburtstag.TotalDays <= birthdayRemindLimit)
                {
                    var vorname = row.Cells["Vorname"].Value.ToString(); // DataGridViewCell
                    var name = vorname + (string.IsNullOrEmpty(vorname) ? "" : " ") + row.Cells["Nachname"].Value.ToString();
                    var alter = heute.Year - geburtsdatum.Year; // if (naechsterGeburtstag > geburtsdatum.AddYears(alter)) { alter--; }
                    var id = row.Cells[idRessource].Value?.ToString() ?? string.Empty;
                    bevorstehendeGeburtstage.Add((Datum: geburtsdatum, Name: name, Alter: alter, Tage: bisGeburtstag.Days, Id: id));
                }
            }
            bevorstehendeGeburtstage = [.. bevorstehendeGeburtstage.OrderBy(g => g.Tage)];
            if (bevorstehendeGeburtstage.Count > 0)
            {
                using var frm = new FrmBirthdays(sColorScheme, bevorstehendeGeburtstage, birthdayRemindLimit, isLocal);
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    birthdayRemindLimit = frm.BirthdayRemindLimit;
                    if (frm.SelectionIndex >= 0 && frm.SelectionIndex < bevorstehendeGeburtstage.Count)
                    {
                        var selectedBirthday = bevorstehendeGeburtstage[frm.SelectionIndex];
                        foreach (DataGridViewRow row in dgv.Rows)
                        {
                            if (row.Cells[idRessource].Value?.ToString() == (string?)selectedBirthday.Id)
                            {
                                row.Selected = true;
                                dgv.FirstDisplayedScrollingRowIndex = row.Index;
                                if (tabControl.SelectedTab == addressTabPage) { AdressEditFields(row.Index); }
                                else if (tabControl.SelectedTab == contactTabPage) { ContactEditFields(row.Index); }
                                break;
                            }
                        }
                    }
                }
                else if (frm.DialogResult == DialogResult.Continue) { birthdayRemindLimit = frm.BirthdayRemindLimit; }
            }
            else
            {
                TaskDialogButton showButton = new TaskDialogCommandLinkButton("Anstehende-Geburtstage öffnen");
                var page = new TaskDialogPage()
                {
                    Caption = appName,
                    Heading = $"Keine Geburtstage innerhalb der nächsten {birthdayRemindLimit} Tage",
                    Text = "Sie können das Limit ändern (max. Anzahl der Tage vor dem Geburtstagstermin).",
                    Icon = TaskDialogIcon.ShieldWarningYellowBar,
                    AllowCancel = true,
                    SizeToContent = true,
                    Buttons = { showButton, TaskDialogButton.OK },
                };
                if (TaskDialog.ShowDialog(Handle, page) == showButton)
                {
                    using var frm = new FrmBirthdays(sColorScheme, [], birthdayRemindLimit, false);
                    if (frm.ShowDialog() == DialogResult.Continue) { birthdayRemindLimit = frm.BirthdayRemindLimit; }
                }
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

    private void AddressDGV_RowContextMenuStripNeeded(object sender, DataGridViewRowContextMenuStripNeededEventArgs e) => e.ContextMenuStrip = contextMenu;
    //private void ContactDGV_RowContextMenuStripNeeded // occurs only when the DataGridView control DataSource property is set or its VirtualMode property is true
    private void ContactDGV_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right)
        {
            var hitTestInfo = contactDGV.HitTest(e.X, e.Y);
            if (hitTestInfo.Type == DataGridViewHitTestType.Cell)
            {
                contactDGV.Rows[hitTestInfo.RowIndex].Selected = true;
                contextMenu.Show(contactDGV, new System.Drawing.Point(e.X, e.Y));
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
            var item = new ToolStripMenuItem(file) { Image = Properties.Resources.address_book16, ShortcutKeyDisplayString = first ? "F12" : string.Empty };
            first = false;
            item.Click += (s, e) =>
            {
                if (dataAdapter != null) { SaveSQLDatabase(true); }
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

    private void DokuListView_SelectedIndexChanged(object sender, EventArgs e)
    {
        dokuMinusTSButton.Enabled = dokuShowTSButton.Enabled = dokuListView.SelectedItems.Count > 0;
    }

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
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            foreach (var pfad in openFileDialog.FileNames) { Add2dokuListView(pfad, false); }
            dokuListView.ListViewItemSorter = new ListViewItemComparer();
            dokuListView.Sort();
            ListView2DataTable();
            //else if (tabControl.SelectedTab == contactTabPage) { ListView2ChangedList(); }
        }
    }

    private void ListView2DataTable()
    {
        if (dataTable == null) { return; }
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
            searchPictureBox.Image = Properties.Resources.DeleteFilter16;
            searchPictureBox.Cursor = Cursors.Hand;
            if (dokuListView.Items.Count > 0) { dokuListView.Items[0].Selected = true; }
        }
        else
        {
            searchPictureBox.Image = Properties.Resources.Search_16;
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

    private void Add2dokuListView(string pfad, bool sortAndSave = true)
    {
        ListViewItem item;
        var info = new FileInfo(pfad);
        var extension = info.Extension.ToLower();
        if (info.Exists)
        {
            if (!dokuImages.Images.ContainsKey(extension))
            {
                var icon = Icon.ExtractAssociatedIcon(pfad);
                if (icon != null) { dokuImages.Images.Add(extension, icon); }
            }
            item = new ListViewItem(info.FullName);
            item.SubItems.Add(Utilities.FormatDateigröße(info.Length));
            item.SubItems.Add(info.LastWriteTime.ToString("dd.MM.yyyy HH:mm"));
            item.ImageKey = extension;
        }
        else { item = new ListViewItem([pfad, string.Empty, string.Empty]); }
        var vorhandenesItem = dokuListView.Items.Cast<ListViewItem>().FirstOrDefault(item => string.Equals(item.Text, pfad, StringComparison.OrdinalIgnoreCase));
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
                Add2dokuListView(text);
                tabulation.SelectedTab = tabPageDoku;
                BringToFront();
            }
        }
        else if (result == nextButton)
        {
            BringToFront();
            ActiveControl = searchTextBox;
            using TaskDialogIcon questionDialogIcon = new(Properties.Resources.question32);
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
                    Add2dokuListView(text);
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
            foreach (var cell in addressDGV.Rows[prevSelectedAddressRowIndex].Cells.Cast<DataGridViewCell>().SkipLast(1).Where(cell => !Equals(originalAddressData[cell.OwningColumn.Name], cell.Value)))
            {
                changedAddressData[cell.OwningColumn.Name] = cell.Value?.ToString() ?? string.Empty;
            }
            if (changedAddressData.Count > 0) { saveTSButton.Enabled = true; }
            else if (addressDGV.Rows[prevSelectedAddressRowIndex].DataBoundItem is DataRowView dataBoundItem)
            {
                dataBoundItem.Row.AcceptChanges();  // keine Änderung, RowState auf Unchanged setzen 
                saveTSButton.Enabled = dataTable?.GetChanges() != null; // andere Rows könnten geändert sein
            }
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

    private void HelpdokuTSMenuItem_Click(object sender, EventArgs e) => Utilities.StartFile(Handle, "Adressen.pdf");

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
        static bool filter(DataGridViewRow row) => !string.IsNullOrEmpty(row.Cells["Dokumente"].Value?.ToString()); // lokale Funktion statt Lambda-Ausdruck (Func<>)
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            isFilteredAddress = true;
            flexiTSStatusLabel.Text = "… mit Briefverweis";
        }
    }

    private void MailPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => !string.IsNullOrEmpty((string?)(row.Cells["Mail1"].Value?.ToString() + row.Cells["Mail2"].Value?.ToString()));
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            isFilteredAddress = true;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            FilterContactDGV(filter);
            isFilteredContact = true;
        }
        flexiTSStatusLabel.Text = "… mit E-Mailadresse";
    }

    private void MailMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => string.IsNullOrEmpty((string?)(row.Cells["Mail1"].Value?.ToString() + row.Cells["Mail2"].Value?.ToString()));
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            isFilteredAddress = true;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            FilterContactDGV(filter);
            isFilteredContact = true;
        }
        flexiTSStatusLabel.Text = "… ohne E-Mailadresse";
    }

    private void TelephonePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => !string.IsNullOrEmpty((string?)(row.Cells["Telefon1"].Value?.ToString() + row.Cells["Telefon2"].Value?.ToString()));
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            isFilteredAddress = true;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            FilterContactDGV(filter);
            isFilteredContact = true;
        }
        flexiTSStatusLabel.Text = "… mit Telefonnummer";
    }

    private void TelephoneMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => string.IsNullOrEmpty((string?)(row.Cells["Telefon1"].Value?.ToString() + row.Cells["Telefon2"].Value?.ToString()));
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            isFilteredAddress = true;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            FilterContactDGV(filter);
            isFilteredContact = true;
        }
        flexiTSStatusLabel.Text = "… ohne Telefonnummer";
    }

    private void MobilePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => !string.IsNullOrEmpty(row.Cells["Mobil"].Value?.ToString());
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            isFilteredAddress = true;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            FilterContactDGV(filter);
            isFilteredContact = true;
        }
        flexiTSStatusLabel.Text = "… mit Mobilfunknummer";
    }

    private void MobileMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        static bool filter(DataGridViewRow row) => string.IsNullOrEmpty(row.Cells["Mobil"].Value?.ToString());
        if (tabControl.SelectedTab == addressTabPage)
        {
            FilterAddressDGV(filter);
            isFilteredAddress = true;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            FilterContactDGV(filter);
            isFilteredContact = true;
        }
        flexiTSStatusLabel.Text = "… ohne Mobilfunknummer";
    }

    private void AdressenSelResetToolStripMenuItem_Click(object sender, EventArgs e)
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
            isFilteredAddress = false;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            var rowIndex = contactDGV.SelectedRows.Count > 0 ? contactDGV.SelectedRows[0].Index : -1;
            FilterContactDGV(filter);
            if (rowIndex >= 0 && contactDGV.Rows[rowIndex] != null)
            {
                contactDGV.Rows[rowIndex].Selected = true;
                contactDGV.FirstDisplayedScrollingRowIndex = rowIndex;
            }
            isFilteredContact = false;
        }
        ignoreSearchChange = true; // F9 löst SearchTSTextBox_TextChanged aus
        searchTSTextBox.TextBox.Clear();
        ignoreSearchChange = false;
        flexiTSStatusLabel.Text = "";
    }

    private void FilterlToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        adressenMitBriefToolStripMenuItem.Enabled = adressenSelResetToolStripMenuItem.Enabled = tabControl.SelectedTab == addressTabPage && addressDGV.Rows.Count > 0;
        adressenSelResetToolStripMenuItem.Enabled = tabControl.SelectedTab == contactTabPage ? isFilteredContact : isFilteredAddress;
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
        AdressEditFields(-1); // Ihre Methode zum Zurücksetzen der Felder
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
                AdressEditFields(firstVisibleIndex); // Ihre Methode zum Füllen der Felder
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

}
