using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using Adressen.cls;
using Adressen.frm;
using Adressen.Properties;
using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using Microsoft.Win32;

namespace Adressen;

public partial class FrmAdressen : Form
{
    private readonly FrmSplashScreen? _splashScreen;
    private static readonly string appPath = Application.ExecutablePath; // EXE-Pfad
    private string _databaseFilePath = string.Empty; // Path.ChangeExtension(appPath, ".adb");
    private AppSettings _settings = new(); // Ein einziges Objekt für alle Einstellungen
    private AdressenDbContext? _context;
    private readonly string _settingsPath;
    private readonly string tokenDir;
    private readonly string secretPath;
    private readonly string cleanRegex = @"[^\+0-9]";
    private readonly string appLong = Application.ProductName ?? "Adressen & Kontakte";
    private readonly string appName = "Adressen";
    private readonly string appCont = "Kontakte";
    private readonly Dictionary<string, string> bookmarkTextDictionary = [];  // wird aus den Edit-Controls befüllt, Datenbank unabhängig
    private readonly Dictionary<Control, string> editControlsDictionary = [];
    private const int latestSchemaVersion = 1; // DB-Ziel-Version: muss bei jeder zukünftigen Änderung an der Datenbankstruktur erhöht werden!!
    private readonly string[] dataFields = ["Anrede", "Praefix", "Nachname", "Vorname", "Zwischenname", "Nickname",
        "Suffix", "Unternehmen", "Position", "Strasse", "PLZ", "Ort", "Postfach", "Land", "Betreff", "Grussformel", "Schlussformel", "Geburtstag",
        "Mail1", "Mail2", "Telefon1", "Telefon2", "Mobil", "Fax", "Internet", "Notizen"]; // Id fehlt absichtlich  
    private readonly bool argsPath = false;
    private bool isSelectionChanging = false;
    private bool ignoreTextChange = false; // ignore when changing text in ContactEditFields
    private bool ignoreSearchChange = false;
    private string lastAddressSearch = string.Empty;
    private string lastContactSearch = string.Empty;
    private ToolStripDropDown? calendarDropdown;
    private MonthCalendar? monthCalendar;
    private readonly string[] formats = ["dd.MM.yyyy", "d.MM.yyyy", "dd.M.yyyy", "d.M.yyyy", "dd.M.yy", "d.MM.yy", "d.M.yy"];
    private readonly CultureInfo culture = new("de-DE");
    private TabPage? deactivatedPage = null;
    private List<ListViewItem> allDokuLVItems = [];
    private int lastColumn = -1;
    private SortOrder lastOrder = SortOrder.None;
    private string lastTooltipText = string.Empty;
    private bool contactBirthdayFlag = true; // false wenn Zugriffstoken für Google-Kontakte fehlt oder abgelaufen ist
    private readonly string[] documentTypes = ["*.doc", "*.dot", "*.docx", "*.doct", "*.docm", "*.odt", "*.ott", "*.fodt", "*.uot", "*.pdf", "*.txt"];
    private readonly string[] imageTypes = ["*.png", "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.tif", "*.tiff", "*.webp"];
    private readonly List<string> grussformelList =
        [
        "Hallo #vorname",
        "Hallo #nickname",
        "Liebe #vorname",
        "Lieber #vorname",
        "Liebe #nickname",
        "Lieber #nickname",
        "Lieber Frau #nachname",
        "Lieber Herr #nachname",
        "Sehr geehrte Frau #nachname",
        "Sehr geehrter Herr #nachname",
        "Sehr geehrte Frau #titel #nachname",
        "Sehr geehrter Herr #titel #nachname",
        "Sehr geehrte Kollegin #nachname",
        "Sehr geehrter Kollege #nachname",
        "Sehr geehrte Kollegin #titel #nachname",
        "Sehr geehrter Kollege #titel #nachname",
        "Sehr geehrte Frau Kollegin #nachname",
        "Sehr geehrter Herr Kollege #nachname",
        "Sehr geehrte Frau Kollegin #titel #nachname",
        "Sehr geehrter Herr Kollege #titel #nachname",
        "Sehr geehrte Damen und Herren"
        ];
    private readonly string[] pictureBoxExtensions = [".bmp", ".jpg", ".jpeg", ".png", ".gif"];
    private readonly SortedSet<string> allAddressMemberships = new(StringComparer.OrdinalIgnoreCase);
    private readonly SortedSet<string> curAddressMemberships = new(StringComparer.OrdinalIgnoreCase);
    private readonly SortedSet<string> allContactMemberships = [];
    private SortedSet<string> curContactMemberships = [];
    private Contact? _lastActiveContact; // Merkt sich den Kontakt, der VOR dem Wechsel aktiv war
    private Contact? _originalContactSnapshot;
    private Dictionary<string, string> contactGroupsDict = [];
    private bool _isClosing = false; // Flag, um Endlosschleife zu verhindern
    private bool _isFiltering = false; // Verhindert Speichern während der Suche
    private BindingList<Contact> _allGoogleContacts = []; // Klassenvariable
    private bool _isDarkMode;
    private CancellationTokenSource? _googleCts; // Wenn der User die Kontakte lädt und kurz darauf erneut klickt, soll der erste Ladevorgang abgebrochen werden: deshalb global!!
    private int _currentDbVersion;
    private bool _isTabSwitchingProgrammatically = false; // Verhindert unerwünschte Event-Auslösung bei Tab-Wechseln durch Code
    private TabPage? _previousTab;  // innerhalb des Selecting-Events kann man sich nicht auf tabControl.SelectedTab verlassen
    private string _lastSearchTerm = string.Empty;
    private bool _lastMatchCase = false;
    private bool _isCheckingContactChanges = false;
    private object? _lastProcessedEntry;
    private int _savedAddressScrollIndex = -1;
    private int _savedContactScrollIndex = -1;
    private Font? _monospaceFont;
    private Font? _standardAppFont;
    private CancellationTokenSource? _scrollCancellation;  // merkt sich den aktuellen Scroll-Vorgang
    private bool _isPanelFrozen = false;
    private bool _isDialogVisible = false;

    [GeneratedRegex("#(vorname|nickname|nachname|titel)", RegexOptions.IgnoreCase)]
    private static partial Regex PlaceholderRegex();

    public FrmAdressen(FrmSplashScreen? splashScreen, string[] args)
    {
        //// 1. Argumente parsen (wie im Original)
        //_migrationRequested = args.Any(a =>
        //        a.Equals("/migrate", StringComparison.OrdinalIgnoreCase) ||
        //        a.Equals("-migrate", StringComparison.OrdinalIgnoreCase));

        if (args.Length >= 1)
        {
            if (File.Exists((string?)args[0]))
            {
                _databaseFilePath = (string?)args[0] ?? string.Empty;
                if (!string.IsNullOrEmpty(_databaseFilePath)) { argsPath = true; }
            }
        }

        InitializeComponent();

        // 2. Basis-Initialisierung
        _splashScreen = splashScreen;
        // DoubleBuffered via Reflection für flüssigeres Rendering
        typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, addressDGV, [true]);
        typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, contactDGV, [true]);
        typeof(TableLayoutPanel).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, tableLayoutPanel, [true]);
        typeof(FlowLayoutPanel).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, flowLayoutPanel, [true]);

        _isDarkMode = DefaultBackColor.R < 128;
        UpdateAppearanceStatus();
        _previousTab = tabControl.SelectedTab;
        // ImageList Setup
        imageList.Images.Add(Resources.address_book);
        imageList.Images.Add(Resources.address_book_blue);
        imageList.Images.Add(Resources.universal24);
        imageList.Images.Add(Resources.inbox24);
        imageList.Images.Add(Resources.inboxdoc24);
        tabControl.ImageList = imageList;
        tabControl.TabPages[0].ImageIndex = 0;
        tabControl.TabPages[1].ImageIndex = 1;
        tabulation.TabPages[0].ImageIndex = 2;
        tabulation.TabPages[1].ImageIndex = 3;

        // 3. Pfade ermitteln (InnoSetup vs. Portable)
        if (Utils.IsInnoSetupValid(appPath))
        {
            _settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName, appName + ".json");
            tokenDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName, "token.json");
            secretPath = Path.Combine(Path.GetDirectoryName(appPath) ?? string.Empty, "client_secret.json");
        }
        else
        {
            _settingsPath = Path.ChangeExtension(appPath, ".json");
            tokenDir = Path.Combine(AppContext.BaseDirectory, "token.json");
            secretPath = Path.Combine(AppContext.BaseDirectory, "client_secret.json");
        }

        // 4. Einstellungen SYNCHRON laden
        if (File.Exists(_settingsPath)) { _settings = SettingsManager.Load(_settingsPath); }
        else
        {
            var dir = Path.GetDirectoryName(_settingsPath);
            if (dir != null) { Directory.CreateDirectory(dir); }
        }

        // 5. Weitere UI-Vorbereitungen
        addressDGV.ColumnHeadersDefaultCellStyle.SelectionBackColor = addressDGV.ColumnHeadersDefaultCellStyle.BackColor;
        contactDGV.ColumnHeadersDefaultCellStyle.SelectionBackColor = contactDGV.ColumnHeadersDefaultCellStyle.BackColor;
        searchTSTextBox.TextBox.PlaceholderText = " Suche";

        FillDictionary();
        editControlsDictionary.Add(cbAnrede, "Anrede");
        editControlsDictionary.Add(cbPraefix, "Praefix");
        editControlsDictionary.Add(tbNachname, "Nachname");
        editControlsDictionary.Add(tbVorname, "Vorname");
        editControlsDictionary.Add(tbZwischenname, "Zwischenname");
        editControlsDictionary.Add(tbNickname, "Nickname");
        editControlsDictionary.Add(tbSuffix, "Suffix");
        editControlsDictionary.Add(tbFirma, "Unternehmen");
        editControlsDictionary.Add(tbPosition, "Position");
        editControlsDictionary.Add(tbStraße, "Strasse");
        editControlsDictionary.Add(cbPLZ, "PLZ");
        editControlsDictionary.Add(cbOrt, "Ort");
        editControlsDictionary.Add(tbPostfach, "Postfach");
        editControlsDictionary.Add(cbLand, "Land");
        editControlsDictionary.Add(tbBetreff, "Betreff");
        editControlsDictionary.Add(cbGrussformel, "Grussformel");
        editControlsDictionary.Add(cbSchlussformel, "Schlussformel");
        //editControlsDictionary.Add(maskedTextBox, "Geburtstag");
        editControlsDictionary.Add(tbMail1, "Mail1");
        editControlsDictionary.Add(tbMail2, "Mail2");
        editControlsDictionary.Add(tbTelefon1, "Telefon1");
        editControlsDictionary.Add(tbTelefon2, "Telefon2");
        editControlsDictionary.Add(tbMobil, "Mobil");
        editControlsDictionary.Add(tbFax, "Fax");
        editControlsDictionary.Add(tbInternet, "Internet");
        editControlsDictionary.Add(tbNotizen, "Notizen");

        //// Event Handler
        foreach (ToolStripItem item in menuStrip.Items)
        {
            if (item is ToolStripDropDownItem dropDownItem) { dropDownItem.DropDown.Opening += new CancelEventHandler(MainDropDown_Opening); }
        }
        // 6. UI basierend auf geladenen Settings anwenden
        //ApplySettingsToUI();
        Utils.RestoreWindowBounds(this, _settings.MainWindowPosition, _settings.WindowMaximized);
        _settings.SplitterPosition = _settings.SplitterPosition > 0 ? _settings.SplitterPosition : splitContainer.SplitterDistance;
        NativeMethods.SendMessage(searchTSTextBox.TextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (AppSettings.TextBoxPadding << 16) | (AppSettings.TextBoxPadding & 0xFFFF));
        SetColorScheme();
        tsClearLabel.Visible = false;

        min2TrayTSButton.BackColor = tableLayoutPanel.BackColor;  // SetColorScheme muss vorher aufgerufen worden sein.
        min2TrayTSButton.MouseEnter += (s, e) => { min2TrayTSButton.BackColor = toolStrip.BackColor; };  // Hover-Effekt (Maus betritt den Button)
        min2TrayTSButton.MouseLeave += (s, e) => { min2TrayTSButton.BackColor = tableLayoutPanel.BackColor; };  // Normalzustand (Maus verlässt den Button)
        min2TrayTSButton.MouseDown += (s, e) => { min2TrayTSButton.BackColor = Color.DarkGray; };  // Klick-Effekt (Maustaste wird gedrückt)
        min2TrayTSButton.MouseUp += (s, e) => { min2TrayTSButton.BackColor = toolStrip.BackColor; };  //  // Geht zurück in den Hover-Zustand
        min2TrayTSButton.Paint += Min2TrayTSButton_Paint;
    }

    private void Min2TrayTSButton_Paint(object? sender, PaintEventArgs e)
    {
        if (sender is not ToolStripButton btn || btn.Owner == null) { return; }
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        var fullRect = new Rectangle(0, 0, btn.Width, btn.Height);  // Das gesamte Rechteck des Buttons
        var borderRect = new Rectangle(0, 0, btn.Width - 1, btn.Height - 1);  // Das etwas kleinere Rechteck für den Rahmen
        var radius = 4;
        using var roundedPath = Utils.GetRoundedRectanglePath(borderRect, radius);
        using (var cornerRegion = new Region(fullRect))
        {
            cornerRegion.Exclude(roundedPath);  // die Ecken radieren, indem wir die runde Form ausschneiden
            using var eraserBrush = new SolidBrush(btn.Owner.BackColor);
            e.Graphics.FillRegion(eraserBrush, cornerRegion);  // mit Hintergrundfarbe des ToolStrips übermalen
        }
        using var pen = new Pen(Color.DarkGray, 1f);
        e.Graphics.DrawPath(pen, roundedPath);  // den runden Rahmen zeichnen   
    }

    private void FrmAdressen_Load(object sender, EventArgs e)
    {
        addressTabPage.ApplyMaxLengthFromEntity<Adresse>();
        contactTabPage.ApplyMaxLengthFromEntity<Contact>();
        ApplyEditControlsFont();
        ApplyFileWatcherSettings();

    }

    private async void FrmAdressen_Shown(object sender, EventArgs e)
    {
        Update();  // sicherer als DoEvents(), da es nur Painting betrifft; soll weiße Flächen verhindern
        UseWaitCursor = true; // NEU: Hält den Lade-Mauszeiger auch bei Mausbewegungen während 'await', statt Cursor.Current = Cursors.WaitCursor;
        searchTSTextBox.TextBox.Focus();
        try
        {
            splitContainer.SplitterDistance = _settings.SplitterPosition;
            UpdateStatusLabelWidth();  // flexiTSStatusLabel.Width = 244 + splitContainer.SplitterDistance - 536;
            await Task.Delay(50);  // Lässt den UI-Thread kurz durchatmen und die Form komplett rendern
            if (Utils.IsUpdateCheckDue(_settings.UpdateIndex, _settings.LastUpdateCheck))
            {
                var (version, date) = await Utils.GetLatestVersionInfoAsync();
                RefreshUpdateUI(version, date);
            }
            if (!argsPath && _settings.ReloadRecent) { _databaseFilePath = _settings.RecentFiles.Count > 0 ? _settings.RecentFiles[0] : string.Empty; }
            if ((_settings.ReloadRecent || argsPath) && !string.IsNullOrEmpty(_databaseFilePath)) { await ConnectSQLDatabaseAsync(_databaseFilePath); }
            else if (!_settings.ReloadRecent && !_settings.NoAutoload && !string.IsNullOrEmpty(_settings.StandardFile)) { await ConnectSQLDatabaseAsync(_settings.StandardFile); }
            if (_settings.ContactsAutoload) { await LoadAndDisplayGoogleContactsAsync(); }

        }
        finally
        {
            UseWaitCursor = false; // NEU: Cursor wieder freigeben, statt Cursor.Current = Cursors.Default;
            if (_splashScreen != null)
            {
                _splashScreen.Close();
                _splashScreen.Dispose();
            }
        }
    }

    private void UpdateStatusLabelWidth()
    {
        var startX = toolStripStatusLabel.Bounds.Right + flexiTSStatusLabel.Margin.Left;
        var progressBarWidth = toolStripProgressBar.Width + toolStripProgressBar.Margin.Horizontal;
        var targetWidth = splitContainer.SplitterDistance - startX - progressBarWidth - 1;
        flexiTSStatusLabel.Width = Math.Max(0, targetWidth);
    }

    private void UpdateSearchBoxWidth()
    {
        var startX = toolStripSeparator2.Bounds.Right + searchTSTextBox.Margin.Left;
        var clearLabelWidth = tsClearLabel.Visible ? (tsClearLabel.Width + tsClearLabel.Margin.Horizontal) : 0;
        var availableWidth = splitContainer.SplitterDistance - startX - clearLabelWidth - 1;
        searchTSTextBox.Width = Math.Max(50, availableWidth);
    }

    private void SaveConfiguration()
    {
        if (WindowState != FormWindowState.Minimized) { _settings.WindowMaximized = WindowState == FormWindowState.Maximized; }  // ignoriere "Minimized", da wir sonst den Maximiert-Status vergessen würden
        var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        _settings.MainWindowPosition = new WindowPlacement { X = bounds.X, Y = bounds.Y, Width = bounds.Width, Height = bounds.Height };
        _settings.SplitterPosition = splitContainer.SplitterDistance;
        var activeDGV = tabControl.SelectedTab == contactTabPage ? contactDGV : addressDGV;
        if (activeDGV != null && activeDGV.Columns.Count > 0)
        {
            _settings.HideColumnArr = [.. activeDGV.Columns.Cast<DataGridViewColumn>().Select(c => !c.Visible)];
            _settings.ColumnWidths = [.. activeDGV.Columns.Cast<DataGridViewColumn>().Select(c => c.Width)];
        }
        SettingsManager.Save(_settings, _settingsPath);
    }

    private async Task ConnectSQLDatabaseAsync(string file)
    {
        // 1. Checks (unverändert)
        if (string.IsNullOrEmpty(file) || !File.Exists(file))
        {
            Utils.MsgTaskDlg(Handle, "Datenbank-Datei nicht gefunden", file, TaskDialogIcon.ShieldWarningYellowBar);
            _settings.RecentFiles.Remove(file);
            return;
        }

        // 2. UI-Feedback: Feste Schritte statt Lauflicht
        toolStripProgressBar.Visible = true;
        toolStripProgressBar.Style = ProgressBarStyle.Continuous; // Oder 'Blocks'
        toolStripProgressBar.Minimum = 0;  // zur Sicherheit, falls es vorher im Lauflichtmodus war
        toolStripProgressBar.Maximum = 100; // 100% als Maximalwert 
        toolStripProgressBar.Value = 15; // Startwert

        toolStripStatusLabel.Text = "Öffne Datenbank...";
        statusStrip.Update();

        try
        {
            CloseDatabaseConnection();
            _databaseFilePath = Utils.CorrectUNC(file);  // hier einmalig CorrectUNC aufrufen, damit wir konsistenten Pfad haben

            _currentDbVersion = DatabaseMigrator.GetDatabaseVersion(_databaseFilePath);
            //MessageBox.Show($"Datenbankversion: {_currentDbVersion}\nErwartete Version: {AppSettings.DatabaseSchemaVersion}", "Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (_currentDbVersion > AppSettings.DatabaseSchemaVersion)  // Downgrade-Schutz
            {
                Utils.MsgTaskDlg(Handle, "Datenbank zu neu",
                    "Diese Datenbank wurde mit einer neueren Version der Software erstellt. " +
                    "Bitte aktualisiere das Programm.", TaskDialogIcon.ShieldErrorRedBar);
                return;
            }


            _context = new AdressenDbContext(_databaseFilePath);

            // OPTIMIERUNG 1: WAL Modus aktivieren (Massiver Performance-Gewinn)
            await _context.Database.OpenConnectionAsync();

            // --- NEU: Eigene Sortierung (Collation) registrieren ---
            // Wir holen uns die rohe SQLite-Verbindung
            if (_context.Database.GetDbConnection() is SqliteConnection sqliteConnection)
            {
                // Wir definieren "GERMAN" als Sortierregel, die C# CultureInfo nutzt
                sqliteConnection.CreateCollation("GERMAN", (x, y) => string.Compare(x, y, new CultureInfo("de-DE"), CompareOptions.IgnoreCase));
            }
            // -------------------------------------------------------

            await _context.Database.ExecuteSqlRawAsync("PRAGMA journal_mode = WAL;");
            // Optional: Synchronous Commit auf NORMAL setzen (schneller, immer noch sicher genug für Desktop)
            await _context.Database.ExecuteSqlRawAsync("PRAGMA synchronous = NORMAL;");

            toolStripProgressBar.Value = 30; // Fortschritt: 30%

            // SCHRITT B: Migration
            var migrationDone = false;
            if (_currentDbVersion < AppSettings.DatabaseSchemaVersion)
            {
                toolStripStatusLabel.Text = "Führe Migration durch...";
                statusStrip.Update();

                // Wir rufen die Migration OHNE Handle auf
                migrationDone = await Task.Run(() => DatabaseMigrator.MigrateLegacyData(_context));

                if (migrationDone)
                {
                    _currentDbVersion = AppSettings.DatabaseSchemaVersion;

                    // Erfolgsdialog sicher im UI-Thread anzeigen
                    Utils.MsgTaskDlg(Handle, "Datenbank aktualisiert",
                        $"Die Datenbank wurde erfolgreich migriert (v{AppSettings.DatabaseSchemaVersion}).",
                        TaskDialogIcon.ShieldSuccessGreenBar);
                }
            }


            // SCHRITT C: Laden (Der längste Teil)
            // Wir setzen ihn auf 50%, wohl wissend, dass er hier kurz "hängt"
            toolStripProgressBar.Value = 50;
            toolStripStatusLabel.Text = "Lade Datensätze...";
            statusStrip.Update();

            // OPTIMIERUNG 2: ChangeTracker temporär pausieren
            // Das verhindert, dass EF Core beim massenhaften Aufbau der Entities
            // ständig intern prüft, ob sich Eigenschaften geändert haben.
            _context.ChangeTracker.AutoDetectChangesEnabled = false;

            try
            {
                // 1. Alle Gruppen vorab in den Cache laden
                await _context.Gruppen.LoadAsync();

                // 2. Adressen laden (mit Split Query Optimierung)
                await _context.Adressen
                    .Include(a => a.Gruppen)
                    .Include(a => a.Dokumente) // Dokumente direkt mitladen, damit Filter funktioniert

                    // OPTIMIERUNG 3: AsSplitQuery verhindert riesige JOIN-Datenmengen
                    .AsSplitQuery()

                    .OrderBy(a => EF.Functions.Collate(a.Nachname, "GERMAN"))
                    .ThenBy(a => EF.Functions.Collate(a.Vorname, "GERMAN"))
                    .LoadAsync();
            }
            finally
            {
                // WICHTIG: Danach unbedingt wieder einschalten, damit spätere
                // Eingaben des Benutzers im DataGridView auch als Änderung erkannt werden!
                _context.ChangeTracker.AutoDetectChangesEnabled = true;
            }


            toolStripProgressBar.Value = 80;
            toolStripStatusLabel.Text = "Erstelle Ansicht...";
            statusStrip.Update();

            addressBSource.DataSource = _context.Adressen.Local.ToBindingList();
            addressDGV.DataSource = addressBSource;
            AutoValidate = AutoValidate.EnableAllowFocusChange; // Fehler im Validating-Event anzeigen, aber Fokuswechsel erlauben; Standard = EnablePreventFocusChange
            ApplyColumnSettings(addressDGV);
            foreach (DataGridViewColumn column in addressDGV.Columns) { column.SortMode = DataGridViewColumnSortMode.NotSortable; }

            PopulateMemberships();
            SwitchDataBinding(addressBSource);

            if (_context != null)
            {
                _settings.RecentFiles.Remove(_databaseFilePath);
                _settings.RecentFiles.Insert(0, _databaseFilePath);
                if (_settings.RecentFiles.Count > AppSettings.MaxRecentFiles) { _settings.RecentFiles = [.. _settings.RecentFiles.Take(AppSettings.MaxRecentFiles)]; }

                //newToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = deleteTSButton.Enabled = newTSButton.Enabled = duplicateToolStripMenuItem.Enabled = 
                //copyTSButton.Enabled = clipboardTSButton.Enabled = wordTSButton.Enabled = envelopeTSButton.Enabled = true;

                SetCommonButtonState(true);
                copyToOtherDGVTSMenuItem.Enabled = false;

                tabControl.SelectTab(0);

                _context.ChangeTracker.StateChanged += OnStateChanged;
                addressBSource.CurrentChanged += AddressBindingSource_CurrentChanged;

                if (addressBSource.Count > 0) { AddressBindingSource_CurrentChanged(this, EventArgs.Empty); }

                if (!migrationDone && _settings.BirthdayAddressShow)
                {
                    _ = InvokeAsync(() => BirthdayReminder(addressDGV));  // erst ausführen, wenn die UI aktualisiert ist, damit der Dialog über dem Hauptfenster erscheint
                }

                _ = Task.Run(() => Utils.StartSearchCacheWarmup(_context.Adressen.Local));

                // SCHRITT E: Fertig
                _lastProcessedEntry = null;
                tableLayoutPanel.SuspendLayout();  // Layout-Berechnung pausieren
                //maskedTextBox.Enabled = false;  // Control für die Dauer des Daten-Pumpens "taub" schalten. Verhindert aufblitzende Cursor und Placeholder-Sprünge
                AddressBindingSource_CurrentChanged(addressBSource, EventArgs.Empty);  // Einmalig feuern für den ersten Datensatz
                tableLayoutPanel.ResumeLayout(true);  // Layout-Berechnung wieder aktivieren und neu zeichnen
                //maskedTextBox.Enabled = true;  // Control wieder aufwecken (zeichnet sich jetzt genau EINMAL in der Endposition)
                toolStripProgressBar.Value = 100; // Voller Balken
                toolStripStatusLabel.Text = $"{addressBSource.Count} Adressen";
                statusStrip.Update();
            }
        }
        catch (Exception ex)
        {
            toolStripStatusLabel.Text = "Fehler beim Laden.";
            Utils.ErrTaskDlg(Handle, ex);
        }
        finally { toolStripProgressBar.Visible = false; }
    }

    private void OnStateChanged(object? sender, EntityStateChangedEventArgs e) => UpdateSaveButton();

    private void PopulateMemberships()
    {
        if (addressBSource is null || _context is null) { return; }
        allAddressMemberships.Clear();
        allAddressMemberships.Add("★"); // Favoriten immer zuerst
        var dbGruppen = _context.Gruppen.Select(g => g.Name).Distinct().ToList();
        allAddressMemberships.UnionWith(dbGruppen);
        UpdateTagComboBoxDataSource();
    }

    private void CreateNewDatabase(string filePath, bool addSampleRecord = false)
    {
        try
        {
            SqliteConnection.ClearAllPools(); // bestehende Pools leeren, um Dateisperren zu vermeiden
            if (File.Exists(filePath)) { File.Delete(filePath); }
            using var dbContext = new AdressenDbContext(filePath);
            dbContext.Database.EnsureCreated(); // Erstellt die Datenbank und ALLE Tabellen (Adressen, Gruppen, Dokumente, Foto)
            if (addSampleRecord)
            {
                var sampleAdresse = new Adresse
                {
                    Anrede = "Herrn",
                    Praefix = "Dr.",
                    Nachname = "Mustermann",
                    Vorname = "Max",
                    Zwischenname = "Moritz",
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

    private async Task<DialogResult> SaveSQLDatabaseAsync(bool closeDB = false, bool askNever = false, bool isFormClosing = false)
    {
        var isInputValid = false;
        addressDGV.CausesValidation = false;
        contactDGV.CausesValidation = false;
        try
        {
            ActiveControl = null;
            isInputValid = ValidateChildren(ValidationConstraints.Enabled);
            addressBSource?.EndEdit();
        }
        finally
        {
            addressDGV.CausesValidation = true;
            contactDGV.CausesValidation = true;
        }
        var analysis = DbChangeAnalyzer.AnalyzeChanges(_context);
        if (_context == null || !analysis.HasChanges)
        {
            if (closeDB) { CloseDatabaseConnection(); } // Ok, hier dürfen wir schließen
            return DialogResult.None;
        }
        if (!askNever && _settings.AskBeforeSaveSQL)
        {
            if (tabControl.SelectedTab != addressTabPage) { tabControl.SelectTab(addressTabPage); }
            var saveButton = new TaskDialogButton("&Speichern");
            var dontSaveButton = new TaskDialogButton("&Nicht speichern");
            var cancelButton = TaskDialogButton.Cancel;
            using var questionDialogIcon = new TaskDialogIcon(Resources.question32); // Ggf. Properties.Resources.question32
            var page = new TaskDialogPage()
            {
                Caption = appName,
                Heading = analysis.DialogHeading,
                Text = analysis.DialogText,
                Icon = questionDialogIcon,
                AllowCancel = true,
                SizeToContent = true,
                Verification = new TaskDialogVerificationCheckBox() { Text = "Immer fragen" },
                Buttons = { saveButton, dontSaveButton, cancelButton }
            };
            if (_settings.AskBeforeSaveSQLExpander && !string.IsNullOrEmpty(analysis.ExpanderText))
            {
                page.Expander = new TaskDialogExpander() { Text = analysis.ExpanderText, Expanded = false };
            }

            if (page.Verification is TaskDialogVerificationCheckBox check) { check.Checked = _settings.AskBeforeSaveSQL; }
            var result = TaskDialog.ShowDialog(this, page);
            if (result != cancelButton)  // CheckBox-Status nur übernehmen, wenn der Vorgang nicht komplett abgebrochen wurde
            {
                if (page.Verification is TaskDialogVerificationCheckBox finalCheck)
                {
                    if (_settings.AskBeforeSaveSQL && !finalCheck.Checked)
                    {
                        Utils.MsgTaskDlg(Handle, "Hinweis", "Du kannst die Sicherheitsabfrage in\nden Einstellungen wieder einschalten.", new(Resources.info32));
                        _settings.AskBeforeSaveSQL = false;
                    }
                    else if (finalCheck.Checked) { _settings.AskBeforeSaveSQL = true; }
                }
            }
            if (result == cancelButton) { return DialogResult.Cancel; }  // WICHTIG: User hat abgebrochen! Nichts tun, DB bleibt offen.
            if (result == dontSaveButton)
            {
                _isFiltering = true;
                try
                {
                    await DbChangeAnalyzer.RevertChangesAsync(analysis.RealChanges);
                    var changedEntries = _context.ChangeTracker.Entries().Where(e => e.State != EntityState.Unchanged).ToList();
                    foreach (var entry in changedEntries) { entry.State = EntityState.Unchanged; }
                }
                finally { _isFiltering = false; }

                if (closeDB) { CloseDatabaseConnection(); } // Ok, User will Änderungen verwerfen und beenden
                return DialogResult.No;
            }
        }
        if (!isInputValid)
        {
            Utils.MsgTaskDlg(Handle, "Speichern nicht möglich", "Einige Eingaben sind ungültig oder unvollständig.", TaskDialogIcon.ShieldErrorRedBar);
            return DialogResult.Cancel; // WICHTIG: DB bleibt offen, damit der User den Fehler korrigieren kann!
        }
        try
        {
            await _context.SaveChangesAsync();
            await _context.Database.ExecuteSqlRawAsync("PRAGMA wal_checkpoint(TRUNCATE);");
            if (!isFormClosing)
            {
                Invoke(() =>
                {
                    saveTSButton.Enabled = false;
                    flexiTSStatusLabel.Text = $"Letztes Speichern: {DateTime.Now:HH:mm:ss}";
                });
            }
            if (_settings.DailyBackup && File.Exists(_databaseFilePath) && Directory.Exists(_settings.BackupDirectory))
            {
                if (isFormClosing) { await Utils.DailyBackupAsync(_databaseFilePath, _settings.BackupDirectory); }
                else { _ = Utils.DailyBackupAsync(_databaseFilePath, _settings.BackupDirectory); }
            }
            if (_settings.AddZipBackup && File.Exists(_databaseFilePath) && !string.IsNullOrWhiteSpace(_settings.AddZipDirectory))
            {
                if (isFormClosing) { await Utils.UpdateZipBackupAsync(_databaseFilePath, _settings.AddZipDirectory); }
                else { _ = Utils.UpdateZipBackupAsync(_databaseFilePath, _settings.AddZipDirectory); }
            }
            if (closeDB) { CloseDatabaseConnection(); } // Ok, erfolgreich gespeichert, wir können schließen
            return DialogResult.Yes;
        }
        catch (DbUpdateConcurrencyException dbEx)
        {
            Utils.MsgTaskDlg(Handle, "Konflikt beim Speichern", $"Details: {dbEx.Message}\nIhre lokalen Änderungen werden verworfen.");
            foreach (var entry in dbEx.Entries) { await entry.ReloadAsync(); }
            Invoke(() => { saveTSButton.Enabled = false; }); // Sicherheitshalber auch hier Invoke nutzen!
            return DialogResult.Abort; // DB bleibt offen
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            return DialogResult.Abort; // DB bleibt offen
        }
    }

    private void CloseDatabaseConnection()
    {
        _lastProcessedEntry = null; // Ganz wichtig!
        // 1. Events abklemmen, damit keine Logik mehr getriggert wird
        addressBSource.CurrentChanged -= AddressBindingSource_CurrentChanged;
        _context?.ChangeTracker.StateChanged -= OnStateChanged;

        // 2. REIHENFOLGE GEÄNDERT: Erst das Grid vom Binding lösen!
        // Wenn das DGV zuerst auf null gesetzt wird, sucht es nicht mehr nach "Nachname",
        // wenn die BindingSource danach geleert wird.
        addressDGV?.DataSource = null;
        contactDGV?.DataSource = null;

        // 3. UI-Controls säubern
        AutoValidate = AutoValidate.Disable;
        maskedTextBox?.DataBindings.Clear();
        maskedTextBox?.Text = string.Empty;
        topAlignZoomPictureBox.Image = Resources.AddressBild100;
        flowLayoutPanel.Controls.Clear();
        dokuListView.Items.Clear();
        tabPageDoku.ImageIndex = 3;

        // 4. BindingSources "neutralisieren"
        // Wir setzen sie auf den Typ zurück, damit Metadaten erhalten bleiben, 
        // aber keine Instanzen mehr da sind. Das verhindert Bindungsfehler.
        addressBSource.DataSource = typeof(Adresse);
        contactBSource.DataSource = typeof(Contact);

        // 5. Context entsorgen
        _context?.Dispose();
        _context = null;

        Debug.WriteLine("Datenbankverbindung sicher getrennt.");
    }

    private async void OpenToolStripMenuItem_Click(object? sender, EventArgs? e)
    {
        await CheckContactChanges(async () =>
        {
            if (_context != null) { await SaveSQLDatabaseAsync(true); }  // CloseDatabaseConnection wird durch 'true' bereits aufgerufen.

            openFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";

            var fullPath = _databaseFilePath;
            var fileName = Path.GetFileName(fullPath) ?? "Adressen.adb";
            var dirName = Path.GetDirectoryName(fullPath);

            openFileDialog.FileName = fileName;
            openFileDialog.InitialDirectory = (_settings.DatabaseFolder is { Length: > 0 } dbDir && Directory.Exists(dbDir)) ? dbDir : dirName ?? string.Empty;
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                await ConnectSQLDatabaseAsync(openFileDialog.FileName);
                SetSearchTextIgnoreChange(string.Empty);  // Textfeld sicher und ohne Event-Sturm leeren
                ApplyGlobalSearch(string.Empty, jumpToFirstRow: true); // Hier true, da wir bei einer neuen DB oben starten wollen
            }
        });
    }

    private async void ExitToolStripMenuItem_Click(object? sender, EventArgs? e)
    {
        if (addressBSource != null) { await SaveSQLDatabaseAsync(true); }
        Close();
    }

    private async void AddressDGV_CellClick(object sender, DataGridViewCellEventArgs e)
    {
        // 1. Validitätsprüfung (Header-Klicks ausschließen)
        if (e.RowIndex < 0 || e.ColumnIndex < 0) { return; }

        // 2. Prüfung auf Strg-Taste (WinForms-Standard)
        if ((ModifierKeys & Keys.Control) == Keys.Control)
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
                if (targetControl is TextBoxBase tb) { tb.SelectAll(); }
                // Für ComboBoxen die Dropdown-Liste öffnen (optional)
                else if (targetControl is ComboBox cb) { cb.DroppedDown = true; }
            }
        }
    }

    private async void AddressBindingSource_ListChanged(object? sender, ListChangedEventArgs e) => UpdateSaveButton();  // ListChanged sollte niemals schwere Logik (DB-Zugriffe) enthalten!

    private async void AddressBindingSource_CurrentChanged(object? sender, EventArgs e)
    {
        if (_isFiltering) { return; }
        ErzeugeGrussformeln(addressBSource.Current);  // Cave, muss hier bleiben und nicht nach Task.Delay => DataBinding und AutoComplete im selben UI-Frame
        if (_scrollCancellation != null)
        {
            _scrollCancellation.Cancel();
            _scrollCancellation.Dispose();
        }
        _scrollCancellation = new CancellationTokenSource();
        var token = _scrollCancellation.Token;
        //try
        //{
        //    await Task.Delay(100, token);
        //    if (!token.IsCancellationRequested) { UpdateAddressDetails(token); }
        //}
        try
        {
            await Task.Delay(150, token);

            // DER NEUE GATEKEEPER: Wenn die Taste immer noch gedrückt ist, 
            // brechen wir ab. Keine Bilder laden, kein EF Core!
            if (Utils.IsArrowKeyDown() || token.IsCancellationRequested)
            {
                return;
            }

            UpdateAddressDetails(token);
        }
        catch (TaskCanceledException) { }
    }



    private void UpdateAddressDetails(CancellationToken token)
    {
        try
        {
            ignoreTextChange = true;
            if (addressBSource?.Current is Adresse currentAdresse)
            {
                _ = ShowPhotoInPictureBox(currentAdresse, token);
                _ = LoadDetailsAsync(currentAdresse, token);

                if (currentAdresse.Geburtstag.HasValue) { AgeLabel_MaskedTB_Set(currentAdresse.Geburtstag.Value); }
                else { AgeLabel_MaskedTB_Clear(); }
            }
            else
            {
                topAlignZoomPictureBox.Image = Resources.AddressBild100;
                delPictboxToolStripButton.Enabled = false;
                curAddressMemberships.Clear();
                UpdateMembershipTags();
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

    private async Task LoadDetailsAsync(Adresse adresse, CancellationToken token)
    {
        if (_context == null) { return; }
        try
        {
            var entry = _context.Entry(adresse);

            // Explizites Nachladen der Gruppen, falls noch nicht geschehen (mit Abbruch-Token)
            if (!entry.Collection(a => a.Gruppen).IsLoaded) { await entry.Collection(a => a.Gruppen).LoadAsync(token); }

            // Explizites Nachladen der Dokumente (mit Abbruch-Token)
            if (!entry.Collection(a => a.Dokumente).IsLoaded) { await entry.Collection(a => a.Dokumente).LoadAsync(token); }

            // Nur aktualisieren, wenn der Nutzer in der Zwischenzeit nicht gescrollt hat
            if (addressBSource.Current == adresse && !token.IsCancellationRequested)
            {
                LoadGroupsForCurrentAddress();
                UpdateDocumentListView(adresse);
            }
        }
        catch (TaskCanceledException) { }  // Völlig in Ordnung. Die EF-Core-Abfrage wurde durch den Nutzer (Weiterscrollen) abgebrochen.
        catch (Exception ex) { Debug.WriteLine($"Fehler beim asynchronen Detail-Laden: {ex.Message}"); }
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
        if (addressBSource.Current is Adresse adresse)
        {
            foreach (var gruppe in adresse.Gruppen) { curAddressMemberships.Add(gruppe.Name); } // EF Core hat die Gruppen (hoffentlich via .Include) geladen
        }
        UpdateMembershipTags(); // UI aktualisieren
        UpdateTagComboBoxDataSource(); // Zwingt die ComboBox, die zugewiesenen Gruppen auszublenden
    }

    private void AgeLabel_MaskedTB_Set(DateOnly date)
    {
        // HIER KEIN maskedTextBox.Text ZUWEISEN! Das macht das Binding.
        var dateAsDateTime = new DateTime(date.Year, date.Month, date.Day);
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
        // HIER KEINE MASKE ODER TEXT LÖSCHEN! Das Binding setzt Text="" bei null.
        ageLabel.Text = string.Empty;
        toolTip.SetToolTip(ageLabel, string.Empty);
    }

    private void AddressDGV_DataSourceChanged(object sender, EventArgs e)
    {
        if (addressDGV.DataSource != null)
        {
            ApplyColumnSettings(addressDGV); // Einfacher Aufruf statt Tuple-Destructuring
            Text = appName + " – " + (string.IsNullOrEmpty(_databaseFilePath) ? "unbenannt" : _databaseFilePath);
        }
        else { Text = appLong; }
    }

    private void ApplyColumnSettings(DataGridView dgv)
    {
        var colCount = dgv.Columns.Count;
        if (colCount == 0) { return; }

        dgv.SuspendLayout(); // Grid einfrieren, verhindert Layout-Kettenreaktion
        try
        {
            for (var i = 0; i < colCount; i++)
            {
                if (i < _settings.HideColumnArr.Length) { dgv.Columns[i].Visible = !_settings.HideColumnArr[i]; }
                if (i < _settings.ColumnWidths.Length) { dgv.Columns[i].Width = Math.Max(20, _settings.ColumnWidths[i]); }
            }
        }
        finally
        {
            dgv.ResumeLayout(); // Grid neu zeichnen
        }
    }

    private void OpenTSButton_Click(object sender, EventArgs e) => OpenToolStripMenuItem_Click(sender, e);

    private void FrmAdressen_Resize(object sender, EventArgs e)
    {
        UpdateStatusLabelWidth();  // flexiTSStatusLabel.Width = 244 + splitContainer.SplitterDistance - 536;
        UpdateSearchBoxWidth();  // searchTSTextBox.Width = 202 + splitContainer.SplitterDistance - 536 - (tsClearLabel.Visible ? tsClearLabel.Width : 0);
    }

    private void SearchTSTextBox_TextChanged(object sender, EventArgs e)
    {
        if (!searchTSTextBox.Focused || ignoreSearchChange) { return; } // Nur reagieren, wenn der User tippt
        tsClearLabel.Visible = !string.IsNullOrWhiteSpace(searchTSTextBox.Text);
        flexiTSStatusLabel.Text = string.Empty;
        searchTimer.Stop();  // Laufenden Timer abbrechen
        searchTimer.Start();
    }

    private void ApplyGlobalSearch(string searchText, bool jumpToFirstRow = true)
    {
        var term = searchText.Trim().ToLowerInvariant();
        var isSearchEmpty = string.IsNullOrWhiteSpace(term);
        var terms = term.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        var isAddressTab = tabControl.SelectedTab == addressTabPage;
        var activeBs = isAddressTab ? addressBSource : contactBSource;
        var activeDgv = isAddressTab ? addressDGV : contactDGV;

        if (activeBs == null || activeDgv == null) { return; }

        if (activeBs.Current != null)
        {
            activeBs.EndEdit();
        }

        _isFiltering = true;

        try
        {
            activeDgv.SuspendLayout();

            if (isAddressTab && _context != null)
            {
                var source = _context.Adressen.Local;

                // 1. Wir bestimmen die Liste explizit als IList, um den Bindungs-Fehler zu vermeiden
                // WinForms braucht eine konkrete Liste für die Metadaten.
                System.Collections.IList filtered;
                if (isSearchEmpty)
                {
                    filtered = source.ToBindingList();
                }
                else
                {
                    filtered = source.Where(a => terms.All(t => a.SearchText.Contains(t))).ToList();
                }

                // 2. Flackerschutz: Nur bei echten Änderungen die DataSource tauschen
                if (activeBs.DataSource is not System.Collections.IList currentList || !currentList.Cast<Adresse>().SequenceEqual(filtered.Cast<Adresse>()))
                {
                    activeBs.DataSource = filtered;
                }
                UpdateAddressStatusBar();
            }
            else if (!isAddressTab && _allGoogleContacts != null)
            {
                System.Collections.IList filtered;
                if (isSearchEmpty)
                {
                    filtered = _allGoogleContacts;
                }
                else
                {
                    filtered = _allGoogleContacts.Where(c => terms.All(t => c.SearchText.Contains(t))).ToList();
                }

                if (activeBs.DataSource is not System.Collections.IList currentList || !currentList.Cast<Contact>().SequenceEqual(filtered.Cast<Contact>()))
                {
                    activeBs.DataSource = filtered;
                }
                UpdateContactStatusBar();
            }
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
        }
        finally
        {
            activeDgv.ResumeLayout();
            _isFiltering = false;

            if (activeBs.Count > 0)
            {
                var currentEntry = activeBs.Current;
                var shouldFocusGrid = !searchTSTextBox.Focused;

                _ = activeDgv.InvokeAsync(() =>
                {
                    if (jumpToFirstRow)
                    {
                        SyncGridToPosition(activeDgv, activeBs, 0, shouldFocusGrid);
                    }

                    if (_lastProcessedEntry != currentEntry)
                    {
                        _lastProcessedEntry = currentEntry;
                        activeBs.ResetCurrentItem();

                        if (isAddressTab)
                        {
                            AddressBindingSource_CurrentChanged(null, EventArgs.Empty);
                        }
                        else
                        {
                            ContactBindingSource_CurrentChanged(null, EventArgs.Empty);
                        }
                    }
                });
            }
            else
            {
                _lastProcessedEntry = null;
                if (isAddressTab) { AddressBindingSource_CurrentChanged(null, EventArgs.Empty); }
                else { ContactBindingSource_CurrentChanged(null, EventArgs.Empty); }
            }

            UpdateFilterUIState();
        }
    }

    private void UpdateAddressStatusBar()
    {
        if (_context == null) { return; }
        var totalCount = _context.Adressen.Local.Count;
        var visibleCount = addressBSource.Count;
        toolStripStatusLabel.Text = visibleCount == totalCount ? $"{totalCount} Adressen" : $"{visibleCount}/{totalCount} Adressen";
    }

    private void SelectFirstAddressRow()
    {
        if (addressBSource.Count > 0 && addressDGV.Rows.Count > 0)
        {
            addressDGV.ClearSelection();
            var firstCol = addressDGV.Columns.GetFirstColumn(DataGridViewElementStates.Visible);
            if (firstCol != null) { addressDGV.CurrentCell = addressDGV.Rows[0].Cells[firstCol.Index]; }
            addressDGV.Rows[0].Selected = true;
            addressBSource.Position = 0;
        }
    }

    private void SyncGridToPosition(DataGridView grid, BindingSource bs, int index, bool setFocus = false)
    {
        if (IsDisposed || grid.RowCount <= index || index < 0) { return; }
        try
        {
            if (setFocus) { grid.Focus(); }
            grid.ClearSelection();

            var firstCol = grid.Columns.GetFirstColumn(DataGridViewElementStates.Visible);
            if (firstCol is not null) { grid.CurrentCell = grid.Rows[index].Cells[firstCol.Index]; }

            grid.Rows[index].Selected = true;
            bs.Position = index;

            var firstVisible = grid.FirstDisplayedScrollingRowIndex;
            if (firstVisible >= 0)
            {
                var fullyVisibleRows = grid.DisplayedRowCount(true);
                var lastVisible = firstVisible + fullyVisibleRows - 1;

                // Nur scrollen, wenn die Zeile NICHT vollständig im aktuellen Sichtfeld liegt
                if (index < firstVisible || index > lastVisible)
                {
                    var displayedRows = grid.DisplayedRowCount(false);
                    if (displayedRows > 0)
                    {
                        var targetTopIndex = index - (displayedRows / 2);
                        targetTopIndex = Math.Max(0, Math.Min(targetTopIndex, grid.RowCount - 1));
                        grid.FirstDisplayedScrollingRowIndex = targetTopIndex;
                    }
                }
            }
        }
        catch { }  // Stille Korrektur
    }

    private void UpdateContactStatusBar()
    {
        if (_allGoogleContacts == null) { return; }
        var total = _allGoogleContacts.Count;
        var visible = contactBSource.Count;
        toolStripStatusLabel.Text = visible == total ? $"{total} Google Kontakte" : $"{visible}/{total} Google Kontakte";
    }

    private async void SaveTSButton_Click(object sender, EventArgs e)
    {
        // 1. Fall: Adressen-Tab
        if (tabControl.SelectedTab == addressTabPage && addressBSource?.Current is Adresse)
        {
            var result = await SaveSQLDatabaseAsync(false, true);
            if (result == DialogResult.Yes || result == DialogResult.None)
            {
                saveTSButton.Enabled = false;
            }
            return;
        }

        // 2. Fall: Kontakt-Tab (Guard Clause: Wenn falscher Tab oder falscher Datensatz -> Abbruch)
        if (tabControl.SelectedTab != contactTabPage || contactBSource.Current != _lastActiveContact)
        {
            Console.Beep();
            return;
        }

        // Prüfen auf valide Daten
        if (_lastActiveContact is not Contact contactToSave || _originalContactSnapshot is null)
        {
            return;
        }

        // Änderungen ermitteln
        //bool photoChanged = false;
        //if (contactDGV.CurrentRow?.IsNewRow == true) { photoChanged = true; }

        contactBSource.EndEdit();
        var changedFields = contactToSave.GetChangedFields(_originalContactSnapshot);
        //if (contactToSave is new Contact) { }
        //var photoChanged = changedFields.Remove("photos");
        //changedFields.Remove("photos");

        // Guard Clause: Nichts zu tun?
        if (changedFields.Count == 0 && !string.IsNullOrEmpty(contactToSave.ResourceName))
        {
            saveTSButton.Enabled = false;
            return;
        }

        // 3. Ausführung: Dialog-Logik ausgelagert
        // Wir übergeben die eigentliche Arbeit als Func<CancellationToken, Task>
        //var success = await Utils.RunWithProgressDialogAsync(this, "Titel", "Text", async token => { ... });

        var success = await Utils.RunWithProgressDialogAsync(this,
            "Google Kontakte",
            "Daten werden an Google übertragen.",
            async token =>
            {
                await ExecuteGoogleSaveAsync(contactToSave, changedFields, topAlignZoomPictureBox.Image, token);
            });

        if (success)
        {
            saveTSButton.Enabled = false;
            contactBSource.ResetBindings(false);
        }
    }

    private async void NewTSButton_Click(object sender, EventArgs e)
    {
        //MessageBox.Show("Achtung: Alle ungespeicherten Änderungen am aktuellen Kontakt gehen verloren!", "Neuen Kontakt erstellen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        if (!string.IsNullOrEmpty(searchTSTextBox.Text)) { Clear_SearchTextBox(); }

        if (tabControl.SelectedTab == contactTabPage)
        {
            // 1. Erst prüfen/speichern (auf dem ALTEN Kontakt).
            // Hier ist isSelectionChanging noch FALSE, damit die Prüfung läuft.
            if (!await ContactChanges_Check()) { return; }

            // 2. JETZT den Lock setzen, damit das folgende RowValidating (ausgelöst durch .Add/.Position)
            // ignoriert wird.
            isSelectionChanging = true;
            try
            {
                var newContact = new Contact();

                // 3. Hinzufügen 
                contactBSource.Add(newContact); // Löst evtl. Events aus -> werden durch Lock ignoriert
                contactBSource.ResetBindings(false);

                var realIndex = contactBSource.IndexOf(newContact);

                if (realIndex >= 0)
                {
                    // 4. Position wechseln
                    contactBSource.Position = realIndex; // Löst RowValidating aus -> wird durch Lock ignoriert

                    if (contactDGV.RowCount > realIndex)
                    {
                        contactDGV.FirstDisplayedScrollingRowIndex = realIndex;
                        contactDGV.Rows[realIndex].Selected = true;
                    }
                }

                // 5. Interne Referenzen auf den NEUEN Kontakt biegen
                _lastProcessedEntry = null; // UI-Update erzwingen
                _lastActiveContact = newContact;
                _originalContactSnapshot = (Contact)newContact.Clone();

                // UI Updates...
                _ = ShowPhotoInPictureBox(newContact, CancellationToken.None);
                UpdateMembershipTags();
                UpdateSaveButton();
            }
            finally { isSelectionChanging = false; }

            if (cbAnrede.CanFocus) { cbAnrede.Focus(); }
        }
        else if (tabControl.SelectedTab == addressTabPage && addressBSource != null)
        {
            _lastProcessedEntry = null; // UI-Update erzwingen
            addressBSource.AddNew();  // noch nicht fest in die zugrunde liegende BindingList "committed".
            addressBSource.EndEdit(); // dadurch wird der Status im EF ChangeTracker zuverlässig auf 'Added' gesetzt
            UpdateSaveButton();
            if (cbAnrede.CanFocus) { cbAnrede.Focus(); }
        }
    }

    private async void CopyTSButton_Click(object sender, EventArgs e)
    {
        // 1. Der Gatekeeper: Prüft zuerst, ob der aktuelle Google-Kontakt 
        // ungespeicherte Änderungen hat, bevor wir etwas Neues erstellen.
        await CheckContactChanges(async () =>
        {
            // 2. LOCK SETZEN: Verhindert, dass RowValidating während des 
            // programmatischen Zeilenwechsels dazwischenfunkt.
            isSelectionChanging = true;
            try
            {
                // ==============================================================================
                // FALL 1: Google Kontakt duplizieren
                // ==============================================================================
                if (tabControl.SelectedTab == contactTabPage && contactBSource.Current is Contact originalContact)
                {
                    // Klonen (ResourceName/ETag leeren für neuen Datensatz)
                    var clone = (Contact)originalContact.Clone();
                    clone.ResourceName = string.Empty;
                    clone.ETag = string.Empty;
                    clone.PhotoUrl = null;

                    _allGoogleContacts ??= [];
                    _allGoogleContacts.Add(clone);

                    // Sortieren und Bindings aktualisieren
                    Utils.SortContacts(_allGoogleContacts);
                    contactBSource.ResetBindings(false);

                    // Position finden und ansteuern
                    var newIndex = _allGoogleContacts.IndexOf(clone);
                    if (newIndex >= 0)
                    {
                        _lastProcessedEntry = null; // UI-Update erzwingen
                        contactBSource.Position = newIndex;

                        if (contactDGV.RowCount > 0 && newIndex < contactDGV.RowCount)
                        {
                            // Kontext wahren: 2 Zeilen Puffer nach oben
                            var scrollIndex = Math.Max(0, newIndex - 2);
                            contactDGV.FirstDisplayedScrollingRowIndex = scrollIndex;
                            contactDGV.Rows[newIndex].Selected = true;

                            // Erste sichtbare Zelle fokussieren
                            var firstCol = contactDGV.Columns.GetFirstColumn(DataGridViewElementStates.Visible);
                            if (firstCol != null)
                            {
                                contactDGV.CurrentCell = contactDGV.Rows[newIndex].Cells[firstCol.Index];
                            }
                        }
                    }

                    // Snapshots für den neuen Klon initialisieren
                    _lastActiveContact = clone;
                    _originalContactSnapshot = (Contact)clone.Clone();

                    saveTSButton.Enabled = true;
                    cbAnrede.Focus();
                }

                // ==============================================================================
                // FALL 2: Lokale Adresse duplizieren
                // ==============================================================================
                else if (tabControl.SelectedTab == addressTabPage && addressBSource?.Current is Adresse originalAdresse && _context != null)
                {
                    // Sauberes EF-Cloning via AsNoTracking
                    var duplikat = _context.Adressen
                        .Include(a => a.Foto)
                        .AsNoTracking()
                        .FirstOrDefault(a => a.Id == originalAdresse.Id);

                    if (duplikat == null)
                    {
                        return;
                    }

                    _lastProcessedEntry = null;
                    duplikat.Id = 0;
                    duplikat.Foto?.Id = 0;

                    // Einfügeposition bestimmen
                    var insertIndex = Utils.GetAddressInsertIndex(addressBSource, duplikat);

                    // In BindingSource einfügen
                    addressBSource.Insert(insertIndex, duplikat);
                    addressBSource.Position = insertIndex;

                    // UI Scrollen & Fokus
                    if (addressDGV.RowCount > 0 && insertIndex < addressDGV.RowCount)
                    {
                        var scrollIndex = Math.Max(0, insertIndex - 2);
                        addressDGV.FirstDisplayedScrollingRowIndex = scrollIndex;
                        addressDGV.Rows[insertIndex].Selected = true;

                        var firstCol = addressDGV.Columns.GetFirstColumn(DataGridViewElementStates.Visible);
                        if (firstCol != null)
                        {
                            addressDGV.CurrentCell = addressDGV.Rows[insertIndex].Cells[firstCol.Index];
                        }
                    }

                    saveTSButton.Enabled = true;
                    cbAnrede.Focus();
                }
                else
                {
                    Console.Beep();
                }
            }
            catch (Exception ex)
            {
                Utils.ErrTaskDlg(Handle, ex);
            }
            finally
            {
                // 3. LOCK AUFHEBEN: Ab jetzt sind manuelle Zeilenwechsel wieder bewacht.
                isSelectionChanging = false;
            }
        });
    }

    private async void CopyToOtherDGVMenuItem_Click(object sender, EventArgs e)
    {
        // ==============================================================================
        // FALL 1: Von Google (Contact) -> Lokal (Adresse)
        // ==============================================================================
        if (tabControl.SelectedTab == contactTabPage && contactBSource.Current is Contact selectedGoogleContact)
        {
            // A. Sofortiges Feedback
            tabControl.SelectedTab = addressTabPage;
            if (!string.IsNullOrEmpty(searchTSTextBox.Text)) { Clear_SearchTextBox(); }

            // B. Arbeit erledigen
            var success = await CopyGoogleToLocalAsync(selectedGoogleContact);

            // C. Nachbearbeitung
            if (success)
            {
                if (addressDGV.RowCount > 0)
                {
                    var currentIdx = addressBSource.Position;
                    if (currentIdx >= 0 && currentIdx < addressDGV.RowCount)
                    {
                        // 1. Scrollen (funktioniert immer)
                        addressDGV.FirstDisplayedScrollingRowIndex = currentIdx;

                        // 2. Zeile markieren
                        addressDGV.Rows[currentIdx].Selected = true;

                        // 3. Fokus auf erste SICHTBARE Zelle setzen (Fix für den Absturz)
                        var firstVisibleCol = addressDGV.Columns.GetFirstColumn(DataGridViewElementStates.Visible);
                        if (firstVisibleCol != null)
                        {
                            addressDGV.CurrentCell = addressDGV.Rows[currentIdx].Cells[firstVisibleCol.Index];
                        }
                    }
                }

                cbAnrede.Focus();
                saveTSButton.Enabled = true;
            }
            else
            {
                tabControl.SelectedTab = contactTabPage;
            }
        }

        // ==============================================================================
        // FALL 2: Von Lokal (Adresse) -> Google (Contact)
        // ==============================================================================
        else if (tabControl.SelectedTab == addressTabPage && addressBSource.Current is Adresse selectedLocalAddress)
        {
            // A. Sofortiges Feedback
            tabControl.SelectedTab = contactTabPage;
            if (!string.IsNullOrEmpty(searchTSTextBox.Text)) { Clear_SearchTextBox(); }

            // B. Arbeit erledigen
            var success = await CopyLocalToGoogleAsync(selectedLocalAddress);

            // C. Nachbearbeitung
            if (success)
            {
                if (contactDGV.RowCount > 0)
                {
                    var currentIdx = contactBSource.Position;
                    if (currentIdx >= 0 && currentIdx < contactDGV.RowCount)
                    {
                        // 1. Scrollen
                        contactDGV.FirstDisplayedScrollingRowIndex = currentIdx;

                        // 2. Zeile markieren
                        contactDGV.Rows[currentIdx].Selected = true;

                        // 3. Fokus auf erste SICHTBARE Zelle setzen (Fix für den Absturz)
                        var firstVisibleCol = contactDGV.Columns.GetFirstColumn(DataGridViewElementStates.Visible);
                        if (firstVisibleCol != null) { contactDGV.CurrentCell = contactDGV.Rows[currentIdx].Cells[firstVisibleCol.Index]; }
                    }
                }

                cbAnrede.Focus();
                saveTSButton.Enabled = false;
                flexiTSStatusLabel.Text = "Kontakt erfolgreich zu Google kopiert.";
            }
            else { tabControl.SelectedTab = addressTabPage; }
        }
        else { Console.Beep(); }
    }

    private async void DeleteTSButton_Click(object sender, EventArgs e)
    {
        // 1. Der Gatekeeper: Prüft auf ungespeicherte Änderungen
        await CheckContactChanges(async () =>
        {
            // === FALL A: GOOGLE KONTAKTE ===
            if (tabControl.SelectedTab == contactTabPage && contactBSource.Current is Contact googleKontakt)
            {
                var (askBefore, deleteNow) = Utils.AskBeforeDeleteContact(Handle, googleKontakt, _settings.AskBeforeDelete, false);
                _settings.AskBeforeDelete = askBefore;

                if (!deleteNow)
                {
                    return;
                }

                isSelectionChanging = true;
                try
                {
                    var success = await Utils.RunWithProgressDialogAsync(
                        this,
                        "Kontakt löschen",
                        "Der Kontakt wird bei Google gelöscht...",
                        async token =>
                        {
                            await DeleteGoogleContactAsync(googleKontakt, token);
                        });

                    if (success)
                    {
                        _lastProcessedEntry = null; // Cache leeren!
                        contactBSource.Remove(googleKontakt);

                        _lastActiveContact = null;
                        _originalContactSnapshot = null;

                        UpdateContactStatusBar();

                        // Synchronisation: Wir springen auf die neue Position (die BindingSource bleibt am gleichen Index)
                        if (contactBSource.Count > 0)
                        {
                            var newPos = contactBSource.Position;
                            _ = contactDGV.InvokeAsync(() => SyncGridToPosition(contactDGV, contactBSource, newPos, true));
                        }
                    }
                }
                finally
                {
                    isSelectionChanging = false;
                }
            }
            // === FALL B: LOKALE ADRESSEN ===
            else if (tabControl.SelectedTab == addressTabPage && addressBSource.Current is Adresse adresseZumLoeschen && _context != null)
            {
                if (addressBSource.IsBindingSuspended || adresseZumLoeschen == null)
                {
                    return;
                }

                if (addressDGV.CurrentRow?.IsNewRow == true)
                {
                    return;
                }

                addressBSource.EndEdit();
                var deleteFinal = true;

                if (_settings.AskBeforeDelete)
                {
                    var (askBefore, deleteNow) = Utils.AskBeforeDeleteAddress(Handle, adresseZumLoeschen, _settings.AskBeforeDelete);
                    _settings.AskBeforeDelete = askBefore;
                    deleteFinal = deleteNow;
                }

                if (!deleteFinal)
                {
                    return;
                }

                isSelectionChanging = true; // Guard jetzt auch hier für lokale Adressen!
                try
                {
                    _lastProcessedEntry = null; // Cache leeren!
                    var entry = _context.Entry(adresseZumLoeschen);
                    var isNewRecord = entry.State == EntityState.Added || adresseZumLoeschen.Id == 0;

                    if (isNewRecord)
                    {
                        if (adresseZumLoeschen.Foto is not null)
                        {
                            var fotoEntry = _context.Entry(adresseZumLoeschen.Foto);
                            if (fotoEntry.State == EntityState.Added || adresseZumLoeschen.Foto.Id == 0)
                            {
                                fotoEntry.State = EntityState.Detached;
                            }
                        }
                        entry.State = EntityState.Detached;
                    }
                    else
                    {
                        _context.Adressen.Remove(adresseZumLoeschen);
                    }

                    if (addressBSource.Contains(adresseZumLoeschen))
                    {
                        addressBSource.Remove(adresseZumLoeschen);
                    }

                    UpdateSaveButton();
                    UpdateAddressStatusBar();

                    // Synchronisation: Fokus sicher auf das nächste Element setzen
                    if (addressBSource.Count > 0)
                    {
                        var newPos = addressBSource.Position;
                        _ = addressDGV.InvokeAsync(() => SyncGridToPosition(addressDGV, addressBSource, newPos, true));
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
            else
            {
                Console.Beep();
            }
        });
    }

    private void ExecuteAndPreserveSelection<T>(BindingSource bindingSource, DataGridView grid, Action dataUpdateAction) where T : class
    {
        if (grid == null || bindingSource == null)
        {
            return;
        }

        // 1. Snapshot VOR der Änderung (Referenz sichern)
        // Das 'currentItem' wird hier deklariert und behält seinen Wert
        var currentItem = bindingSource.Current as T;

        var currencyManager = BindingContext?[bindingSource] as CurrencyManager;
        currencyManager?.SuspendBinding();

        try
        {
            grid.CurrentCell = null;
            dataUpdateAction();
        }
        finally
        {
            currencyManager?.ResumeBinding();
        }

        // 3. Selektion mit Zentrierung wiederherstellen
        // Hier nutzen wir den modernen Pattern Matching Check
        if (currentItem is T target)
        {
            var newIndex = bindingSource.IndexOf(target);
            if (newIndex >= 0)
            {
                var shouldFocusGrid = !searchTSTextBox.Focused;
                _ = grid.InvokeAsync(() => SyncGridToPosition(grid, bindingSource, newIndex, shouldFocusGrid));
            }
        }
        else if (grid.RowCount > 0)
        {
            _ = grid.InvokeAsync(SelectFirstAddressRow);
        }
    }

    private async void FrmAdressen_FormClosing(object sender, FormClosingEventArgs e)
    {
        // 1. Rekursions-Check: Wenn wir am Ende der Methode Close() rufen, springen wir hier raus
        // und Windows darf das Fenster nun endgültig schließen.
        if (_isClosing) { return; }

        // 2. DAS WICHTIGSTE: Den synchronen Schließvorgang SOFORT stoppen!
        // Nur so überlebt das Fenster die asynchronen Speicher-Dialoge.
        e.Cancel = true;

        // 3. Laufende Google-Requests sofort abbrechen
        _googleCts?.Cancel();

        // -------------------------------------------------------------
        // SCHRITT A: Prüfungen durchführen (Abbruch ermöglichen)
        // -------------------------------------------------------------

        // Fall 1: SQL Datenbank (EF Core 10)
        if (_context != null)
        {
            var result = await SaveSQLDatabaseAsync(false, false, true);
            // Wenn der User abbricht, machen wir nichts weiter. 
            // Das Fenster bleibt offen (da e.Cancel bereits auf true steht).
            if (result == DialogResult.Cancel) { return; }
        }

        // Fall 2: Google Kontakte (Zentraler Gatekeeper)
        // isClosing: true teilt dem Gatekeeper mit, dass keine UI-Resets mehr nötig sind
        var readyToCloseGoogle = await ContactChanges_Check(isClosing: true);
        if (!readyToCloseGoogle) { return; }

        // -------------------------------------------------------------
        // SCHRITT B: Aufräumen und Endgültig Schließen
        // -------------------------------------------------------------

        // Ab hier gibt es kein Zurück mehr. UI einfrieren für den Cleanup.
        AutoValidate = AutoValidate.Disable;
        Enabled = false;
        Cursor = Cursors.WaitCursor;

        try
        {
            SaveConfiguration();

            // Ressourcen sauber freigeben
            _googleCts?.Dispose();
            CloseDatabaseConnection();

            addressBSource?.Dispose();
            contactBSource?.Dispose();

            // Timer stoppen und entsorgen
            searchTimer?.Dispose();
            debounceTimer?.Dispose();
            //scrollTimer?.Dispose();
            _monospaceFont?.Dispose();
            _standardAppFont?.Dispose();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Fehler beim Cleanup: {ex.Message}");
        }
        finally
        {
            // 4. Finales Flag setzen, Cursor zurücksetzen und Schließen neu triggern
            _isClosing = true;
            Cursor = Cursors.Default;

            // Jetzt rufen wir Close() erneut auf. Das Event feuert wieder,
            // läuft aber oben in 'if (_isClosing) { return; }' rein und schließt die App sauber.
            Close();
        }
    }

    private void AboutToolStripMenuItem_Click(object sender, EventArgs e) => Utils.HelpMsgTaskDlg(Handle, appLong, Icon, _currentDbVersion);

    private void AddressDGV_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e) => toolStripStatusLabel.Text = addressDGV.RowCount.ToString() + " Adressen";

    private void AddressDGV_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e) => toolStripStatusLabel.Text = addressDGV.RowCount.ToString() + " Adressen";

    private void ErzeugeGrussformeln(object? item)
    {
        if (item == null)
        {
            if (cbGrussformel.AutoCompleteCustomSource.Count > 0)
            {
                SetGrussformelAutoCompleteSilently(new AutoCompleteStringCollection());
            }
            return;
        }

        string? vName = null, nName = null, nick = null, tit = null;
        if (item is Adresse a) { (vName, nName, nick, tit) = (a.Vorname, a.Nachname, a.Nickname, a.Praefix); }
        else if (item is Contact c) { (vName, nName, nick, tit) = (c.Vorname, c.Nachname, c.Nickname, c.Praefix); }

        var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(vName)) { replacements["#vorname"] = vName; }
        if (!string.IsNullOrWhiteSpace(nick)) { replacements["#nickname"] = nick; }
        if (!string.IsNullOrWhiteSpace(nName)) { replacements["#nachname"] = nName; }
        if (!string.IsNullOrWhiteSpace(tit)) { replacements["#titel"] = tit; }

        if (replacements.Count == 0)
        {
            if (cbGrussformel.AutoCompleteCustomSource.Count > 0)
            {
                SetGrussformelAutoCompleteSilently(new AutoCompleteStringCollection());
            }
            return;
        }

        var newSuggestions = grussformelList
            .Select(template => PlaceholderRegex().Replace(template, match =>
                replacements.TryGetValue(match.Value, out var replacement) ? replacement : match.Value))
            .Where(text => !text.Contains('#'))
            .Distinct()
            .Order()
            .ToArray();

        var currentSource = cbGrussformel.AutoCompleteCustomSource.Cast<string>().ToArray();
        if (!newSuggestions.SequenceEqual(currentSource))
        {
            var newCollection = new AutoCompleteStringCollection();
            newCollection.AddRange(newSuggestions);
            SetGrussformelAutoCompleteSilently(newCollection);
        }
    }

    private void SetGrussformelAutoCompleteSilently(AutoCompleteStringCollection newCollection)
    {
        // 1. Zeichen-Updates für die TextBox hart auf Betriebssystem-Ebene verbieten
        NativeMethods.SendMessage(cbGrussformel.Handle, NativeMethods.WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);

        // 2. Quelle austauschen (WinForms baut im Hintergrund um, aber die UI bleibt stumm)
        cbGrussformel.AutoCompleteCustomSource = newCollection;

        // 3. Zeichen-Updates wieder erlauben und exakt EINMAL sauber neu zeichnen lassen
        NativeMethods.SendMessage(cbGrussformel.Handle, NativeMethods.WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
        cbGrussformel.Invalidate();
    }

    private void ImportToolStripMenuItem_Click(object sender, EventArgs e)
    {
        // Wir übergeben das vorhandene Array. "Gruppen" hängen wir an, damit der Nutzer auch kommagetrennte Gruppen mappen kann.
        var availableFields = dataFields.ToList();
        availableFields.Add("Gruppen");

        // Wir übergeben auch den aktuellen Datenbankpfad, falls der Nutzer "In aktuelle DB importieren" wählt
        var importDialog = new FrmImportCsv(availableFields, _databaseFilePath);

        if (importDialog.ShowDialog(this) == DialogResult.OK)
        {
            // Wenn der Import (und ggf. das Anlegen einer neuen DB) erfolgreich war, verbinden wir uns neu bzw. laden die Ansicht neu.
            var targetDbPath = importDialog.TargetDatabasePath;
            _ = ConnectSQLDatabaseAsync(targetDbPath);
        }
    }

    private void SearchTSTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.Tab)
        {
            tabControl.SelectedIndex = (tabControl.SelectedIndex == 1) ? 0 : 1;
            e.SuppressKeyPress = true;  // Ton unterdrücken
            e.Handled = true;  // als "erledigt" markieren
        }
        else if (e.KeyCode == Keys.Enter)
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

    private async void HandleSwitchDatabaseAsync(string currentDbPath)
    {
        foreach (var file in _settings.RecentFiles)
        {
            if (file == currentDbPath) { continue; }

            if (File.Exists(file))
            {
                if (addressBSource != null) { await SaveSQLDatabaseAsync(true); }
                //ConnectSQLDatabase(file);  // Erst wenn das Speichern fertig ist, geht es hier weiter:
                await ConnectSQLDatabaseAsync(file);
                SetSearchTextIgnoreChange(string.Empty);
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
                WordManager.ShowWordBookmarksInfoDialog(Handle, [.. bookmarkTextDictionary.Keys]);
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
                Utils.StartFile(Handle, Path.Combine(Path.GetDirectoryName(appPath) ?? string.Empty, "AdressenKontakte.pdf"));
                return true;
            case Keys.I | Keys.Control:
                Utils.HelpMsgTaskDlg(Handle, appLong, Icon, _currentDbVersion);
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
            case Keys.F9 | Keys.Alt:
                filterToolStripMenuItem.ShowDropDown();
                return true;
            case Keys.Enter | Keys.Control:
            case Keys.Tab | Keys.Control:   // funktioniert nicht
                tabControl.SelectedIndex = tabControl.SelectedIndex == 1 ? 0 : 1;
                return true;
            case Keys.F | Keys.Control:
                if (tbNotizen.Focused)
                {
                    OpenSearchDialog();
                }
                else if (dokuListView.Focused)
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
                BirthdayReminder(tabControl.SelectedTab == addressTabPage ? addressDGV : contactDGV, showIfEmpty: true);
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
                RejectChangesToolStripMenuItem_Click(null!, null!);
                return true;
            case Keys.Delete | Keys.Control:
                DeleteTSButton_Click(null!, null!);
                return true;
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
            case Keys.H | Keys.Control:
                if (min2TrayTSButton.Visible)
                {
                    HideToTray();
                }
                return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void TextBox_KeyDown(object sender, KeyEventArgs e)
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
            tbMail1.Focus();  // SelectNextControl((Control)sender, true, true, true, true);
        }
        else if (e.KeyCode == Keys.Space)
        {
            e.SuppressKeyPress = true;
            BtnCalendar_Click(null!, null!);
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
        if (tabControl.SelectedTab == addressTabPage && addressDGV.SelectedRows.Count > 0 || tabControl.SelectedTab == contactTabPage && addressDGV.SelectedRows.Count > 0)
        {
            var useWord = _settings.WordProcessorProgram ?? Utils.AskWordProcessingProgram(Handle);
            if (useWord is null) { return; }
            if (useWord == true)
            {
                if (!WordManager.IsWordInstalled)
                {
                    Utils.MsgTaskDlg(Handle, "Word fehlt", "Microsoft Word wurde nicht gefunden. Bitte installiere es.");
                    return;
                }
                WordProcess();
            }
            else
            {
                if (!WordManager.IsLibreOfficeInstalled)
                {
                    Utils.MsgTaskDlg(Handle, "LibreOffice fehlt", "LibreOffice Writer wurde nicht gefunden. Bitte installiere es.");
                    return;
                }
                LibreProcess();
            }
        }
        else { Utils.MsgTaskDlg(Handle, "Keine Auswahl", "Es könne keine Daten übertragen werden."); }
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
        FillDictionary(); // bookmarkTextDictionary mit aktuellen Werten füllen
        WordManager.TransferDataToActiveDocument(bookmarkTextDictionary, Handle);
    }

    private void FillDictionary()
    {  //Alle Werte einmalig sauber auslesen und trimmen; das verbessert die Lesbarkeit und Performance
        var anrede = cbAnrede.Text.Trim();
        var praefix = cbPraefix.Text.Trim();
        var vorname = tbVorname.Text.Trim();
        var zwischen = tbZwischenname.Text.Trim();
        var nachname = tbNachname.Text.Trim();
        var nickname = tbNickname.Text.Trim();
        var suffix = tbSuffix.Text.Trim();
        var firma = tbFirma.Text.Trim();
        var position = tbPosition.Text.Trim();
        var strasse = tbStraße.Text.Trim();
        var postfach = tbPostfach.Text.Trim();
        var plz = cbPLZ.Text.Trim();
        var ort = cbOrt.Text.Trim();
        var land = cbLand.Text.Trim();
        var betreff = tbBetreff.Text.Trim();
        var gruss = cbGrussformel.Text.Trim();
        var schluss = cbSchlussformel.Text.Trim();
        var geburtstag = maskedTextBox.Text.Trim();
        var mail1 = tbMail1.Text.Trim();
        var mail2 = tbMail2.Text.Trim();
        var tel1 = tbTelefon1.Text.Trim();
        var tel2 = tbTelefon2.Text.Trim();
        var mobil = tbMobil.Text.Trim();
        var fax = tbFax.Text.Trim();
        var internet = tbInternet.Text.Trim();

        var zwischenInitial = string.IsNullOrEmpty(zwischen) ? null : $"{zwischen[0]}.";

        bookmarkTextDictionary["Anrede"] = anrede;
        bookmarkTextDictionary["Praefix"] = praefix; // Empfehlung: "ae" statt "ä"
        bookmarkTextDictionary["Vorname"] = vorname;
        bookmarkTextDictionary["Zwischenname"] = zwischen;
        bookmarkTextDictionary["Zwischenname_initial"] = zwischenInitial ?? ""; // Falls null, leerer String
        bookmarkTextDictionary["Nickname"] = nickname;
        bookmarkTextDictionary["Nachname"] = nachname;
        bookmarkTextDictionary["Suffix"] = suffix;
        bookmarkTextDictionary["Unternehmen"] = firma;
        bookmarkTextDictionary["Position"] = position;

        bookmarkTextDictionary["Anrede_Praefix_Vorname_Nachname"] =
            string.Join(" ", new[] { anrede, praefix, vorname, nachname }.Where(static s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Anrede_Praefix_Vorname_Zwischenname_Nachname"] =
            string.Join(" ", new[] { anrede, praefix, vorname, zwischen, nachname }.Where(static s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Anrede_Praefix_Vorname_Zwischenname_initial_Nachname"] =
            string.Join(" ", new[] { anrede, praefix, vorname, zwischenInitial, nachname }.Where(static s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Praefix_Vorname_Nachname"] =
            string.Join(" ", new[] { praefix, vorname, nachname }.Where(static s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Praefix_Vorname_Zwischenname_Nachname"] =
            string.Join(" ", new[] { praefix, vorname, zwischen, nachname }.Where(static s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Praefix_Vorname_Zwischenname_initial_Nachname"] =
            string.Join(" ", new[] { praefix, vorname, zwischenInitial, nachname }.Where(static s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Vorname_Nachname"] =
            string.Join(" ", new[] { vorname, nachname }.Where(static s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Vorname_Zwischenname_Nachname"] =
            string.Join(" ", new[] { vorname, zwischen, nachname }.Where(static s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Vorname_Zwischenname_initial_Nachname"] =
            string.Join(" ", new[] { vorname, zwischenInitial, nachname }.Where(static s => !string.IsNullOrWhiteSpace(s)));

        // Adressdaten
        bookmarkTextDictionary["Strasse"] = strasse; // "ss" statt "ß" ist in Keys sicherer
        bookmarkTextDictionary["Postfach"] = postfach;

        // Logik: Wenn Postfach leer, nimm Straße. Sonst "Postfach XYZ"
        bookmarkTextDictionary["Postfach_sonst_Strasse"] = string.IsNullOrEmpty(postfach) ? strasse : $"Postfach {postfach}";

        bookmarkTextDictionary["PLZ"] = plz;
        bookmarkTextDictionary["Ort"] = ort;
        bookmarkTextDictionary["PLZ_Ort"] = $"{plz} {ort}".Trim();

        bookmarkTextDictionary["Land"] = land;
        bookmarkTextDictionary["Land_Gross"] = land.ToUpper();

        // Sonstiges
        bookmarkTextDictionary["Betreff"] = betreff;
        bookmarkTextDictionary["Grussformel"] = gruss; // "ss" statt "ß"
        bookmarkTextDictionary["Schlussformel"] = schluss;
        bookmarkTextDictionary["Geburtstag"] = geburtstag;
        bookmarkTextDictionary["Mail1"] = mail1;
        bookmarkTextDictionary["Mail2"] = mail2;
        bookmarkTextDictionary["Telefon1"] = tel1;
        bookmarkTextDictionary["Telefon2"] = tel2;
        bookmarkTextDictionary["Mobil"] = mobil;
        bookmarkTextDictionary["Fax"] = fax;
        bookmarkTextDictionary["Internet"] = internet;
    }

    private void WordHelpToolStripMenuItem_Click(object sender, EventArgs e)
    {
        FillDictionary();
        WordManager.ShowWordBookmarksInfoDialog(Handle, [.. bookmarkTextDictionary.Keys]);
    }

    private void StatusbarToolStripMenuItem_Click(object sender, EventArgs e) => statusStrip.Visible = statusbarToolStripMenuItem.Checked = !statusbarToolStripMenuItem.Checked;
    private void NewToolStripMenuItem_Click(object sender, EventArgs e) => NewTSButton_Click(sender, e);
    private void DuplicateToolStripMenuItem_Click(object sender, EventArgs e) => CopyTSButton_Click(sender, e);
    private void DeleteToolStripMenuItem_Click(object sender, EventArgs e) => DeleteTSButton_Click(sender, e);

    private void SwitchDataBinding(BindingSource targetSource)
    {
        if (targetSource == null || (targetSource.DataSource == null && targetSource == contactBSource)) { return; }
        var useNullConversion = targetSource == addressBSource;  // Unterscheidung: Lokale DB (null erlaubt) vs. Google (leerer String bevorzugt)
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
        UpdateTextBoxAutoComplete(targetSource); // Aktualisierung der ComboBox-Listen (Suggest-Listen)
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

    private void UpdateTextBoxAutoComplete(BindingSource targetSource)
    {
        cbAnrede.AutoCompleteCustomSource.Clear();
        cbPraefix.AutoCompleteCustomSource.Clear();
        cbPLZ.AutoCompleteCustomSource.Clear();
        cbOrt.AutoCompleteCustomSource.Clear();
        cbLand.AutoCompleteCustomSource.Clear();
        cbSchlussformel.AutoCompleteCustomSource.Clear();
        cbGrussformel.AutoCompleteCustomSource.Clear();

        // HashSets ignorieren Duplikate blitzschnell und automatisch
        var anreden = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var praefixe = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var plzs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var orte = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var laender = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var schlussformeln = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (targetSource == addressBSource && _context != null)
        {
            // Nur EIN Durchlauf durch die EF Core Daten
            foreach (var item in _context.Adressen.Local)
            {
                if (!string.IsNullOrWhiteSpace(item.Anrede)) { anreden.Add(item.Anrede); }
                if (!string.IsNullOrWhiteSpace(item.Praefix)) { praefixe.Add(item.Praefix); }
                if (!string.IsNullOrWhiteSpace(item.PLZ)) { plzs.Add(item.PLZ); }
                if (!string.IsNullOrWhiteSpace(item.Ort)) { orte.Add(item.Ort); }
                if (!string.IsNullOrWhiteSpace(item.Land)) { laender.Add(item.Land); }
                if (!string.IsNullOrWhiteSpace(item.Schlussformel)) { schlussformeln.Add(item.Schlussformel); }
            }
        }
        else if (targetSource == contactBSource && contactBSource.DataSource is BindingList<Contact> contactList)
        {
            // Nur EIN Durchlauf durch die Google Kontakte
            foreach (var item in contactList)
            {
                if (!string.IsNullOrWhiteSpace(item.Anrede)) { anreden.Add(item.Anrede); }
                if (!string.IsNullOrWhiteSpace(item.Praefix)) { praefixe.Add(item.Praefix); }
                if (!string.IsNullOrWhiteSpace(item.PLZ)) { plzs.Add(item.PLZ); }
                if (!string.IsNullOrWhiteSpace(item.Ort)) { orte.Add(item.Ort); }
                if (!string.IsNullOrWhiteSpace(item.Land)) { laender.Add(item.Land); }
                if (!string.IsNullOrWhiteSpace(item.Schlussformel)) { schlussformeln.Add(item.Schlussformel); }
            }
        }

        // Am Ende einmalig sortieren und zuweisen
        cbAnrede.AutoCompleteCustomSource.AddRange([.. anreden.Order()]);
        cbPraefix.AutoCompleteCustomSource.AddRange([.. praefixe.Order()]);
        cbPLZ.AutoCompleteCustomSource.AddRange([.. plzs.Order()]);
        cbOrt.AutoCompleteCustomSource.AddRange([.. orte.Order()]);
        cbLand.AutoCompleteCustomSource.AddRange([.. laender.Order()]);
        cbSchlussformel.AutoCompleteCustomSource.AddRange([.. schlussformeln.Order()]);
    }

    private async Task ShowPhotoInPictureBox(object item, CancellationToken token)
    {
        // 1. Reset
        topAlignZoomPictureBox.Image = tabControl.SelectedTab == contactTabPage ? Resources.ContactBild100 : Resources.AddressBild100;
        delPictboxToolStripButton.Enabled = false;

        if (item is IContactEntity entity)
        {
            try
            {
                // --- EF Core "Explicit Loading" für SQL-Adressen ---
                if (item is Adresse adresse && _context != null)
                {
                    var entry = _context.Entry(adresse);
                    if (!entry.Reference(a => a.Foto).IsLoaded)
                    {
                        // Das Laden aus der Datenbank sofort abbrechen, falls gescrollt wird
                        await entry.Reference(a => a.Foto).LoadAsync(token);
                    }
                }

                // 3. Bild abrufen (Lokal oder Web) und Token DURCHREICHEN!
                var image = await entity.GetPhotoAsync(token);

                // 4. Harter Cut: Wenn der Task als "abgebrochen" markiert ist,
                // entsorgen wir das Bild sofort.
                if (token.IsCancellationRequested)
                {
                    image?.Dispose();
                    return;
                }

                // 5. Redundante Prüfung (zur absoluten Sicherheit)
                var currentBindingSource = tabControl.SelectedTab == addressTabPage
                    ? addressBSource
                    : contactBSource;

                if (currentBindingSource.Current != item)
                {
                    image?.Dispose();
                    return;
                }

                // 6. Anzeigen
                if (image != null)
                {
                    topAlignZoomPictureBox.Image = image;
                    delPictboxToolStripButton.Enabled = true;
                }
            }
            catch (TaskCanceledException)
            {
                // Völlig normal: Das Bildladen wurde durch Weiterscrollen abgebrochen.
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Fehler beim Laden des Fotos: " + ex.Message);
            }
        }
    }

    private async Task LoadAndDisplayGoogleContactsAsync()
    {
        // A. UI Status Checks & Datensicherheit
        if (tabControl.SelectedTab == addressTabPage && addressBSource != null)
        {
            if (filterRemoveToolStripMenuItem.Visible)
            {
                FilterRemoveToolStripMenuItem_Click(null!, EventArgs.Empty);
            }

            if (searchTSTextBox.TextBox.TextLength > 0)
            {
                lastAddressSearch = searchTSTextBox.TextBox.Text;
                SetSearchTextIgnoreChange(string.Empty);
            }
        }
        else
        {
            // DER NEUE GATEKEEPER: Erst prüfen, ob noch Änderungen offen sind.
            if (!await ContactChanges_Check())
            {
                return;
            }

            lastContactSearch = searchTSTextBox.TextBox.Text;
            SetSearchTextIgnoreChange(string.Empty);
        }

        // B. Netzwerk-Check (JETZT ASYNCHRON!)
        if (!await Utils.GoogleConnectionCheckAsync(Handle, secretPath))
        {
            return;
        }
        else
        {
            topAlignZoomPictureBox.Image = Resources.ContactBild100;
        }

        // 1. Alten Prozess abbrechen UND aufräumen
        if (_googleCts != null)
        {
            _googleCts.Cancel();
            _googleCts.Dispose();
        }

        // 2. Neues Token erstellen
        _googleCts = new CancellationTokenSource();
        var ct = _googleCts.Token; // Token in lokale Var

        _isFiltering = true;
        try
        {
            var tokenFileName = "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";
            var tokenFilePath = Path.Combine(tokenDir, tokenFileName);
            var isNewLogin = !File.Exists(tokenFilePath);

            toolStripStatusLabel.Text = "Verbindung zu Google wird hergestellt...";
            toolStripProgressBar.Style = ProgressBarStyle.Continuous;
            toolStripProgressBar.Value = 15;
            toolStripProgressBar.Visible = true;

            var manager = new GooglePeopleManager(secretPath, tokenDir);
            var stopwatch = Stopwatch.StartNew();

            // C. MANAGER AUFRUF
            var result = await manager.LoadContactsAsync(ct);
            toolStripProgressBar.Value = 30;
            stopwatch.Stop();

            // D. Auth-Logik
            if (isNewLogin || stopwatch.ElapsedMilliseconds > 2000)
            {
                contactBirthdayFlag = false;
            }

            // E. Gruppen verarbeiten (KORRIGIERT: Duplikate beim Stern vermeiden)
            contactGroupsDict = result.GroupMap;
            allContactMemberships.Clear();
            flowLayoutPanel.Controls.Clear();

            // Erst alle regulären Gruppen aus der Map hinzufügen
            foreach (var kvp in contactGroupsDict)
            {
                // Starred lassen wir hier bewusst aus, da es unten fix hinzugefügt wird
                if (!kvp.Value.Equals("starred", StringComparison.OrdinalIgnoreCase))
                {
                    allContactMemberships.Add(kvp.Value);
                }
            }
            // Und dann exakt einmal den Stern als UI-Label hinzufügen
            allContactMemberships.Add("★");

            toolStripProgressBar.Value = 50;

            // F. Datenbindung mit LOCK
            isSelectionChanging = true;
            try
            {
                _lastActiveContact = null;
                _originalContactSnapshot = null;

                var contactList = new BindingList<Contact>([.. result.Contacts]);

                if (contactList.Count == 0)
                {
                    toolStripStatusLabel.Text = "Keine Kontakte gefunden.";
                    contactDGV.DataSource = null;
                    return;
                }

                _allGoogleContacts = contactList;
                toolStripStatusLabel.Text = $"{contactList.Count} Kontakte";

                contactBSource.DataSource = contactList;
                contactDGV.DataSource = contactBSource;

                ApplyColumnSettings(contactDGV);
                toolStripProgressBar.Value = 80;

                SwitchDataBinding(contactBSource);

                tabControl.SelectedIndex = 1;
                Text = $"Kontakte - Google Kontakte";
            }
            finally { isSelectionChanging = false; }

            // G. UI Finalisierung
            var hasRows = contactDGV.Rows.Count > 0;
            //copyTSButton.Enabled = copyToOtherDGVTSMenuItem.Enabled = wordToolStripMenuItem.Enabled = googlebackupToolStripMenuItem.Enabled =
            //    envelopeToolStripMenuItem.Enabled = clipboardTSButton.Enabled = wordTSButton.Enabled = envelopeTSButton.Enabled = hasRows;
            SetCommonButtonState(hasRows);
            googlebackupToolStripMenuItem.Enabled = hasRows;
            duplicateToolStripMenuItem.Enabled = false;
            btnEditContact.Visible = true;
            if (hasRows)
            {
                contactDGV.ClearSelection();  // Grid optisch an den absoluten Anfang setzen
                contactDGV.FirstDisplayedScrollingRowIndex = 0;
                contactDGV.Rows[0].Selected = true;
                contactBSource.Position = 0;  // Zur Sicherheit auch die BindingSource-Position hart auf 0 nageln
                ContactBindingSource_CurrentChanged(contactBSource, EventArgs.Empty);  // dadurch wird _lastActiveContact aktualisiert, Foto geladen, etc.
                contactBSource.ResetBindings(false);  // Der WinForms-Hammer: Zwingt alle Editier-Textboxen rechts zum sofortigen Refresh!
            }
            else { ContactBindingSource_CurrentChanged(contactBSource, EventArgs.Empty); }  // Auch bei einer leeren Liste müssen wir aufräumen (Felder leeren)
            if (tabulation.TabPages.Contains(tabPageDoku))  // Doku-Tab wegräumen falls nötig
            {
                deactivatedPage = tabPageDoku;
                tabulation.TabPages.Remove(tabPageDoku);
            }
            if (contactBirthdayFlag && _settings.BirthdayContactShow)  // Geburtstagserinnerung
            {
                toolStripProgressBar.Visible = false;
                _ = InvokeAsync(() => BirthdayReminder(contactDGV));  // sorgt dafür, dass dieser Code erst ausgeführt wird, wenn die aktuelle Methode (inkl. finally!) beendet ist
            }
            contactBirthdayFlag = true;
            toolStripProgressBar.Value = 100;
            Utils.StartSearchCacheWarmup(_allGoogleContacts);  // Background Warmup
            UpdateTagComboBoxDataSource();
            UpdatePlaceholderVis();
        }
        catch (UnauthorizedAccessException)
        {
            contactBirthdayFlag = false;
            Utils.MsgTaskDlg(Handle, "Autorisierung erforderlich",
                "Das Zugriffstoken ist abgelaufen. Bitte im Browser erneut anmelden.",
                TaskDialogIcon.Information);
        }
        catch (Exception ex) when (!IsDisposed) { Utils.ErrTaskDlg(Handle, ex); }
        finally
        {
            _isFiltering = false;
            await Task.Delay(400);
            if (!IsDisposed && toolStripProgressBar != null)
            {
                toolStripProgressBar.Visible = false;
                toolStripStatusLabel.Visible = true;
            }
        }
    }

    private async Task DeleteGoogleContactAsync(Contact contact, CancellationToken token)
    {
        if (contact == null) { return; }   // 1. Nur auf null prüfen
        if (string.IsNullOrEmpty(contact.ResourceName)) { return; }  // 2. Wenn keine ResourceName da ist (Kontakt war nie bei Google), ist nichts zu tun
        var manager = new GooglePeopleManager(secretPath, tokenDir);
        await manager.DeleteContactAsync(contact.ResourceName, token);
    }

    private async Task UpdateContactPhotoAsync(Contact contact, Image imageToUpload, ImageFormat formatToUse, Action onClose)
    {
        try
        {
            var manager = new GooglePeopleManager(secretPath, tokenDir);
            var newUrl = await manager.UpdateContactPhotoAsync(contact.ResourceName, imageToUpload, formatToUse);

            if (!string.IsNullOrEmpty(newUrl))
            {
                contact.PhotoUrl = newUrl;
                contact.ResetSearchCache();

                var index = contactBSource.IndexOf(contact);
                if (index >= 0) { contactBSource.ResetItem(index); }
            }
        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        finally { onClose?.Invoke(); }
    }

    private async Task DeleteContactPhotoAsync(Contact contact)
    {
        if (contact == null || string.IsNullOrEmpty(contact.ResourceName)) { return; }

        try
        {
            var manager = new GooglePeopleManager(secretPath, tokenDir);
            var newUrl = await manager.DeleteContactPhotoAsync(contact.ResourceName);

            contact.PhotoUrl = newUrl; // Ist null oder Platzhalter
            contact.ResetSearchCache();
            _ = ShowPhotoInPictureBox(contact, CancellationToken.None);
        }
        catch (Exception ex)
        {
            if (ex.Message.Contains("NotFound")) // Einfacher Check statt using Google...
            {
                Utils.MsgTaskDlg(Handle, "Kein Foto", "Es konnte online kein Foto gefunden werden.", TaskDialogIcon.Information);
                contact.PhotoUrl = null;
                _ = ShowPhotoInPictureBox(contact, CancellationToken.None);
            }
            else
            {
                Utils.ErrTaskDlg(Handle, ex);
            }
        }
    }

    private async void GoogleTSButton_Click(object sender, EventArgs e) => await CheckContactChanges(LoadAndDisplayGoogleContactsAsync);

    private async void ContactDGV_CellClick(object sender, DataGridViewCellEventArgs e)
    {
        // 1. Validitätsprüfung (keine Header, keine ungültigen Klicks)
        if (e.RowIndex < 0 || e.ColumnIndex < 0) { return; }

        // 2. Prüfung auf Strg-Taste via WinForms ModifierKeys
        if ((ModifierKeys & Keys.Control) == Keys.Control)
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
                if (targetControl is TextBoxBase tb) { tb.SelectAll(); }
                else if (targetControl is ComboBox cb) { cb.DroppedDown = true; }
            }
        }
    }

    private async void ContactBindingSource_CurrentChanged(object? sender, EventArgs e)
    {
        if (_isFiltering || isSelectionChanging) { return; }
        ErzeugeGrussformeln(contactBSource.Current);  // Synchrone Ausführung VOR dem Delay
        if (_scrollCancellation != null)
        {
            _scrollCancellation.Cancel();
            _scrollCancellation.Dispose();
        }
        _scrollCancellation = new CancellationTokenSource();
        var token = _scrollCancellation.Token;
        try
        {
            await Task.Delay(100, token);
            if (!token.IsCancellationRequested) { UpdateContactDetails(token); }
        }
        catch (TaskCanceledException) { }
    }

    private void UpdateContactDetails(CancellationToken token)
    {
        if (contactBSource.Current is not Contact contact)
        {
            _originalContactSnapshot = null;
            _lastActiveContact = null;
            topAlignZoomPictureBox.Image = Resources.ContactBild100;
            delPictboxToolStripButton.Enabled = false;
            AgeLabel_MaskedTB_Clear();
            flowLayoutPanel.Controls.Clear();
            btnEditContact.Visible = false;
            saveTSButton.Enabled = false;
            curContactMemberships.Clear();
            UpdateTagComboBoxDataSource();
            return;
        }

        ignoreTextChange = true;

        try
        {
            _lastActiveContact = contact;
            _originalContactSnapshot = (Contact)contact.Clone();
            _ = ShowPhotoInPictureBox(contact, token);
            if (contact.Geburtstag.HasValue) { AgeLabel_MaskedTB_Set(contact.Geburtstag.Value); }
            else { AgeLabel_MaskedTB_Clear(); }
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

            UpdateTagComboBoxDataSource();
            LinkLabel_Enabled();
            btnEditContact.Visible = true;
            saveTSButton.Enabled = false;
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
        // 1. REKURSIONS-SCHUTZ:
        // Wenn wir den Wechsel selbst im Code ausgelöst haben (siehe unten), 
        // dann Prüfungen überspringen und durchwinken.
        if (_isTabSwitchingProgrammatically) { return; }

        // ------------------------------------------------------------------------
        // FALL A: WIR VERLASSEN DEN GOOGLE-TAB
        // (Wir sind aktuell auf 'contactTabPage' und wollen woanders hin)
        // ------------------------------------------------------------------------
        if (_previousTab == contactTabPage && e.TabPage != contactTabPage) // tabControl.SelectedTab == contactTabPage ist unsicher, weil der Tab-Wechsel schon im Gange ist
        {
            //Console.Beep();
            // 1. Wechsel VORSORGLICH abbrechen.
            // Warum? Windows Forms wartet nicht auf 'await'. Ohne das Cancel würde der 
            // Tab sofort wechseln, während der Speicher-Dialog noch lädt.
            e.Cancel = true;

            // 2. Den "Gatekeeper" fragen
            // Diese Methode kümmert sich um alles: Validierung, Frage an User, 
            // Speichern (Progressbar), Verwerfen oder Aufräumen leerer neuer Kontakte.
            var readyToLeave = await ContactChanges_Check();

            // 3. Entscheidung auswerten
            if (readyToLeave)
            {
                var targetTab = e.TabPage;

                // MODERN: InvokeAsync statt BeginInvoke
                // Durch das 'await' kehrt diese Methode hier zum TabControl zurück,
                // der Abbruch (e.Cancel) wird wirksam, und DANACH läuft der Code im Block.
                await InvokeAsync(() =>
                {
                    _isTabSwitchingProgrammatically = true;
                    try
                    {
                        tabControl.SelectedTab = targetTab;
                    }
                    finally
                    {
                        _isTabSwitchingProgrammatically = false;
                    }

                    // Filter zurücksetzen
                    if (filterRemoveToolStripMenuItem.Visible)
                    {
                        FilterRemoveToolStripMenuItem_Click(null!, null!);
                    }
                });
            }
            // Wenn readyToLeave == false (User hat "Abbrechen" im Dialog geklickt), 
            // bleibt e.Cancel = true und wir bleiben auf dem Google-Tab.
            return;
        }

        // ------------------------------------------------------------------------
        // FALL B: WIR BETRETEN DEN GOOGLE-TAB (Laden der Daten)
        // ------------------------------------------------------------------------
        if (e.TabPage == contactTabPage)
        {
            // Prüfen, ob geladen werden muss
            if (contactBSource.DataSource == null || contactBSource.Count == 0)
            {
                // Hinweis: Um "async void" Probleme zu minimieren, lagern wir das Laden oft aus.
                // Hier ist es okay, aber der Dialog blockiert kurz den Tab-Wechsel visuell.
                var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Google Kontakte", "Keine Kontakte vorhanden", "Möchtest du die Kontakte jetzt laden?");

                if (isYes) { await LoadAndDisplayGoogleContactsAsync(); }
            }
        }
    }

    private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
    {
        _previousTab = tabControl.SelectedTab;
        // ========================================================================
        // TAB: ADRESSEN (SQL)
        // ========================================================================
        if (tabControl.SelectedTab == addressTabPage)
        {
            // Snapshot-Cleanup: Da wir jetzt im SQL-Tab sind, gibt es keinen "aktiven" Google-Kontakt
            _originalContactSnapshot = null;
            _lastActiveContact = null;

            if (deactivatedPage != null && !tabulation.TabPages.Contains(deactivatedPage))
            {
                tabulation.TabPages.Insert(1, deactivatedPage);
                deactivatedPage = null;
            }

            // Suche sichern/wiederherstellen
            HandleSearchTransition(ref lastContactSearch, ref lastAddressSearch);

            // Binding umschalten
            SwitchDataBinding(addressBSource);

            if (addressBSource.Current != null)
            {
                _ = ShowPhotoInPictureBox(addressBSource.Current, CancellationToken.None);
            }

            // UI Status
            if (addressBSource?.Count > 0)
            {
                Text = $"{appName} – {(string.IsNullOrEmpty(_databaseFilePath) ? "unbenannt" : _databaseFilePath)}";
                btnEditContact.Visible = false;
                UpdateSaveButton();

                // Buttons aktivieren
                SetCommonButtonState(true);
                copyToOtherDGVTSMenuItem.Enabled = false;

                // Statuszeile
                var rowCount = _context?.Adressen.Local.Count ?? 0;
                var visibleRowCount = addressBSource.Count;
                toolStripStatusLabel.Text = rowCount == visibleRowCount
                    ? $"{visibleRowCount} Adressen"
                    : $"{visibleRowCount}/{rowCount} Adressen";
            }
        }

        // ========================================================================
        // TAB: GOOGLE KONTAKTE
        // ========================================================================
        else if (tabControl.SelectedTab == contactTabPage)
        {
            // Snapshot Logik initialisieren (Wichtig für den Gatekeeper beim nächsten Wechsel)
            if (contactBSource.Current is Contact current)
            {
                _lastActiveContact = current;
                _originalContactSnapshot = (Contact)current.Clone();
            }

            // Tabulation (Doku Tab entfernen)
            if (tabulation.TabPages.Contains(tabPageDoku))
            {
                deactivatedPage = tabPageDoku;
                tabulation.TabPages.Remove(tabPageDoku);
            }

            // Suche sichern/wiederherstellen
            HandleSearchTransition(ref lastAddressSearch, ref lastContactSearch);

            // Binding umschalten
            if (contactBSource.DataSource != null) { SwitchDataBinding(contactBSource); }
            // UI Status
            if (contactBSource.Count > 0)
            {
                //Text = !string.IsNullOrWhiteSpace(userEmail) ? $"Kontakte - {userEmail}" : "Google-Kontakte";
                Text = "Kontakte - Google Kontakte";
                btnEditContact.Visible = true;

                // Menü Items gemäß Logik (Google-Tab hat andere Regeln für Neu/Löschen im Menü)
                newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = false;

                // Toolbar Buttons aktivieren
                SetCommonButtonState(true);
                copyToOtherDGVTSMenuItem.Enabled = true;

                toolStripStatusLabel.Text = $"{contactBSource.Count} Kontakte";
            }
        }

        UpdateMembershipTags();
        flexiTSStatusLabel.Text = string.Empty;
        searchTSTextBox.TextBox.Focus();
        UpdateFilterUIState();
    }

    private void SetCommonButtonState(bool enabled)
    {
        newTSButton.Enabled = duplicateToolStripMenuItem.Enabled =
        deleteToolStripMenuItem.Enabled = deleteTSButton.Enabled =
        copyTSButton.Enabled = clipboardTSButton.Enabled =
        wordTSButton.Enabled = envelopeTSButton.Enabled = enabled;
    }

    private void HandleSearchTransition(ref string sourceStorage, ref string targetStorage)
    {
        if (searchTSTextBox.TextBox.TextLength > 0)
        {
            sourceStorage = searchTSTextBox.Text;
            SetSearchTextIgnoreChange(string.Empty);
        }

        if (!string.IsNullOrEmpty(targetStorage))
        {
            SetSearchTextIgnoreChange(targetStorage);
            targetStorage = string.Empty;
        }
    }

    private void SetSearchTextIgnoreChange(string newText)
    {
        ignoreSearchChange = true;
        try
        {
            searchTSTextBox.Text = newText;   // Einheitlich .Text verwenden (greift ohnehin auf die TextBox durch) 
            tsClearLabel.Visible = !string.IsNullOrWhiteSpace(newText); // Den X-Button im Suchfeld direkt passend mitsynchronisieren
        }
        finally { ignoreSearchChange = false; }
    }

    private void AuthentMenuItem_Click(object sender, EventArgs e)
    {
        using TaskDialogIcon questionDialogIcon = new(Resources.question32);
        TaskDialogPage page = new()
        {
            Caption = appCont,
            Heading = "Möchtest du die Zugangsdaten löschen?",
            Text = "Wenn du den Request-Token löschst, können Google-\nKontakte nur nach erneuter Autorisierung herunter-\ngeladen werden. Hierzu öffnet sich beim nächsten Versuch automatisch die Google-Anmeldeseite.",
            Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
            Icon = questionDialogIcon,
            DefaultButton = TaskDialogButton.No
        };
        if (TaskDialog.ShowDialog(this, page) == TaskDialogButton.Yes)
        {
            var tokenFile = Path.Combine(tokenDir, "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            try { if (File.Exists(tokenFile)) { File.Delete(tokenFile); } }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            finally { GooglePeopleManager.ClearServiceCache(); }
        }
    }

    private void ExtraToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        authentMenuItem.Enabled = Directory.Exists(tokenDir);
        manageGroupsToolStripMenuItem.Enabled = tabControl.SelectedTab == contactTabPage ? contactDGV.Rows.Count > 0 : addressBSource != null;
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
        Cursor = Cursors.WaitCursor;
        FillDictionary();
        using var frm = new FrmPrintSetting(_settings, bookmarkTextDictionary);
        Utils.RestoreWindowBounds(frm, _settings.PrintWindowPosition);
        Cursor = Cursors.Default;
        if (frm.ShowDialog() == DialogResult.OK)
        {
            var bounds = frm.WindowState == FormWindowState.Normal ? frm.DesktopBounds : frm.RestoreBounds;
            _settings.PrintWindowPosition = new WindowPlacement
            {
                X = bounds.X,
                Y = bounds.Y,
                Width = bounds.Width,
                Height = bounds.Height
            };
            SettingsManager.Save(_settings, _settingsPath);  // Optional: Sofortiges Speichern der JSON-Datei
        }
    }

    private void OptionsToolStripMenuItem_Click(object sender, EventArgs e)
    {
        // 1. Wir erstellen einen Klon der aktuellen Einstellungen.
        // Das Original (_settings) bleibt völlig unberührt, egal was der User im Dialog macht.
        var tempSettings = _settings.DeepClone();

        // 2. Wir übergeben den Klon an das Formular.
        // Das DataBinding arbeitet jetzt "live" auf 'tempSettings'.
        using var frm = new FrmProgSettings(tempSettings);

        if (frm.ShowDialog(this) == DialogResult.OK)
        {
            // 3. Nur bei OK: Wir tauschen das Original gegen den bearbeiteten Klon aus.
            _settings = tempSettings;

            // UI & System-Trigger auf Basis der neuen Werte ausführen
            ApplyEditControlsFont();
            SetColorScheme();
            ApplyFileWatcherSettings();

            // Einstellungen dauerhaft speichern
            SaveConfiguration();
        }
        // Bei "Abbrechen" passiert gar nichts. 
        // 'tempSettings' wird verworfen und _settings bleibt, wie es war.
    }

    private void ApplyFileWatcherSettings()
    {
        var docPath = _settings.DocumentFolder;

        // Basiskonfiguration
        fileSysWatcher.IncludeSubdirectories = true;
        fileSysWatcher.Filters.Clear();

        // Beide Arrays (Dokumente und Bilder) kombinieren und die Filter setzen
        var allTypes = documentTypes.Concat(imageTypes);
        foreach (var pattern in allTypes) { fileSysWatcher.Filters.Add(pattern); }

        // Pfad setzen und nur aktivieren, wenn alles passt
        if (_settings.WatchFolder && !string.IsNullOrEmpty(docPath) && Directory.Exists(docPath))
        {
            fileSysWatcher.Path = docPath;
            fileSysWatcher.EnableRaisingEvents = true;
            min2TrayTSButton.Visible = true;
        }
        else
        {
            fileSysWatcher.EnableRaisingEvents = false;
            min2TrayTSButton.Visible = false;
        }
    }

    private void SetColorScheme()
    {
        switch (_settings.ColorScheme)
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

    private void ApplyEditControlsFont()
    {
        var fontName = _settings.AppFontName;
        var fontSize = _settings.AppFontSize;

        if (string.IsNullOrWhiteSpace(fontName)) { fontName = "Segoe UI"; }
        if (fontSize < 9f || fontSize > 12f) { fontSize = 10f; }

        _monospaceFont?.Dispose();
        _standardAppFont?.Dispose();
        _monospaceFont = Utils.GetSafeFont("Courier New", fontSize, FontFamily.GenericMonospace);
        _standardAppFont = Utils.GetSafeFont(fontName, fontSize, FontFamily.GenericSansSerif);

        try
        {
            addressDGV.DefaultCellStyle.Font = _standardAppFont;
            contactDGV.DefaultCellStyle.Font = _standardAppFont;

            // 2. Zeilenhöhe anhand der gewählten Schriftart exakt EINMAL berechnen
            // _standardAppFont.GetHeight() gibt die Höhe der Schrift in Pixeln zurück. 
            // Die +8 dienen als komfortabler Abstand (Padding) oben und unten.
            var calculatedRowHeight = (int)_standardAppFont.GetHeight() + 8;

            // 3. Diese Höhe als Standard für alle NEUEN Zeilen festlegen
            addressDGV.RowTemplate.Height = calculatedRowHeight;
            contactDGV.RowTemplate.Height = calculatedRowHeight;

            // 4. Die Höhe auch auf alle BEREITS GELADENEN Zeilen anwenden
            foreach (DataGridViewRow row in addressDGV.Rows) { row.Height = calculatedRowHeight; }
            foreach (DataGridViewRow row in contactDGV.Rows) { row.Height = calculatedRowHeight; }

            foreach (var control in tableLayoutPanel.Controls.OfType<Control>())
            {
                if (control is PaddedTextBox || control is PaddedMaskedTextBox) { control.Font = _standardAppFont; }
            }

            tbNotizen.Font = _standardAppFont;
            maskedTextBox.Font = _standardAppFont;
        }
        catch (Exception ex) { Debug.WriteLine($"Fehler beim Setzen der Schriftart: {ex.Message}"); }
    }

    private void BtnEditContact_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == contactTabPage && contactBSource.Current is Contact contact)
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

    private async void TsClearLabel_Click(object sender, EventArgs e) => await CheckContactChanges(Clear_SearchTextBox);

    private void TsClearLabel_VisibleChanged(object sender, EventArgs e) => UpdateSearchBoxWidth();  // searchTSTextBox.Width = 202 + splitContainer.SplitterDistance - 536 - (tsClearLabel.Visible ? tsClearLabel.Width : 0);

    private void ToolStrip_Paint(object? sender, PaintEventArgs e)
    {
        if (tsClearLabel is { Visible: true })   // C# 14 Property Pattern: Prüft auf null UND Sichtbarkeit in einem Rutsch
        {
            var rect = new Rectangle(
                tsClearLabel.Bounds.Location.X - 2,
                tsClearLabel.Bounds.Location.Y + 2,
                tsClearLabel.Width + 1,
                tsClearLabel.Height - 4);
            e.Graphics.DrawRectangle(Pens.Black, rect);
        }
    }

    private void AddressDGV_KeyDown(object sender, KeyEventArgs e) => HandleDGVKeyDown(e);

    private void AddressDGV_KeyUp(object sender, KeyEventArgs e) => HandleDGVKeyUp(e, () => AddressBindingSource_CurrentChanged(addressBSource, EventArgs.Empty));

    private void ContactDGV_KeyDown(object sender, KeyEventArgs e) => HandleDGVKeyDown(e);

    private void ContactDGV_KeyUp(object sender, KeyEventArgs e) => HandleDGVKeyUp(e, () => ContactBindingSource_CurrentChanged(contactBSource, EventArgs.Empty));

    private void HandleDGVKeyDown(KeyEventArgs e)
    {
        // 1. Panel einfrieren beim Scrollen
        if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
        {
            if (!_isPanelFrozen)
            {
                NativeMethods.SendMessage(splitContainer.Panel2.Handle, NativeMethods.WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
                _isPanelFrozen = true;
            }
            return; // Direkt raus, hier muss nichts mehr geprüft werden
        }

        // 2. Kopieren abfangen
        if (e.Control && e.KeyCode == Keys.C)
        {
            ClipboardTSMenuItem_Click(null!, null!);
            e.Handled = true;
            e.SuppressKeyPress = true;
            return;
        }

        // 3. Suche ansteuern (Buchstaben & Zahlen)
        var isLetter = e.KeyCode >= Keys.A && e.KeyCode <= Keys.Z;
        var isNumber = e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9;

        // Erlaubt Buchstaben (mit/ohne Shift) und Zahlen (ohne Shift)
        if ((e.Modifiers == Keys.None || e.Modifiers == Keys.Shift) && (isLetter || (isNumber && e.Modifiers == Keys.None)))
        {
            searchTSTextBox.Focus();

            if (isLetter)
            {
                var c = (char)e.KeyValue;
                searchTSTextBox.Text += e.Shift ? c.ToString() : char.ToLower(c).ToString();
            }
            else if (isNumber)
            {
                // Zahlen dürfen nicht +32 gerechnet werden
                searchTSTextBox.Text += ((char)e.KeyValue).ToString();
            }

            searchTSTextBox.SelectionStart = searchTSTextBox.Text.Length;
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
    }

    private void HandleDGVKeyUp(KeyEventArgs e, Action loadDetailsAction)
    {
        if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
        {
            if (_isPanelFrozen)
            {
                NativeMethods.SendMessage(splitContainer.Panel2.Handle, NativeMethods.WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
                splitContainer.Panel2.Invalidate(true);
                _isPanelFrozen = false;

                // Die übergebene Methode für den aktuellen Datensatz ausführen
                loadDetailsAction?.Invoke();
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

    private void TextBox_Enter(object sender, EventArgs e)
    {
        if (sender is TextBox tb)
        {
            tb.SelectAll();
            // Dark Mode: Dunkles Gold/Gelb | Light Mode: LightYellow
            tb.BackColor = _isDarkMode ? Color.FromArgb(80, 80, 0) : Color.LightYellow;
            tb.ForeColor = _isDarkMode ? Color.White : Color.Black;
        }
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

        maskedTextBox.BackColor = _isDarkMode ? Color.FromArgb(80, 80, 0) : Color.LightYellow;
        maskedTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;

        if (maskedTextBox.Focused)
        {
            maskedTextBox.Font = _monospaceFont; // Cached Font nutzen
            var cleanText = maskedTextBox.Text.Replace(".", "").Replace("_", "").Replace(" ", "").Trim();

            // Falls leer, Cursor an den Anfang setzen, ansonsten alles markieren
            if (string.IsNullOrWhiteSpace(cleanText))
            {
                maskedTextBox.SelectionStart = 0;
                maskedTextBox.SelectionLength = 0;
            }
            else
            {
                maskedTextBox.SelectAll();
            }
        }

        ignoreTextChange = false;
    }

    private void MaskedTextBox_KeyPress(object sender, KeyPressEventArgs e)
    {
        // Ermöglicht das flüssige Springen zum nächsten Block per Punkt oder Komma,
        // ohne dass die Auto-Formatierung dazwischenfunkt.
        if (e.KeyChar == '.' || e.KeyChar == ',')
        {
            e.Handled = true;

            if (maskedTextBox.SelectionStart <= 2)
            {
                maskedTextBox.SelectionStart = 3;
            }
            else if (maskedTextBox.SelectionStart <= 5)
            {
                maskedTextBox.SelectionStart = 6;
            }
        }
    }

    private void MaskedTextBox_TextChanged(object sender, EventArgs e)
    {
        if (!maskedTextBox.Focused || ignoreTextChange) { return; }
        maskedTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
        if (!maskedTextBox.MaskFull)
        {
            var cleanText = maskedTextBox.Text.Replace(".", "").Replace("_", "").Replace(" ", "").Trim();
            if (string.IsNullOrWhiteSpace(cleanText)) { AgeLabel_MaskedTB_Clear(); }
        }
        else
        {
            var rawText = maskedTextBox.Text;
            if (DateOnly.TryParseExact(rawText, formats, culture, DateTimeStyles.None, out var geburtsdatum))
            {
                if (geburtsdatum > DateOnly.FromDateTime(DateTime.Today))
                {
                    maskedTextBox.ForeColor = Color.Red;
                    AgeLabel_MaskedTB_Clear();
                }
                else
                {
                    maskedTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
                    AgeLabel_MaskedTB_Set(geburtsdatum); // Nutzt deine bestehende Methode für Label & Tooltip
                }
            }
            else
            {
                maskedTextBox.ForeColor = Color.Red;
                AgeLabel_MaskedTB_Clear();
            }
        }
        UpdateSaveButton();
    }

    private void MaskedTextBox_Leave(object sender, EventArgs e)
    {
        ignoreTextChange = true;
        maskedTextBox.BackColor = _isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
        maskedTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
        try
        {
            maskedTextBox.Font = _standardAppFont; // Cached Font nutzen
            // Prüfen, ob der Nutzer überhaupt Zahlen eingetippt hat
            if (maskedTextBox.MaskedTextProvider != null && maskedTextBox.MaskedTextProvider.AssignedEditPositionCount > 0) { FormatAndSetDate(); }

            // Abschließende Auswertung und farbliche Markierung (falls FormatAndSetDate es nicht fixen konnte)
            if (DateOnly.TryParseExact(maskedTextBox.Text, formats, culture, DateTimeStyles.None, out var geburtsdatum))
            {
                if (geburtsdatum > DateOnly.FromDateTime(DateTime.Today))
                {
                    maskedTextBox.ForeColor = Color.Red;
                    AgeLabel_MaskedTB_Clear();
                }
                else { AgeLabel_MaskedTB_Set(geburtsdatum); }
            }
            else
            {
                AgeLabel_MaskedTB_Clear();

                // Wenn wirklich Text da ist, der aber kein gültiges Datum ergibt, markieren wir ihn rot.
                if (maskedTextBox.MaskedTextProvider != null && maskedTextBox.MaskedTextProvider.AssignedEditPositionCount > 0) { maskedTextBox.ForeColor = Color.Red; }
            }
        }
        finally { ignoreTextChange = false; }
    }

    private void FormatAndSetDate()
    {
        var originalFormat = maskedTextBox.TextMaskFormat;

        // Maske inkl. Platzhalter auslesen, um korrekte Positionen zu garantieren (z. B. " 2. 5.64  ")
        maskedTextBox.TextMaskFormat = MaskFormat.IncludePromptAndLiterals;
        var fullText = maskedTextBox.Text;
        maskedTextBox.TextMaskFormat = originalFormat;

        if (string.IsNullOrWhiteSpace(fullText) || fullText.Length < 10) { return; }

        var prompt = maskedTextBox.PromptChar.ToString();
        var dStr = fullText[..2].Replace(prompt, " ").Trim();
        var mStr = fullText.Substring(3, 2).Replace(prompt, " ").Trim();
        var yStr = fullText.Substring(6, 4).Replace(prompt, " ").Trim();

        // Wenn Tag oder Monat fehlen, brechen wir ab (wird dann durch das Leave-Event als fehlerhaft markiert)
        if (string.IsNullOrEmpty(dStr) || string.IsNullOrEmpty(mStr)) { return; }

        // Tag und Monat ggf. mit führenden Nullen auffüllen
        dStr = dStr.PadLeft(2, '0');
        mStr = mStr.PadLeft(2, '0');

        var currentYear = DateTime.Today.Year;

        if (string.IsNullOrEmpty(yStr))
        {
            // Bei Geburtsdaten nehmen wir nicht automatisch das aktuelle Jahr an.
            // Wenn das Jahr komplett fehlt, lassen wir es unvollständig -> Validation schlägt fehl.
            return;
        }
        else if (yStr.Length <= 2)
        {
            if (int.TryParse(yStr, out var shortYear))
            {
                var assumedYear = 2000 + shortYear;

                if (assumedYear > currentYear) { assumedYear = 1900 + shortYear; }

                yStr = assumedYear.ToString();
            }
        }
        else if (yStr.Length == 3) { yStr = yStr.PadLeft(4, '0'); }

        var dateString = $"{dStr}.{mStr}.{yStr}";

        ignoreTextChange = true;
        try
        {
            if (DateTime.TryParseExact(dateString, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            {
                maskedTextBox.Text = dateString;
                maskedTextBox.DataBindings["Text"]?.WriteValue();
                UpdateSaveButton(); // Da wir die Text-Property manuell setzen, müssen wir Save evtl. triggern
            }
        }
        finally { ignoreTextChange = false; }
    }

    private void MaskedTextBox_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            // Wenn das Feld noch völlig leer ist, setzen wir den Cursor strikt nach ganz links.
            if (maskedTextBox.MaskedTextProvider == null || maskedTextBox.MaskedTextProvider.AssignedEditPositionCount == 0)
            {
                maskedTextBox.SelectionStart = 0;
                maskedTextBox.SelectionLength = 0;
                return;
            }

            var charIndex = maskedTextBox.GetCharIndexFromPosition(e.Location);

            // Die Maske ist fix: "00.00.0000" (10 Zeichen)
            switch (charIndex)
            {
                case <= 2: // Tag (Indizes 0, 1 und Trennzeichen 2)
                    {
                        maskedTextBox.SelectionStart = 0;
                        maskedTextBox.SelectionLength = 2;
                        break;
                    }
                case >= 3 and <= 5: // Monat (Indizes 3, 4 und Trennzeichen 5)
                    {
                        maskedTextBox.SelectionStart = 3;
                        maskedTextBox.SelectionLength = 2;
                        break;
                    }
                case >= 6: // Jahr (Indizes 6 bis 9)
                    {
                        maskedTextBox.SelectionStart = 6;
                        maskedTextBox.SelectionLength = 4;
                        break;
                    }
            }
        }
    }

    private void BtnResetDate_Click(object sender, EventArgs e)
    {
        ignoreTextChange = true;
        try
        {
            maskedTextBox.Clear();
            AgeLabel_MaskedTB_Clear(); // Alter-Label und Tooltip ebenfalls leeren

            // Das WinForms-Binding scheitert oft daran, eine leere Maske in DateOnly? umzuwandeln.
            // Daher setzen wir den Wert direkt im zugrundeliegenden Objekt auf null.
            if (tabControl.SelectedTab == addressTabPage)
            {
                if (addressBSource?.Current is Adresse adresse) { adresse.Geburtstag = null; }
            }
            else
            {
                if (contactBSource?.Current is Contact contact) { contact.Geburtstag = null; }
            }

            // Weist das Binding an, den jetzt genullten Wert neu aus dem Modell in die TextBox zu laden
            maskedTextBox.DataBindings["Text"]?.ReadValue();
        }
        finally { ignoreTextChange = false; }

        maskedTextBox.Focus();
        UpdateSaveButton();
    }

    private void TextBox_TextChanged(object sender, EventArgs e)
    {
        if (sender is not Control senderControl || !senderControl.Focused || ignoreTextChange || _isFiltering) { return; }
        var isLocal = tabControl.SelectedTab == addressTabPage;
        var isGoogle = tabControl.SelectedTab == contactTabPage;
        if (!isLocal && !isGoogle) { return; }
        if (isLocal) { senderControl.DataBindings["Text"]?.WriteValue(); } // Zwinge das Binding, den Wert SOFORT in das Entity zu schreiben
        if (isGoogle && contactBSource.Current is not Contact) { return; }
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
            //saveFileDialog.InitialDirectory = string.IsNullOrEmpty(sDatabaseFolder) || !Directory.Exists(sDatabaseFolder) ? null : sDatabaseFolder;
            saveFileDialog.InitialDirectory = string.IsNullOrEmpty(_settings.DatabaseFolder) || !Directory.Exists(_settings.DatabaseFolder) ? null : _settings.DatabaseFolder;
            saveFileDialog.DefaultExt = "adb";
            saveFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (addressBSource != null) { await SaveSQLDatabaseAsync(true); }
                _databaseFilePath = saveFileDialog.FileName;
            }
            else { return; }
            CreateNewDatabase(_databaseFilePath, true);
            //ConnectSQLDatabase(_databaseFilePath);
            await ConnectSQLDatabaseAsync(_databaseFilePath);
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            _databaseFilePath = string.Empty;
        }
    }

    private async void ExportToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (addressBSource.Count == 0)
        {
            Utils.MsgTaskDlg(Handle, "Export nicht möglich", "Es gibt keine Datensätze zum Exportieren.");
            return;
        }
        saveFileDialog.FileName = "Adressen_Export.csv";
        saveFileDialog.DefaultExt = "csv";
        saveFileDialog.Filter = "CSV-Datei (*.csv)|*.csv|Alle Dateien (*.*)|*.*";

        if (saveFileDialog.ShowDialog() != DialogResult.OK) { return; }
        var fileName = saveFileDialog.FileName;
        // 1. Spalten vorbereiten (und sicherstellen, dass "Gruppen" auch wirklich dabei ist!)
        var exportColumns = dataFields.Where(f => f != "Id").ToList();
        if (!exportColumns.Contains("Gruppen")) { exportColumns.Add("Gruppen"); }

        // 2. Daten threadsicher in eine Liste kopieren (noch im UI-Thread)
        // Wir filtern direkt auf den Typ 'Adresse', um Fehler zu vermeiden
        var addressesToExport = addressBSource.List.OfType<Adresse>().ToList();

        // 3. UI für den Export vorbereiten
        toolStripProgressBar.Visible = true;
        toolStripProgressBar.Maximum = addressesToExport.Count;
        toolStripProgressBar.Value = 0;
        toolStripStatusLabel.Text = "Export läuft...";
        IProgress<int> progress = new Progress<int>(v => { toolStripProgressBar.Value = v; });
        try
        {
            // 4. Export asynchron im Hintergrund ausführen
            await Task.Run(async () =>
            {
                // Wir nutzen WriteLineAsync für noch bessere I/O-Performance
                using var writer = new StreamWriter(fileName, append: false, System.Text.Encoding.UTF8);
                await writer.WriteLineAsync(string.Join(";", exportColumns));
                var processedCount = 0;
                foreach (var adresse in addressesToExport)
                {
                    var fields = exportColumns.Select(columnName =>
                    {
                        var value = default(object);
                        if (columnName == "Gruppen") { value = string.Join(", ", adresse.Gruppen.Select(g => g.Name)); }
                        else if (columnName == "Geburtstag") { value = adresse.Geburtstag?.ToString("dd.MM.yyyy"); } // Festes, deutsches Format
                        else { value = adresse.GetPropertyValue(columnName); }

                        var fieldString = value?.ToString() ?? string.Empty;
                        return $"\"{fieldString.Replace("\"", "\"\"")}\"";
                    });
                    await writer.WriteLineAsync(string.Join(";", fields));
                    processedCount++;
                    if (processedCount % 50 == 0) { progress.Report(processedCount); }  // Fortschritt nur alle 50 Datensätze ans UI melden, um Flackern zu vermeiden
                }
                // Sicherstellen, dass der Ladebalken am Ende voll ist
                progress.Report(addressesToExport.Count);
            });

            toolStripStatusLabel.Text = "Export abgeschlossen.";
            Utils.MsgTaskDlg(Handle, "Export abgeschlossen", $"{addressesToExport.Count} Datensätze wurden erfolgreich exportiert.", TaskDialogIcon.ShieldSuccessGreenBar);
        }
        catch (Exception ex)
        {
            toolStripStatusLabel.Text = "Fehler beim Export.";
            Utils.ErrTaskDlg(Handle, ex);
        }
        finally { toolStripProgressBar.Visible = false; }
    }

    private void ColumnSelectToolStripMenuItem_Click(object sender, EventArgs e)
    {
        // 1. Wir übergeben den aktuellen Status und die Defaults direkt an das Formular
        using var frm = new FrmColumns(_settings.HideColumnArr, AppSettings.DefaultHideColumns);

        // 2. Auswertung bei OK
        if (frm.ShowDialog() == DialogResult.OK)
        {
            // 3. Formular liefert das fertige Array
            _settings.HideColumnArr = frm.GetNewVisibilityArray();

            ApplyColumnSettings(addressDGV);
            ApplyColumnSettings(contactDGV);
            SettingsManager.Save(_settings, _settingsPath);
        }
    }

    private void ColumnWidthsResetToolStripMenuItem_Click(object sender, EventArgs e)
    {
        // 1. Wir holen die Factory-Defaults aus der Klasse und überschreiben die aktuellen Einstellungen.
        // .Clone() ist extrem wichtig, damit wir eine neue Kopie erhalten und nicht das statische Original referenzieren.
        _settings.ColumnWidths = (int[])AppSettings.DefaultColumnWidths.Clone();

        // 2. Anwenden auf die beiden Grids
        // Wir nutzen einfach die Methode, die wir vorhin optimiert haben.
        ApplyColumnSettings(addressDGV);
        ApplyColumnSettings(contactDGV);

        // 3. Speichern
        SettingsManager.Save(_settings, _settingsPath);
    }

    private void SplitterAutomaticToolStripMenuItem_Click(object sender, EventArgs e) => splitContainer.SplitterDistance = toolStripSeparator.Bounds.Left;

    //private void SplitContainer_SplitterMoved(object sender, SplitterEventArgs e) => flexiTSStatusLabel.Width = 244 + splitContainer.SplitterDistance - 536;
    private void SplitContainer_SplitterMoved(object sender, SplitterEventArgs e)
    {
        UpdateStatusLabelWidth();
        UpdateSearchBoxWidth();
    }

    private void WordToolStripMenuItem_Click(object sender, EventArgs e) => WordTSButton_Click(sender, e);

    private void EnvelopeToolStripMenuItem_Click(object sender, EventArgs e) => EnvelopeTSButton_Click(sender, e);

    private void ClipboardTSMenuItem_Click(object sender, EventArgs e)
    {
        FillDictionary();

        // 1. Klon erstellen (für sauberes Abbrechen)
        var tempSettings = _settings.DeepClone();

        // 2. Form mit Settings-Objekt initialisieren
        // Hinweis: FrmCopyScheme muss angepasst werden (siehe unten)
        using var frm = new FrmCopyScheme(tempSettings, bookmarkTextDictionary);
        if (_settings.CopyWindowPosition != null) { frm.StartPosition = FormStartPosition.Manual; }
        Utils.RestoreWindowBounds(frm, _settings.CopyWindowPosition);
        if (frm.ShowDialog() == DialogResult.OK)
        {
            _settings = tempSettings;  // Die Konvertierung der Listen in Arrays ist bereits im Dialog passiert
            var bounds = frm.WindowState == FormWindowState.Normal ? frm.DesktopBounds : frm.RestoreBounds;
            _settings.CopyWindowPosition = new WindowPlacement
            {
                X = bounds.X,
                Y = bounds.Y,
                Width = bounds.Width,
                Height = bounds.Height
            };
            SettingsManager.Save(_settings, _settingsPath);
        }
    }

    private void ContextMenu_Opening(object sender, CancelEventArgs e)
    {
        // 1. Grundsätzliche Prüfung: Ist überhaupt etwas ausgewählt?
        var isAddressTab = tabControl.SelectedTab == addressTabPage;
        var isContactTab = tabControl.SelectedTab == contactTabPage;

        // Wir nutzen die BindingSource.Current statt SelectedRows, da dies robuster ist
        if ((isAddressTab && addressBSource.Current == null) ||
            (isContactTab && contactBSource.Current == null))
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
            newTSMenuItem.Text = "Adresse hinzufügen";
            dupTSMenuItem.Text = "Adresse duplizieren";
            delTSMenuItem.Text = "Adresse löschen";
            copy2OtherDGVMenuItem.Text = "Zu Google-Kontakte hinzufügen";
            // Nur anzeigen, wenn Google-Kontakte grundsätzlich geladen wurden
            copy2OtherDGVMenuItem.Visible = _allGoogleContacts?.Count > 0;
            //move2OtherDGVToolStripMenuItem.Visible = false;
        }
        else if (isContactTab)
        {
            if (contactDGV.CurrentRow != null && !Utils.RowIsVisible(contactDGV, contactDGV.CurrentRow))
            {
                contactDGV.FirstDisplayedScrollingRowIndex = contactDGV.CurrentRow.Index;
            }
            newTSMenuItem.Text = "Kontakt hinzufügen";
            dupTSMenuItem.Text = "Kontakt duplizieren";
            delTSMenuItem.Text = "Kontakt löschen";
            copy2OtherDGVMenuItem.Text = "In Lokale Adressen kopieren";
            // Immer möglich, sofern eine Datenbankverbindung besteht
            copy2OtherDGVMenuItem.Visible = _context != null;
            //move2OtherDGVToolStripMenuItem.Visible = _context != null;
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

    private void MainToolStripMenuItem_DropDownClosed(object sender, EventArgs e) => ((ToolStripMenuItem)sender).ForeColor = _settings.ColorScheme == "dark" ? SystemColors.HighlightText : SystemColors.ControlText;

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

    private async void RejectChangesToolStripMenuItem_Click(object sender, EventArgs e)
    {
        // 1. Google Kontakte Logik (Snapshot)
        if (tabControl.SelectedTab == contactTabPage && contactBSource.Current is Contact currentContact)
        {
            if (_originalContactSnapshot is null) { return; }
            foreach (var propName in editControlsDictionary.Values.Distinct())
            {
                var propInfo = typeof(Contact).GetProperty(propName);
                if (propInfo is not null && propInfo.CanWrite)
                {
                    propInfo.SetValue(currentContact, propInfo.GetValue(_originalContactSnapshot));
                }
            }
            currentContact.Geburtstag = _originalContactSnapshot.Geburtstag;
            currentContact.PhotoUrl = _originalContactSnapshot.PhotoUrl;
            currentContact.GroupNames.Clear();
            if (_originalContactSnapshot.GroupNames is not null) { currentContact.GroupNames.AddRange(_originalContactSnapshot.GroupNames); }
            currentContact.ResetSearchCache();

            contactBSource.ResetBindings(false);

            UpdateSaveButton();
            UpdateContactStatusBar();
        }
        // 2. Lokale EF Core Adressen
        else if (tabControl.SelectedTab == addressTabPage && _context is not null)
        {
            var analysis = DbChangeAnalyzer.AnalyzeChanges(_context);
            if (!analysis.HasChanges) { return; }

            var confirmHeading = "Möchtest du die Änderungen verwerfen?";

            var (isYes, _, _) = Utils.YesNo_TaskDialog(
                this,
                "Änderungen rückgängig machen",
                confirmHeading,
                analysis.DialogText,
                "Änderungen verwerfen",
                "Abbrechen"   //, false
            );

            if (!isYes) { return; }

            var topRowIndex = addressDGV.FirstDisplayedScrollingRowIndex;
            var currentId = (addressBSource.Current as Adresse)?.Id;
            try
            {
                _isFiltering = true;
                Cursor = Cursors.WaitCursor;
                addressDGV.UseWaitCursor = true;
                addressDGV.DataSource = null;  // NUR das Grid abkoppeln. Die Textboxen bleiben an der addressBindingSource hängen
                await DbChangeAnalyzer.RevertChangesAsync(analysis.RealChanges, addressBSource);  // Änderungen in EF rückgängig machen
                Utils.SortAddresses(addressBSource);  // Sortierung (jetzt via DataSource-Tausch, extrem schnell)
                foreach (var entry in _context.ChangeTracker.Entries().Where(x => x.State != EntityState.Unchanged)) { entry.State = EntityState.Unchanged; }
            }
            catch (Exception ex)
            {
                Utils.ErrTaskDlg(Handle, ex);
                return; // Abbruch bei Fehler
            }
            finally
            {
                addressDGV.DataSource = addressBSource;
                _isFiltering = false;

                addressDGV.UseWaitCursor = false;
                Cursor = Cursors.Default;

                addressBSource.ResetBindings(false);
                UpdateSaveButton();
            }

            // 3. Selektion und Scroll-Position sanft wiederherstellen
            var positionRestored = false;

            if (currentId.HasValue)
            {
                var itemToSelect = addressBSource.List.OfType<Adresse>().FirstOrDefault(a => a.Id == currentId.Value);
                if (itemToSelect is not null)
                {
                    var newIndex = addressBSource.IndexOf(itemToSelect);
                    if (newIndex >= 0)
                    {
                        positionRestored = true;
                        // Wir nutzen InvokeAsync, um dem Grid Zeit zum Neuzeichnen zu geben
                        _ = addressDGV.InvokeAsync(() => SyncGridToPosition(addressDGV, addressBSource, newIndex, true));
                    }
                }
            }

            // Falls der Datensatz weg ist (z.B. neu erstellt und dann verworfen), springen wir nach oben
            if (!positionRestored) { _ = addressDGV.InvokeAsync(SelectFirstAddressRow); }
        }
    }

    private void EditToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage)
        {
            // Nutzt jetzt die verfeinerte EF-Logik (ohne Phantom-Änderungen)
            rejectChangesToolStripMenuItem.Enabled = HasRealEFChanges();

            copyToOtherDGVTSMenuItem.Text = "Zu Google-&Kontakte hinzufügen";
            copyToOtherDGVTSMenuItem.Enabled = addressDGV.SelectedRows.Count > 0 && contactDGV.Rows.Count > 0;
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            // Ersetzt OldContactChanges_Check() durch den präzisen Snapshot-Vergleich
            rejectChangesToolStripMenuItem.Enabled = HasRealContactChanges(_lastActiveContact, _originalContactSnapshot);

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
        saveFileDialog.Title = "Wähle einen Speicherort";
        saveFileDialog.FileName = "GoogleKontakte.adb";
        saveFileDialog.InitialDirectory = Directory.Exists(_settings.DatabaseFolder) ? _settings.DatabaseFolder : Path.GetDirectoryName(_databaseFilePath);
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
                    Text = $"Die Google-Kontakte wurden erfolgreich in\n{backupPath} gespeichert.\n\nMöchtest du die Datei jetzt öffnen?",
                    Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
                    Footnote = "Bitte beachte, dass das Backup insofern unvollständig ist, dass nur\ndie in diesem Programm verwendeten Felder gesichert wurden.",
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
                        if (addressBSource != null) { await SaveSQLDatabaseAsync(true); }
                        await ConnectSQLDatabaseAsync(backupPath);
                        SetSearchTextIgnoreChange(string.Empty);
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

    private void BirthdaysToolStripMenuItem_Click(object sender, EventArgs e) => BirthdayReminder(tabControl.SelectedTab == addressTabPage ? addressDGV : contactDGV, true);

    private void BirthdayReminder(DataGridView dgv, bool showIfEmpty = false)
    {
        if (dgv.DataSource is not BindingSource bs) { return; }
        var isLocal = (dgv == addressDGV);
        var autoShow = isLocal ? _settings.BirthdayAddressShow : _settings.BirthdayContactShow;
        if (!showIfEmpty && !autoShow) { return; }
        IEnumerable<IContactEntity>? source = isLocal ? _context?.Adressen.Local : _allGoogleContacts;
        if (source == null || (!source.Any() && !showIfEmpty)) { return; }
        var bevorstehendeGeburtstage = Utils.CalculateUpcomingBirthdays(source, _settings.BirthdayRemindAfter, _settings.BirthdayRemindLimit);
        if (bevorstehendeGeburtstage.Count > 0 || showIfEmpty)
        {
            using var frm = new FrmBirthdays(_settings, bevorstehendeGeburtstage, isLocal);
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                SettingsManager.Save(_settings, _settingsPath);
                if (frm.SelectionIndex >= 0)
                {
                    var selectedId = bevorstehendeGeburtstage[frm.SelectionIndex].Id;
                    var item = bs.List.Cast<IContactEntity>().FirstOrDefault(x => x.UniqueId == selectedId);
                    if (item != null)
                    {
                        bs.Position = bs.IndexOf(item);
                        if (dgv.CurrentRow != null) { dgv.FirstDisplayedScrollingRowIndex = dgv.CurrentRow.Index; }
                        //if (!isLocal && contactBindingSource.Current is Contact selectedContact) { ShowPhotoInPictureBox(selectedContact); }
                    }
                }
            }
            searchTSTextBox.Focus();
        }
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
        // Sicherheitscheck: Nur prüfen, wenn wir im Kontakt-Tab arbeiten
        if (tabControl.SelectedTab == contactTabPage && _lastActiveContact != null)
        {
            // Wir prüfen zuerst, ob es überhaupt ECHTE Änderungen gibt.
            // Falls nicht, muss das Menü gar nicht erst abgebrochen werden (bessere Performance).
            if (HasRealContactChanges(_lastActiveContact, _originalContactSnapshot))
            {
                // 1. Menü-Öffnen sofort abbrechen, um Platz für den asynchronen Dialog zu machen
                e.Cancel = true;

                // 2. Den zentralen Gatekeeper nutzen.
                // Dieser kümmert sich um:
                // - Neue leere Kontakte (TidyUp)
                // - Änderungen speichern (AskSave)
                // - Änderungen verwerfen (CopyFromSnapshot)
                var readyToProceed = await ContactChanges_Check();

                // 3. Wenn der User fertig ist (Speichern/Verwerfen) und NICHT "Abbrechen" geklickt hat:
                if (readyToProceed)
                {
                    // Das Menü automatisch wieder öffnen, damit der User seinen Klick nicht wiederholen muss
                    if (sender is ToolStripDropDown dropDown && dropDown.OwnerItem is ToolStripDropDownItem ownerItem)
                    {
                        ownerItem.ShowDropDown();
                    }
                }
            }
        }
    }

    private void RecentToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        recentToolStripMenuItem.DropDownItems.Clear();
        var first = true;
        foreach (var file in _settings.RecentFiles)
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
                if (addressBSource != null)
                {
                    // Jetzt funktioniert await, weil das Lambda async ist
                    await SaveSQLDatabaseAsync(true);
                }

                // ConnectSQLDatabase wird erst ausgeführt, wenn SaveSQLDatabaseAsync fertig ist
                //ConnectSQLDatabase(file);
                await ConnectSQLDatabaseAsync(file);
                SetSearchTextIgnoreChange(string.Empty);
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
            newTSButton.Visible = copyTSButton.Visible = deleteTSButton.Visible = clipboardTSButton.Visible = wordTSButton.Visible = envelopeTSButton.Visible = detailSeparator1.Visible = detailSeparator2.Visible = true;
            dokuPlusTSButton.Visible = dokuMinusTSButton.Visible = dokuShowTSButton.Visible = dokuSeparator1.Visible = dokuSeparator2.Visible = false;
        }
        else if (tabulation.SelectedTab == tabPageDoku)
        {
            newTSButton.Visible = copyTSButton.Visible = deleteTSButton.Visible = clipboardTSButton.Visible = wordTSButton.Visible = envelopeTSButton.Visible = detailSeparator1.Visible = detailSeparator2.Visible = false;
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
        var imageFilter = string.Join(";", imageTypes);
        var allSupported = $"{documentFilter};{imageFilter}";
        openFileDialog.Filter =
            $"Alle unterstützten Dateien|{allSupported}|" +
            $"Dokumente ({documentFilter})|{documentFilter}|" +
            $"Bilder ({imageFilter})|{imageFilter}|" +
            $"Alle Dateien (*.*)|*.*";
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
        if (addressBSource?.Current is not Adresse selectedAddress) { return; }

        // 1. Liste der aktuellen Dateipfade aus der GUI holen
        var currentUiPaths = new HashSet<string>(dokuListView.Items.Cast<ListViewItem>().Select(i => i.Text), StringComparer.OrdinalIgnoreCase);

        // 2. Zu löschende Elemente finden (sind in DB, aber nicht mehr in GUI)
        // Wir erstellen eine separate Liste mit ToList(), um die Collection während der Iteration modifizieren zu können.
        var itemsToDelete = selectedAddress.Dokumente
            .Where(doc => !currentUiPaths.Contains(doc.Dateipfad))
            .ToList();

        foreach (var doc in itemsToDelete)
        {
            selectedAddress.Dokumente.Remove(doc);
        }

        // 3. Neue Elemente finden (sind in GUI, aber noch nicht in DB)
        var existingDbPaths = new HashSet<string>(selectedAddress.Dokumente.Select(d => d.Dateipfad), StringComparer.OrdinalIgnoreCase);

        //foreach (ListViewItem item in dokuListView.Items)
        //{
        //    if (!existingDbPaths.Contains(item.Text))
        //    {
        //        selectedAddress.Dokumente.Add(new Dokument
        //        {
        //            Dateipfad = item.Text,
        //            AdressId = selectedAddress.Id,
        //            Adresse = selectedAddress
        //        });
        //    }
        //}
        // Vereinfachte Version innerhalb von SyncDocumentsToEntity
        foreach (ListViewItem item in dokuListView.Items)
        {
            if (!existingDbPaths.Contains(item.Text))
            {
                selectedAddress.Dokumente.Add(new Dokument { Dateipfad = item.Text });  // EF Core setzt automatisch die AdressId und die Navigationseigenschaft Adresse
            }
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
        // 1. Sperre: Keine neuen Events verarbeiten, wenn der Dialog gerade offen ist
        if (_isDialogVisible) { return; }

        // 2. Temp-Dateien von Word (~$) und PDF-Programm (~) komplett ignorieren
        if (e.Name is { Length: > 0 } name && name.StartsWith('~')) { return; }

        debounceTimer.Stop();
        Debug.WriteLine($"ChangedEvent: {e.ChangeType} - {e.FullPath} - {e.Name}");

        debounceTimer.Tag = e.FullPath;
        if (!string.IsNullOrEmpty(e.FullPath)) { debounceTimer.Start(); }
    }

    private void FileSystemWatcher_OnRenamed(object sender, RenamedEventArgs e)
    {
        if (_isDialogVisible) { return; }
        if (e.Name is { Length: > 0 } name && name.StartsWith('~')) { return; }

        debounceTimer.Stop();
        Debug.WriteLine($"RenamedEvent: {e.ChangeType} - {e.FullPath}");

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
        else
        {
            item = new ListViewItem([info.FullName, string.Empty, string.Empty]) { ForeColor = Color.Red };
        }
        var vorhandenesItem = dokuListView.Items.Cast<ListViewItem>().FirstOrDefault(i => string.Equals(i.Text, info.FullName, StringComparison.OrdinalIgnoreCase));
        // SubItems über .Count prüfen, nicht auf null (verhindert ArgumentOutOfRangeException)
        //if (vorhandenesItem != null && vorhandenesItem.SubItems.Count > 2)
        //{
        //    vorhandenesItem.SubItems[1].Text = item.SubItems[1].Text;
        //    vorhandenesItem.SubItems[2].Text = item.SubItems[2].Text;
        //}
        if (vorhandenesItem != null)
        {
            // Wir aktualisieren die SubItems und die Farbe
            if (vorhandenesItem.SubItems.Count > 2 && item.SubItems.Count > 2)
            {
                vorhandenesItem.SubItems[1].Text = item.SubItems[1].Text;
                vorhandenesItem.SubItems[2].Text = item.SubItems[2].Text;
            }
            vorhandenesItem.ForeColor = item.ForeColor;
            vorhandenesItem.ImageKey = item.ImageKey;
        }
        else
        {
            dokuListView.Items.Add(item);
        }
        if (sortAndSave)
        {
            dokuListView.ListViewItemSorter = new ListViewItemComparer();
            dokuListView.Sort();
        }
    }

    private void DebounceTimer_Tick(object sender, EventArgs e)
    {
        debounceTimer.Stop();

        var text = debounceTimer.Tag as string ?? string.Empty;
        if (string.IsNullOrEmpty(text) || !File.Exists(text)) { return; }

        // 3. Neu: Prüfen, ob das PDF-Programm noch speichert
        if (!Utils.IsFileReady(text))
        {
            debounceTimer.Start();  // Datei ist noch blockiert -> Wir warten weitere 500ms
            return;
        }

        // Sperre setzen, damit keine neuen Events dazwischenfunken
        _isDialogVisible = true;

        try
        {


            if (!Visible) { RestoreFromTray(); }
            NativeMethods.SetForegroundWindow(Handle);
            var ort = cbOrt.Text;
            var nameEtc = string.Join(" ", new[] { tbVorname.Text, tbNachname.Text, tbFirma.Text }.Where(s => !string.IsNullOrWhiteSpace(s)));
            var inOrt = string.IsNullOrWhiteSpace(ort) ? "" : $" in {ort}";
            TaskDialogButton linkButton = new TaskDialogCommandLinkButton("Mit Adresse verknüpfen", $"{nameEtc}{inOrt}");
            TaskDialogButton nextButton = new TaskDialogCommandLinkButton("Eine andere Adresse wählen…", "… und neuen Dialog bestätigen");
            TaskDialogButton copyButton = new TaskDialogCommandLinkButton("In Zwischenablage kopieren", "Briefe lassen sich auch manuell hinzufügen.");
            using TaskDialogIcon questionDialogIcon = new(Resources.question32);
            var page = new TaskDialogPage
            {
                Caption = appName,
                Heading = "Änderung im Briefordner erkannt",
                Text = $"Datei: {text}",
                Icon = questionDialogIcon,  // TaskDialogIcon.ShieldWarningYellowBar,
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
                //using TaskDialogIcon questionDialogIcon = new(Resources.question32);
                var next = new TaskDialogPage
                {
                    Caption = appName,
                    Heading = "Möchtest du die Datei verknüpfen?",
                    Text = $"{text}",
                    Icon = questionDialogIcon,
                    Footnote = $"WÄHLE DIE PASSENDE ADRESSE, bevor du auf 'Ja' klickst.",
                    Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
                    AllowCancel = true,
                    SizeToContent = true
                };
                next.Created += static (sender, e) =>
                {
                    var dialogHandle = NativeMethods.GetActiveWindow();
                    if (dialogHandle != nint.Zero)
                    {
                        // Verschiebt das Fenster auf die TopMost-Ebene (HWND_TOPMOST), behält Position & Größe bei (NOMOVE, NOSIZE) und klaut der Suchbox NICHT den Fokus (NOACTIVATE).
                        NativeMethods.SetWindowPos(dialogHandle, NativeMethods.HWND_TOPMOST, 0, 0, 0, 0, NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
                    }
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
                        Utils.MsgTaskDlg(Handle, "Funktion nicht verfügbar", "Google-Kontakte unterstützen keine Verweise auf lokale Dateien.", TaskDialogIcon.Information);
                    }
                }
            }

            else if (result == copyButton)
            {
                try { Clipboard.SetText(text); }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
        }
        finally { _isDialogVisible = false; }  // Die Sperre am Ende garantiert wieder aufheben
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
        if (_isFiltering || IsDisposed) { return; }
        if (InvokeRequired)
        {
            _ = InvokeAsync(UpdateSaveButton);
            return;
        }
        static void UpdateTabTitle(TabPage tab, string baseTitle, bool hasChanges)
        {
            var targetText = hasChanges ? $"{baseTitle}*" : baseTitle;
            if (tab.Text != targetText) { tab.Text = targetText; }
        }
        var hasSqlChanges = HasRealEFChanges();
        var hasGoogleChanges = HasRealContactChanges(_lastActiveContact, _originalContactSnapshot);
        UpdateTabTitle(addressTabPage, " Lokale Adressen", hasSqlChanges);
        UpdateTabTitle(contactTabPage, " Google Kontakte", hasGoogleChanges);
        var (enabled, toolTip) = tabControl.SelectedTab switch  // "Switch Expression" für bessere Lesbarkeit
        {
            var t when t == addressTabPage => (hasSqlChanges, hasSqlChanges ? "Lokale Adressen speichern (Strg+S)" : "Keine Änderungen"),
            var t when t == contactTabPage => (hasGoogleChanges, hasGoogleChanges ? "Google-Kontakt hochladen (Strg+S)" : "Keine Änderungen"),
            _ => (false, "Keine Auswahl")
        };
        if (saveTSButton.Enabled != enabled) { saveTSButton.Enabled = enabled; }
        saveTSButton.ToolTipText = toolTip;
    }

    private bool HasRealEFChanges()
    {
        if (_context == null) { return false; }
        _context.ChangeTracker.DetectChanges();  // DetectChanges erkennt Added/Deleted bei Beziehungen automatisch
        foreach (var entry in _context.ChangeTracker.Entries())
        {
            if (entry.State == EntityState.Added || entry.State == EntityState.Deleted) { return true; }
            if (entry.State == EntityState.Modified)
            {
                if (entry.Metadata.ClrType == null) { return true; }  // selten, Phantom-Änderung ohne CLR-Typ, sicherheitshalber als echte Änderung behandeln
                foreach (var prop in entry.Properties)  // Bei echten Klassen prüfen wir die Properties auf relevante Änderungen
                {
                    if (!prop.IsModified) { continue; }
                    var original = prop.OriginalValue;
                    var current = prop.CurrentValue;
                    if (prop.Metadata.ClrType == typeof(string))
                    {
                        var sOriginal = (original as string ?? string.Empty).Trim();
                        var sCurrent = (current as string ?? string.Empty).Trim();
                        if (string.Equals(sOriginal, sCurrent, StringComparison.Ordinal)) { continue; }
                    }
                    else if (Equals(original, current)) { continue; }
                    return true; // Echte Property-Änderung gefunden
                }
            }
        }
        return false;
    }

    private async Task<bool> ContactChanges_Check(bool isClosing = false)
    {
        // 1. RE-ENTRANCY SCHUTZ: Wenn wir schon prüfen (Dialog ist offen), blockieren wir alles weitere.
        if (_isCheckingContactChanges) { return false; }

        if (_lastActiveContact == null || _originalContactSnapshot == null) { return true; }

        _isCheckingContactChanges = true; // Tür abschließen
        try
        {
            contactDGV.CausesValidation = false;
            addressDGV.CausesValidation = false;

            // Validierung im UI-Thread absichern
            var isValid = true;
            try
            {
                isValid = ValidateChildren(ValidationConstraints.Enabled);
                contactBSource.EndEdit();
            }
            finally
            {
                contactDGV.CausesValidation = true;
                addressDGV.CausesValidation = true;
            }

            if (!isValid) { return false; }

            var currentContact = _lastActiveContact;
            var isNewContact = string.IsNullOrEmpty(currentContact.ResourceName);

            if (!HasRealContactChanges(currentContact, _originalContactSnapshot))
            {
                if (isNewContact) { RemoveContactFromList(currentContact); }
                return true;
            }

            var result = await AskSaveContactChangesAsync(isClosing);

            if (result == DialogResult.Cancel) { return false; }
            if (result == DialogResult.No)
            {
                if (isNewContact) { RemoveContactFromList(currentContact); }
                else
                {
                    currentContact.CopyFrom(_originalContactSnapshot);
                    if (!isClosing) { contactBSource.ResetCurrentItem(); }
                }
            }
            return true;
        }
        finally
        {
            _isCheckingContactChanges = false; // Tür wieder aufschließen
        }
    }

    private async Task CheckContactChanges(Func<Task> action)  // Func<Task> action -> Erwartet eine Methode, die Task zurückgibt
    {
        if (await ContactChanges_Check()) { await action(); }
    }

    private async Task CheckContactChanges(Action action)  // Rückgabetyp hier auch 'Task' statt 'void', damit man es awaiten kann!
    {
        if (await ContactChanges_Check()) { action(); }
    }

    private async Task<DialogResult> AskSaveContactChangesAsync(bool isClosing)
    {
        if (_originalContactSnapshot == null || _lastActiveContact == null) { return DialogResult.None; }

        //ValidateChildren();  // Sicherstellen, dass die UI aktuell ist
        //contactBindingSource.EndEdit();

        var changedFields = _lastActiveContact.GetChangedFields(_originalContactSnapshot);

        //var photoChanged = changedFields.Remove("photos");  // wird separat behandelt, aber für die API-Logik müssen wir wissen, ob es sich geändert hat, siehe unten

        if (changedFields.Count == 0) { return DialogResult.None; }

        var nameParts = new[] { _lastActiveContact.Vorname, _lastActiveContact.Nachname }.Where(s => !string.IsNullOrWhiteSpace(s));
        var fullName = string.Join(" ", nameParts);
        var headingText = "Möchtet du die Änderung" + (changedFields.Count > 1 ? "en" : "") + " speichern?";

        var fieldList = string.Join("\n", changedFields.Select(f => "• " + char.ToUpper(f[0]) + f[1..]));
        //if (photoChanged) { fieldList += "\n• Foto"; }

        var shortSummary = changedFields.Count == 1 ? $"Ein Bereich wurde geändert:\n{fieldList}" : $"{changedFields.Count} Bereiche wurden geändert:\n{fieldList}";

        var detailedDiff = Utils.GenerateDetailedDiff(_lastActiveContact, _originalContactSnapshot, dataFields);

        var btnSave = new TaskDialogButton("&Hochladen") { AllowCloseDialog = false }; // Wichtig: Schließt nicht sofort
        var btnDiscard = new TaskDialogButton("&Verwerfen");
        var btnCancel = TaskDialogButton.Cancel;
        using TaskDialogIcon questionDialogIcon = new(Resources.question32);
        var pageMain = new TaskDialogPage()
        {
            Caption = "Google Kontakte",
            Heading = headingText,
            Text = shortSummary, // detailedDiff hier entfernt
            Icon = questionDialogIcon,  // TaskDialogIcon.ShieldBlueBar,
            AllowCancel = true,
            Buttons = { btnSave, btnDiscard, btnCancel },
            DefaultButton = btnSave,
            Expander = new TaskDialogExpander()
            {
                Text = detailedDiff,
                Position = TaskDialogExpanderPosition.AfterText // Platziert den Expander direkt unter dem Haupttext
            }
        };

        _googleCts?.Dispose();
        _googleCts = new CancellationTokenSource();
        var token = _googleCts.Token;

        var pageProgress = new TaskDialogPage()
        {
            Caption = "Google Kontakte",
            Heading = "Bitte warten…",
            Text = "Daten werden an Google übertragen.",
            Icon = TaskDialogIcon.Information,
            ProgressBar = new TaskDialogProgressBar() { State = TaskDialogProgressBarState.Marquee },
            Buttons = { TaskDialogButton.Close }
        };

        pageProgress.Buttons[0].Enabled = false; // "Schließen" erst nach Abschluss erlauben
        var saveSuccess = false;  // Status-Flag für den Rückgabewert

        btnSave.Click += (s, e) => { pageMain.Navigate(pageProgress); };

        pageProgress.Created += async (s, e) =>
        {
            try
            {
                var currentImage = topAlignZoomPictureBox.Image;

                await ExecuteGoogleSaveAsync(_lastActiveContact, changedFields, currentImage, token);

                saveSuccess = true;

                // UI Feedback im Dialog
                pageProgress.ProgressBar.Value = 100;
                pageProgress.ProgressBar.State = TaskDialogProgressBarState.Normal;

                // Sauberer Schließbefehl für den TaskDialog
                pageProgress.BoundDialog?.Close();
            }
            catch (Exception ex)
            {
                // Fehlerbehandlung im Dialog
                pageProgress.Heading = "Fehler beim Speichern";
                pageProgress.Text = ex.Message; // Ggf. Stacktrace kürzen
                pageProgress.Icon = TaskDialogIcon.Error;
                pageProgress.ProgressBar.State = TaskDialogProgressBarState.Error;
                pageProgress.Buttons[0].Enabled = true; // User muss Button klicken zum Schließen
            }
        };

        // Dialog anzeigen
        var clickedButton = TaskDialog.ShowDialog(Handle, pageMain);

        // Rückgabe ermitteln
        if (saveSuccess)
        {
            if (!isClosing)
            {
                saveTSButton.Enabled = false;
                //var position = contactBindingSource.IndexOf(_lastActiveContact);
                //if (position >= 0)
                //{
                //  _ =  InvokeAsync(() => { contactBindingSource.ResetItem(position); });  // Aktualisiert nur diese eine Zeile im Grid, statt alle Zeilen zu löschen
                //}
            }
            return DialogResult.Yes;
        }
        if (clickedButton == btnDiscard) { return DialogResult.No; }

        return DialogResult.Cancel;
    }

    private async Task ExecuteGoogleSaveAsync(Contact contactToSave, List<string> changedFields, Image? currentImage, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var manager = new GooglePeopleManager(secretPath, tokenDir);

        if (string.IsNullOrEmpty(contactToSave.ResourceName))
        {
            var createdContact = await manager.CreateContactAsync(contactToSave, currentImage, token);
            contactToSave.ResourceName = createdContact.ResourceName;
            contactToSave.ETag = createdContact.ETag;
            contactToSave.PhotoUrl = createdContact.PhotoUrl;
        }
        else // === FALL B: UPDATE ===
        {
            if (changedFields.Count > 0 || changedFields.Contains("memberships"))
            {
                var retry = false;
                do
                {
                    retry = false;
                    try
                    {
                        var updatedPerson = await manager.UpdateContactAsync(contactToSave, changedFields, contactGroupsDict, _originalContactSnapshot, checkEmptyGroups: true, token: token);
                        contactToSave.ETag = updatedPerson.ETag;
                        contactToSave.ResourceName = updatedPerson.ResourceName;
                    }
                    catch (Google.GoogleApiException ex) when (ex.HttpStatusCode == System.Net.HttpStatusCode.BadRequest || ex.HttpStatusCode == System.Net.HttpStatusCode.Conflict)
                    {
                        // Prüfen, ob der Fehler durch einen ETag-Mismatch (Konflikt) ausgelöst wurde
                        if (ex.Message.Contains("etag", StringComparison.OrdinalIgnoreCase) || ex.Message.Contains("precondition", StringComparison.OrdinalIgnoreCase))
                        {
                            var overwrite = false;

                            // 1. UI blockieren und User fragen (Invoke ist Pflicht, da wir im Task-Thread sind!)
                            Invoke(() =>
                            {
                                var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Konflikt beim Speichern",
                                    "Dieser Kontakt wurde in der Zwischenzeit online (z. B. auf einem Smartphone) geändert.",
                                    "Möchtest du das Speichern erzwingen und die Online-Version mit deinen lokalen Eingaben überschreiben?");
                                overwrite = isYes;
                            });

                            if (overwrite)
                            {
                                // 2. Den GANZ FRISCHEN Kontakt von Google holen, nicht nur den ETag!
                                // Sonst würden wir mit unserem alten RawGooglePerson-Payload die Online-Änderungen überschreiben.
                                var freshPerson = await manager.GetRawPersonAsync(contactToSave.ResourceName, token);

                                contactToSave.RawGooglePerson = freshPerson; // Den Hintergrund-Payload aktualisieren
                                contactToSave.ETag = freshPerson.ETag;       // Den neuen ETag anheften

                                retry = true; // Schleife von vorne -> UpdateContactAsync mischt jetzt unsere lokalen Änderungen in die neuen Online-Daten!
                            }
                            else
                            {
                                // 3. User hat abgebrochen: Nichts weiter tun, Methode verlassen.
                                return;
                            }
                        }
                        else
                        {
                            // Ein anderer Bad Request Fehler (z.B. ungültige E-Mail-Adresse), diesen werfen wir ganz normal weiter
                            throw;
                        }
                    }
                } while (retry);
            }
            token.ThrowIfCancellationRequested();
        }

        _originalContactSnapshot = (Contact)contactToSave.Clone();
        contactToSave.ResetSearchCache();
    }


    private void RemoveContactFromList(Contact contact)
    {
        isSelectionChanging = true;
        try
        {
            // WICHTIG: Immer über die BindingSource löschen, um das DataGridView synchron zu halten!
            contactBSource.Remove(contact);

            _lastActiveContact = null;
            _originalContactSnapshot = null;
        }
        finally { isSelectionChanging = false; }
    }

    private bool HasRealContactChanges(Contact? current, Contact? original)
    {
        // 1. Schnelle Referenz- und Null-Prüfung
        if (ReferenceEquals(current, original)) { return false; }
        if (current is null || original is null) { return true; }

        var type = typeof(Contact);

        // 2. Iteration über alle Standard-Felder (Strings & Datum)
        foreach (var fieldName in dataFields)
        {
            var prop = type.GetProperty(fieldName);
            if (prop == null) { continue; } // Sicherheitscheck

            var valCurrent = prop.GetValue(current);
            var valOriginal = prop.GetValue(original);

            // Unterscheidung String vs. Rest (z.B. DateOnly/DateTime)
            if (prop.PropertyType == typeof(string))
            {
                // Strings: null und "" als gleich behandeln
                var s1 = (valCurrent as string) ?? string.Empty;
                var s2 = (valOriginal as string) ?? string.Empty;

                if (!string.Equals(s1, s2, StringComparison.Ordinal)) { return true; }
            }
            else
            {
                // Werttypen (z.B. Geburtstag): Standard-Vergleich
                if (!Equals(valCurrent, valOriginal)) { return true; }
            }
        }

        // 3. Spezialfelder (nicht in dataFields enthalten)

        // Foto (String-Vergleich, aber war nicht im Array)
        //if (!string.Equals(current.PhotoUrl ?? string.Empty, original.PhotoUrl ?? string.Empty, StringComparison.Ordinal)) { return true; }

        // Gruppen (Listen-Vergleich)
        var currentGroups = current.GroupNames ?? [];
        var originalGroups = original.GroupNames ?? [];

        if (currentGroups.Count != originalGroups.Count) { return true; }

        // SequenceEqual prüft, ob die Inhalte gleich sind (sortiert, um Reihenfolge zu ignorieren)
        if (!currentGroups.OrderBy(x => x).SequenceEqual(originalGroups.OrderBy(x => x))) { return true; }

        return false;
    }

    private async void Clear_SearchTextBox()
    {
        filterRemoveToolStripMenuItem.Visible = false;

        var isAddressTab = tabControl.SelectedTab == addressTabPage;
        var activeBs = isAddressTab ? addressBSource : contactBSource;
        var activeDgv = isAddressTab ? addressDGV : contactDGV;
        var selectedItem = activeBs?.Current;

        SetSearchTextIgnoreChange(string.Empty);
        searchTimer.Stop();

        // WICHTIG: jumpToFirstRow auf false, da wir das Item manuell ansteuern
        ApplyGlobalSearch(string.Empty, jumpToFirstRow: false);

        if (activeBs != null && activeDgv != null && selectedItem != null)
        {
            var newIndex = activeBs.IndexOf(selectedItem);
            if (newIndex >= 0)
            {
                _ = activeDgv.InvokeAsync(() =>
                {
                    // Wir erzwingen hier kein Zentrieren (-1), sondern bleiben 
                    // beim Standard-Scrollverhalten für den manuellen Reset.
                    SyncGridToPosition(activeDgv, activeBs, newIndex, setFocus: false);
                });
            }
        }
        await Task.Yield();  // kurzer Atemzug, damit die UI die Änderungen rendern kann, bevor wir den Fokus setzen
        searchTSTextBox.Focus();
    }

    private void WebsiteToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartLink(Handle, @"https://www.netradio.info/address");

    private void GithubToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartLink(Handle, @"https://github.com/ophthalmos/Adressen");

    private void HelpdokuTSMenuItem_Click(object sender, EventArgs e) => Utils.StartFile(Handle, Path.Combine(Path.GetDirectoryName(appPath) ?? string.Empty, "AdressenKontakte.pdf"));

    private void TermsofuseToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartLink(Handle, "https://www.netradio.info/adressen-terms-of-use/");
    private void PrivacypolicyToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartLink(Handle, "https://www.netradio.info/adressen-privacy-policy/");
    private void LicenseTxtToolStripMenuItem_Click(object sender, EventArgs e) => Utils.StartFile(Handle, Path.Combine(Path.GetDirectoryName(appPath) ?? string.Empty, "Lizenzvereinbarung.txt"));

    private void AdressenMitBriefToolStripMenuItem_Click(object sender, EventArgs e)  // gibt es nur bei Adressen
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV, a => a.Dokumente.Count != 0, "… mit Briefverweis", "Adressen");
        }
    }

    private void AdressenOhneBriefToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV, a => a.Dokumente.Count == 0, "… ohne Briefverweis", "Adressen");
        }
    }

    private void PhotoPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // 1. Wir fragen die DB: Welche IDs haben ein Foto? ("SELECT Id FROM Adressen WHERE FotoId IS NOT NULL")
            var idsWithPhoto = _context.Adressen.Where(a => a.Foto != null).Select(a => a.Id).ToHashSet(); // HashSet für extrem schnelle Suche
            // 2. Wir filtern die lokale Liste anhand dieser IDs
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV, a => idsWithPhoto.Contains(a.Id), "… mit Bild", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV, c => !string.IsNullOrWhiteSpace(c.PhotoUrl), "… mit Bild", "Google Kontakte");
        }
    }

    private void PhotoMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // 1. Gleiches Spiel: IDs holen
            var idsWithPhoto = _context.Adressen.Where(a => a.Foto != null).Select(a => a.Id).ToHashSet();
            // 2. Filter umdrehen: Zeige alle, deren ID NICHT in der Liste ist
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV, a => !idsWithPhoto.Contains(a.Id), "… ohne Bild", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV, c => string.IsNullOrWhiteSpace(c.PhotoUrl), "… ohne Bild", "Google Kontakte");
        }
    }

    private void MailPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV, a => !string.IsNullOrWhiteSpace(a.Mail1) || !string.IsNullOrWhiteSpace(a.Mail2), "… mit E-Mail", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV, c => !string.IsNullOrWhiteSpace(c.Mail1) || !string.IsNullOrWhiteSpace(c.Mail2), "… mit E-Mail", "Google Kontakte");
        }
    }

    private void MailMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV,
                a => string.IsNullOrWhiteSpace(a.Mail1) && string.IsNullOrWhiteSpace(a.Mail2), "… ohne E-Mail", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV,
                c => string.IsNullOrWhiteSpace(c.Mail1) && string.IsNullOrWhiteSpace(c.Mail2), "… ohne E-Mail", "Google Kontakte");
        }
    }

    private void TelephonePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV,
                a => !string.IsNullOrWhiteSpace(a.Telefon1) || !string.IsNullOrWhiteSpace(a.Telefon2), "… mit Telefonnummer", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV,
                c => !string.IsNullOrWhiteSpace(c.Telefon1) || !string.IsNullOrWhiteSpace(c.Telefon2), "… mit Telefonnummer", "Google Kontakte");
        }
    }

    private void TelephoneMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV,
                a => string.IsNullOrWhiteSpace(a.Telefon1) && string.IsNullOrWhiteSpace(a.Telefon2), "… ohne Telefonnummer", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV,
                c => string.IsNullOrWhiteSpace(c.Telefon1) && string.IsNullOrWhiteSpace(c.Telefon2), "… ohne Telefonnummer", "Google Kontakte");
        }
    }

    private void MobilePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV,
                a => !string.IsNullOrWhiteSpace(a.Mobil), "… mit Mobilnummer", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV,
                c => !string.IsNullOrWhiteSpace(c.Mobil), "… mit Mobilnummer", "Google Kontakte");
        }
    }

    private void MobileMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV, a => string.IsNullOrWhiteSpace(a.Mobil), "… ohne Mobilnummer", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV, c => string.IsNullOrWhiteSpace(c.Mobil), "… ohne Mobilnummer", "Google Kontakte");
        }
    }

    private void DatePlusFilterMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // Wir prüfen direkt das Nullable DateOnly Feld "Geburtstag"
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV,
                a => a.Geburtstag.HasValue, "… mit Geburtsdatum", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            // Auch für Google Kontakte (vorausgesetzt, das Feld heißt dort ähnlich)
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV,
                c => c.Geburtstag.HasValue, "… mit Geburtsdatum", "Google Kontakte");
        }
    }

    private void DateMinusFilterMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // Wir filtern auf alle Adressen, deren Geburtstag NICHT gesetzt ist (null)
            ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV,
                a => !a.Geburtstag.HasValue, "… ohne Geburtsdatum", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            // Dieselbe Logik für die Google Kontakte
            ExecuteFilter(_allGoogleContacts, contactBSource, contactDGV,
                c => !c.Geburtstag.HasValue, "… ohne Geburtsdatum", "Google Kontakte");
        }
    }

    private void OrphanedDocumentsToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (_context == null || tabControl.SelectedTab != addressTabPage) { return; }
        ExecuteFilter(_context.Adressen.Local, addressBSource, addressDGV, a => a.Dokumente.Any(d => !File.Exists(d.Dateipfad)), "… mit verwaisten Dokumenten", "Adressen");
        if (addressBSource.Count > 0)
        {
            tabulation.SelectedTab = tabPageDoku;
            ShowOrphanCleanupDialog();
        }
    }

    private async void ShowOrphanCleanupDialog()
    {
        var affectedAddresses = addressBSource.Cast<Adresse>().Where(a => a.Dokumente.Any(d => !File.Exists(d.Dateipfad))).ToList();
        var addressCount = affectedAddresses.Count;
        var orphanCount = affectedAddresses.Sum(a => a.Dokumente.Count(d => !File.Exists(d.Dateipfad)));
        if (orphanCount == 0)
        {
            Utils.MsgTaskDlg(Handle, "Alles sauber!", "Keine verwaisten Dokumente gefunden.");
            return;
        }
        var cleanLink = new TaskDialogCommandLinkButton("&Verknüpfungen löschen", "Entfernt die ungültigen Pfadangaben.");
        var replaceLink = new TaskDialogCommandLinkButton("&Dateipfade anpassen", "Ersetzt Pfadangaben oder Teile davon.");
        var closeButton = TaskDialogButton.Close;
        var addressText = addressCount == 1 ? "einer Adresse" : $"{addressCount} Adressen";
        var page = new TaskDialogPage()
        {
            Caption = appName,
            Heading = "Verwaiste Dokumente gefunden",
            Text = $"Es wurden {orphanCount} Dateiverknüpfungen bei {addressText}\ngefunden, deren Zieldateien nicht mehr existieren.",
            Icon = TaskDialogIcon.ShieldWarningYellowBar,
            Buttons = { cleanLink, replaceLink, closeButton }
        };

        page.Created += static (s, e) =>
        {
            var dialogHandle = NativeMethods.GetActiveWindow();
            if (dialogHandle != nint.Zero) { NativeMethods.SetWindowPos(dialogHandle, NativeMethods.HWND_TOPMOST, 0, 0, 0, 0, NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE); }
        };
        var result = TaskDialog.ShowDialog(page);
        if (result == cleanLink)  // OPTION 1: Löschen
        {
            var totalRemoved = 0;
            foreach (var adresse in affectedAddresses)
            {
                var toRemove = adresse.Dokumente.Where(d => !File.Exists(d.Dateipfad)).ToList();
                foreach (var doc in toRemove)
                {
                    adresse.Dokumente.Remove(doc);
                    totalRemoved++;
                }
            }
            if (totalRemoved > 0)
            {
                if (addressBSource.Current is Adresse currentAddress) { UpdateDocumentListView(currentAddress); }
                UpdateSaveButton();
                Utils.MsgTaskDlg(Handle, "Bereinigung abgeschlossen", $"{totalRemoved} verwaiste Verknüpfungen  bei {addressText} wurden entfernt.\n\nDer Filter wird nun zurückgesetzt (Anzeige aller Adressen).", TaskDialogIcon.ShieldSuccessGreenBar);
                tabulation.SelectedTab = tabPageDetail;
                FilterRemoveToolStripMenuItem_Click(this, EventArgs.Empty);
            }
        }
        else if (result == replaceLink)  // OPTION 2: Suchen & Ersetzen
        {
            var commonRoot = string.Empty;
            if (_context != null)
            {
                var mostFrequentBrokenPath = _context.Dokumente.Local
                    .Where(d => !File.Exists(d.Dateipfad))
                    .Select(d =>
                    {
                        try { return Path.GetDirectoryName(d.Dateipfad) ?? string.Empty; }
                        catch { return string.Empty; }
                    })
                    .Where(p => !string.IsNullOrEmpty(p)).GroupBy(p => p, StringComparer.OrdinalIgnoreCase).OrderByDescending(g => g.Count()).FirstOrDefault();
                if (mostFrequentBrokenPath != null) { commonRoot = mostFrequentBrokenPath.Key; }
            }
            using var dialog = new PathReplacement();
            dialog.SearchText = commonRoot;
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                var oldPath = dialog.SearchText.Trim();
                var newPath = dialog.ReplaceText.Trim();
                if (!string.IsNullOrEmpty(oldPath))  // Wenn was eingegeben wurde, asynchron ersetzen
                {
                    var (docCount, addrCount) = await ExecutePathReplacementAsync(oldPath, newPath);
                    if (docCount > 0)
                    {
                        var docText = docCount == 1 ? "Eine Verknüpfung" : $"{docCount} Verknüpfungen";
                        var addrText = addrCount == 1 ? "einer Adresse" : $"{addrCount} Adressen";
                        var verbText = docCount == 1 ? "wurde" : "wurden";
                        Utils.MsgTaskDlg(Handle, "Pfade aktualisiert", $"{docText} bei {addrText} {verbText} erfolgreich angepasst.", TaskDialogIcon.ShieldSuccessGreenBar);
                        OrphanedDocumentsToolStripMenuItem_Click(this, EventArgs.Empty);  // Filter neu anwenden, um die reparierten Adressen ausblenden zu lassen!
                    }
                    else { Utils.MsgTaskDlg(Handle, "Keine Änderungen", "Es wurden keine verwaisten Pfade ge-\nfunden, die diesen Text enthalten.", TaskDialogIcon.Information); }
                }
            }
        }
    }

    private async Task<(int docCount, int addrCount)> ExecutePathReplacementAsync(string oldPath, string newPath)
    {
        if (string.IsNullOrEmpty(oldPath) || _context == null) { return (0, 0); }
        var changedCount = 0;
        var affectedAddresses = new HashSet<Adresse>(); // Verhindert, dass Adressen doppelt gezählt werden
        await Task.Yield();  // damit die UI nicht einfriert, während wir die Pfade anpassen
        var documentsToUpdate = _context.Dokumente.Local.Where(d => !File.Exists(d.Dateipfad) && d.Dateipfad.Contains(oldPath, StringComparison.OrdinalIgnoreCase)).ToList();
        foreach (var doc in documentsToUpdate)
        {
            var updatedPath = doc.Dateipfad.Replace(oldPath, newPath, StringComparison.OrdinalIgnoreCase);
            if (doc.Dateipfad != updatedPath)
            {
                doc.Dateipfad = updatedPath;
                changedCount++;
                if (doc.Adresse != null) { affectedAddresses.Add(doc.Adresse); }
            }
        }
        if (changedCount > 0)
        {
            if (addressBSource.Current is Adresse currentAddress) { UpdateDocumentListView(currentAddress); }  // UI aktualisieren, falls die gerade angewählte Adresse betroffen ist
            UpdateSaveButton();
        }
        return (changedCount, affectedAddresses.Count);
    }

    private void FilterToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        var isAddressTab = tabControl.SelectedTab == addressTabPage && addressDGV.Rows.Count > 0;
        var isContactTab = tabControl.SelectedTab == contactTabPage && contactDGV.Rows.Count > 0;
        var enableCommon = isAddressTab || isContactTab;
        foreach (ToolStripItem item in filterToolStripMenuItem.DropDownItems)
        {
            if (item == adressenMitBriefToolStripMenuItem || item == adressenOhneBriefToolStripMenuItem || item == orphanedDocumentsToolStripMenuItem) { item.Enabled = isAddressTab; }
            else if (item is ToolStripMenuItem) { item.Enabled = enableCommon; }
        }
    }

    private void ExecuteFilter<T>(IEnumerable<T> sourceList, BindingSource bs, DataGridView dgv, Func<T, bool> predicate, string statusText, string entityName)
    {
        if (sourceList == null || bs == null) { return; }
        SetSearchTextIgnoreChange(string.Empty);
        searchTimer.Stop();
        _isFiltering = true; // Guard setzen
        var currencyManager = BindingContext?[bs] as CurrencyManager;
        try
        {
            dgv.SuspendLayout();
            currencyManager?.SuspendBinding();
            dgv.CurrentCell = null;
            var filteredList = sourceList.Where(predicate).ToList();
            bs.DataSource = new BindingList<T>(filteredList);  // Magie: Filtern und konsistente BindingList verwenden
            flexiTSStatusLabel.Text = statusText;
            var totalCount = sourceList.TryGetNonEnumeratedCount(out var count) ? count : sourceList.Count();  // falls sourceList keine echte ICollection ist
            var visibleCount = filteredList.Count;
            toolStripStatusLabel.Text = visibleCount == totalCount ? $"{totalCount} {entityName}" : $"{visibleCount}/{totalCount} {entityName}";
            if (visibleCount > 0 && dgv.Rows.Count > 0) { dgv.CurrentCell = dgv.Rows[0].Cells[0]; }
        }
        catch (Exception ex) { Debug.WriteLine(ex.Message); }
        finally
        {
            currencyManager?.ResumeBinding();
            dgv.ResumeLayout();
            _isFiltering = false;
            if (typeof(T) == typeof(Adresse)) { AddressBindingSource_CurrentChanged(bs, EventArgs.Empty); }  // Sichere Event-Auslösung mit dem richtigen Sender
            else if (typeof(T) == typeof(Contact)) { ContactBindingSource_CurrentChanged(bs, EventArgs.Empty); }
            UpdateFilterUIState();
        }
    }

    private void FilterRemoveToolStripMenuItem_Click(object sender, EventArgs e)
    {
        SetSearchTextIgnoreChange(string.Empty);  // Suche stumm leeren
        searchTimer.Stop();
        _isFiltering = true;
        try
        {
            if (tabControl.SelectedTab == addressTabPage)
            {
                if (_context == null) { return; }
                ExecuteAndPreserveSelection<Adresse>(addressBSource, addressDGV, () => { addressBSource.DataSource = _context.Adressen.Local.ToBindingList(); });
                UpdateAddressStatusBar();
            }
            else if (tabControl.SelectedTab == contactTabPage)
            {
                if (_allGoogleContacts != null)
                {
                    ExecuteAndPreserveSelection<Contact>(contactBSource, contactDGV, () => { contactBSource.DataSource = _allGoogleContacts; });
                }
                UpdateContactStatusBar();
            }
        }
        finally { _isFiltering = false; }
        flexiTSStatusLabel.Text = string.Empty;
        searchTSTextBox.Focus();
        UpdateFilterUIState();
    }

    private void UpdateFilterUIState()
    {
        var isSubset = false;

        // 1. Prüfen, ob eine Teilmenge im aktuell aktiven Tab angezeigt wird
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            var total = _context.Adressen.Local.Count;
            var current = addressBSource.Count;
            isSubset = current < total;
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            var total = _allGoogleContacts.Count;
            var current = contactBSource.Count;
            isSubset = current < total;
        }

        // 2. Prüfen, ob Text im Suchfeld steht
        var hasSearchText = !string.IsNullOrWhiteSpace(searchTSTextBox.Text);

        // 3. Sichtbarkeit setzen (Wenn Teilmenge ODER Suchtext vorhanden)
        filterRemoveToolStripMenuItem.Visible = isSubset || hasSearchText;

        // Das kleine 'X' im Suchfeld wird nur über den Text gesteuert
        tsClearLabel.Visible = hasSearchText;
    }

    private void ContextPhotoMenu_Opening(object sender, CancelEventArgs e)
    {
        copyToolStripMenuItem.Enabled = topAlignZoomPictureBox.Image != null && delPictboxToolStripButton.Enabled;
        var hasClipboardImage = Clipboard.ContainsImage();
        var isAddressValid = tabControl.SelectedTab == addressTabPage && addressBSource.Current != null;
        var isContactValid = tabControl.SelectedTab == contactTabPage && contactBSource.Current != null;
        if (isAddressValid || isContactValid) { pasteToolStripMenuItem.Enabled = hasClipboardImage; }
        else { pasteToolStripMenuItem.Enabled = false; }
    }

    private async Task<TaskDialogButton> ProcessAndUploadGoogleContactPhotoAsync(Image rawImage, ImageFormat rawFormat, Contact targetContact, Image? previewImageForSync)
    {
        Image? workingImage = null;
        Image? finalImageToUpload = null;
        Image? finalImageForDisplay = null; // Wir behalten die Variable, falls ein Beschnitt stattfindet
        var finalDialogResult = TaskDialogButton.Cancel; // Standardmäßig Cancel
        try
        {
            // 1. Skalierung: Bei Google Kontakte max 250px Breite
            if (rawImage.Width > 250) { workingImage = Utils.SkaliereBildDaten(rawImage, 250); }
            else { workingImage = (Image)rawImage.Clone(); }

            var isNewContact = string.IsNullOrEmpty(targetContact.ResourceName);
            var buttonText = isNewContact ? "Übernehmen" : "Hochladen";
            var initialButtonYes = new TaskDialogButton(buttonText) { AllowCloseDialog = isNewContact };
            var initialButtonNo = TaskDialogButton.Cancel;

            var caveText = string.Empty;
            var radioButtons = new List<TaskDialogRadioButton>();
            TaskDialogRadioButton? centerRadio = null, topRadio = null, downRadio = null, skipRadio = null;

            // 2. Beschnitt-Logik konfigurieren (wenn Bild höher als breit)
            if (workingImage.Height > workingImage.Width && workingImage.Width > topAlignZoomPictureBox.Width)
            {
                topRadio = new TaskDialogRadioButton("&Oben priorisieren, nur unten abschneiden");
                centerRadio = new TaskDialogRadioButton("&Mitte priorisieren, oben/unten abschneiden") { Checked = true };
                downRadio = new TaskDialogRadioButton("&Unten priorisieren, nur oben abschneiden");
                skipRadio = new TaskDialogRadioButton("&Nicht beschneiden (nicht empfohlen)");
                radioButtons.AddRange([topRadio, centerRadio, downRadio, skipRadio]);

                caveText =
                    "\n\nDas Bild ist höher als breit. Es wird beim Download\n" +
                    "gemäß den Google-Vorgaben in einer auf 100 Pixel\n" +
                    "Höhe skalierten Version ausgegeben. Dies führt da-\n" +
                    "zu, dass das Foto den horizontal verfügbaren Platz\n" +
                    "nicht vollständig ausfüllen wird.\n\n" +
                    "Du kannst das hochzuladende Bild mit einer der\n" +
                    "folgenden Optionen zum Quadrat beschneiden:";
            }

            var replaceWarning = string.Empty;
            if (!string.IsNullOrEmpty(targetContact.PhotoUrl)) { replaceWarning += $"Das vorhandene Foto wird überschrieben!\n\n"; }

            var infoText = isNewContact
                ? $"Information: Abmessung {workingImage.Width}×{workingImage.Height} Pixel.\nDas Bild wird erst beim Speichern des Kontakts hochgeladen.{caveText}"
                : $"{replaceWarning}Information: Abmessung {workingImage.Width}×{workingImage.Height} Pixel.{caveText}";

            using TaskDialogIcon questionDialogIcon = new(Resources.question32);
            var initialPage = new TaskDialogPage()
            {
                Caption = "Google Kontakte",
                Heading = isNewContact ? "Foto übernehmen?" : "Möchtest du das Bild speichern?",
                Text = infoText,
                Icon = questionDialogIcon,
                AllowCancel = true,
                SizeToContent = true,
                Buttons = { initialButtonNo, initialButtonYes }
            };

            foreach (var rb in radioButtons) { initialPage.RadioButtons.Add(rb); }

            var inProgressCloseButton = TaskDialogButton.Close;
            inProgressCloseButton.Enabled = false;
            var progressPage = new TaskDialogPage()
            {
                Caption = "Google Kontakte", // oder AppName
                Heading = "Bitte warten…",
                Text = "Das Foto wird hochgeladen.",
                Icon = TaskDialogIcon.Information,
                ProgressBar = new TaskDialogProgressBar() { State = TaskDialogProgressBarState.Marquee },
                Buttons = { inProgressCloseButton }
            };

            // 3. User klickt auf Ja/Hochladen
            initialButtonYes.Click += (sender, ev) =>
            {
                // Wenn wir workingImage re-assignen, disposen wir das Zwischenergebnis
                void ReassignWorkingImage(Image newImage)
                {
                    var temp = workingImage;
                    workingImage = newImage;
                    temp?.Dispose();
                }

                // Beschnitt-Logik
                if (topRadio?.Checked == true)
                {
                    ReassignWorkingImage(Utils.BeschneideZuQuadrat(workingImage, null));
                    finalImageToUpload = workingImage;
                    finalImageForDisplay = (Image)workingImage.Clone();
                }
                else if (centerRadio?.Checked == true)
                {
                    ReassignWorkingImage(Utils.BeschneideZuQuadrat(workingImage, false));
                    finalImageToUpload = workingImage;
                    finalImageForDisplay = (Image)workingImage.Clone();
                }
                else if (downRadio?.Checked == true)
                {
                    ReassignWorkingImage(Utils.BeschneideZuQuadrat(workingImage, true));
                    finalImageToUpload = workingImage;
                    finalImageForDisplay = (Image)workingImage.Clone();
                }
                else if (skipRadio?.Checked == true)
                {
                    finalImageToUpload = workingImage;
                    // Nicht quadratisch beschneiden, aber wie Google skalieren
                    finalImageForDisplay = Utils.ReduziereWieGoogle(workingImage, 100);
                }
                else
                {
                    // Quadratisch oder breiter als hoch
                    finalImageToUpload = workingImage;
                    finalImageForDisplay = (Image)workingImage.Clone();
                }

                // PictureBox wird jetzt im Aufrufer synchronisiert 
                // Falls finalImageForDisplay ein anderes Bild ist als previewImageForSync (weil beschnitten wurde),
                // müssen wir die Anzeige aktualisieren.
                if (previewImageForSync != null && !finalImageForDisplay.Equals(previewImageForSync))
                {
                    topAlignZoomPictureBox.Image = finalImageForDisplay;
                    previewImageForSync.Dispose(); // Das ungeschnittene Vorschaubild wegwerfen
                }

                if (isNewContact) { delPictboxToolStripButton.Enabled = true; }
                else { initialPage.Navigate(progressPage); }
            };

            // 4. Der eigentliche Upload (unverändert)
            if (!isNewContact)
            {
                progressPage.Created += async (s, ev) =>
                {
                    try { await UpdateContactPhotoAsync(targetContact, finalImageToUpload!, rawFormat, () => progressPage.Buttons.First().PerformClick()); }
                    finally
                    {
                        workingImage?.Dispose();
                        // Das Bild, das wir für die Anzeige erstellt haben, disposen, 
                        // PictureBox hat es bereits (oder es wurde verworfen).
                        finalImageForDisplay?.Dispose();
                    }
                };
            }

            // Zeige den Dialog und speichere das Ergebnis
            finalDialogResult = TaskDialog.ShowDialog(Handle, initialPage);
        }
        catch (Exception ex)
        {
            Utils.MsgTaskDlg(Handle, $"Fehler bei der Bildverarbeitung: {ex.GetType()}", ex.Message, TaskDialogIcon.Error);
            finalImageForDisplay?.Dispose();
        }
        finally
        {
            // workingImage nur disposen, wenn kein Upload stattgefunden hat (Dialog abgebrochen oder Fehler)
            if (finalDialogResult == TaskDialogButton.Cancel)
            {
                workingImage?.Dispose();
                finalImageForDisplay?.Dispose();
            }
        }

        return finalDialogResult;
    }

    private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (topAlignZoomPictureBox.Image != null && delPictboxToolStripButton.Enabled)  // kopiere nur echtes Bild (Indikator: Löschen-Button ist aktiv)
        {
            try { Clipboard.SetImage(topAlignZoomPictureBox.Image); }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
        else { Console.Beep(); }
    }

    //private async void PasteToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    // 1. Sicherheitschecks
    //    if ((tabControl.SelectedTab == addressTabPage && addressBSource.Current == null) ||
    //        (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count == 0))
    //    {
    //        return;
    //    }

    //    // 2. Prüfen, ob die Zwischenablage Bilddaten enthält
    //    if (!Clipboard.ContainsImage())
    //    {
    //        Console.Beep();
    //        return;
    //    }

    //    // 3. Bild in den Speicher laden
    //    var clipboardImage = Clipboard.GetImage();
    //    if (clipboardImage == null) { return; }

    //    // ---------------------------------------------------------
    //    // FALL 1: Lokale Datenbank (EF Core)
    //    // ---------------------------------------------------------
    //    if (tabControl.SelectedTab == addressTabPage && addressBSource.Current is Adresse adresse)
    //    {
    //        try
    //        {
    //            topAlignZoomPictureBox.Image?.Dispose();
    //            topAlignZoomPictureBox.Image = null;

    //            Image finalImage;
    //            // Lokale Adressen: Wir skalieren auf 100 Pixel Breite
    //            if (clipboardImage.Width > 100)
    //            {
    //                finalImage = Utils.SkaliereBildDaten(clipboardImage, 100);
    //                clipboardImage.Dispose(); // GDI+ Objekt sauber freigeben!
    //            }
    //            else { finalImage = clipboardImage; }  // Bild direkt übernehmen

    //            topAlignZoomPictureBox.Image = finalImage;
    //            delPictboxToolStripButton.Enabled = true;

    //            // Bilddaten für DB vorbereiten (aus Clipboard immer als Jpeg sichern)
    //            byte[] datenZumSpeichern;
    //            using (var outputMs = new MemoryStream())
    //            {
    //                finalImage.Save(outputMs, ImageFormat.Jpeg);
    //                datenZumSpeichern = outputMs.ToArray();
    //            }

    //            adresse.Foto ??= new Foto();
    //            adresse.Foto.Fotodaten = datenZumSpeichern;

    //            addressBSource.ResetCurrentItem();
    //            UpdateSaveButton();
    //        }
    //        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
    //    }
    //    // ---------------------------------------------------------
    //    // FALL 2: Google Kontakte (aus Zwischenablage)
    //    // ---------------------------------------------------------
    //    else if (tabControl.SelectedTab == contactTabPage && contactBSource.Current is Contact currentContact)
    //    {
    //        // --- NEU: Backup des aktuellen Zustands (analog zum FileDialog) ---
    //        var previousImage = topAlignZoomPictureBox.Image;
    //        var previousDelButtonState = delPictboxToolStripButton.Enabled;

    //        try
    //        {
    //            // --- NEU: Sofortige Vorschau anzeigen ---
    //            Image? scaledPreviewImage = null;
    //            if (clipboardImage.Width > 250)
    //            {
    //                scaledPreviewImage = Utils.SkaliereBildDaten(clipboardImage, 250);
    //            }
    //            else
    //            {
    //                scaledPreviewImage = (Image)clipboardImage.Clone();
    //            }

    //            topAlignZoomPictureBox.Image = scaledPreviewImage;
    //            delPictboxToolStripButton.Enabled = true;
    //            topAlignZoomPictureBox.Refresh(); // Zwingt zum sofortigen Neuzeichnen vor dem Dialog

    //            // Standardformat für Clipboard-Bilder ist Jpeg.
    //            // Aufruf der angepassten Methode mit allen 4 Parametern!
    //            var dialogResult = await ProcessAndUploadGoogleContactPhotoAsync(clipboardImage, ImageFormat.Jpeg, currentContact, scaledPreviewImage);

    //            // Auswertung des Dialogs
    //            if (dialogResult == TaskDialogButton.Cancel)
    //            {
    //                // Benutzer hat abgebrochen: Altes Bild und Button-Zustand wiederherstellen
    //                topAlignZoomPictureBox.Image = previousImage;
    //                delPictboxToolStripButton.Enabled = previousDelButtonState;
    //                scaledPreviewImage.Dispose(); // Das verworfene Vorschaubild wegschmeißen
    //            }
    //            else
    //            {
    //                // Benutzer hat zugestimmt: Das alte Backup-Bild aus dem Speicher werfen
    //                previousImage?.Dispose();
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            // Fehlerfall: Altes Bild wiederherstellen
    //            topAlignZoomPictureBox.Image = previousImage;
    //            delPictboxToolStripButton.Enabled = previousDelButtonState;
    //            Utils.ErrTaskDlg(Handle, ex);
    //        }
    //        finally
    //        {
    //            // Das Originalbild aus der Zwischenablage wird am Ende immer sauber freigegeben
    //            clipboardImage?.Dispose();
    //        }
    //    }
    //}

    private async void PasteToolStripMenuItem_Click(object sender, EventArgs e)
    {
        // 1. Sicherheitschecks
        if ((tabControl.SelectedTab == addressTabPage && addressBSource.Current == null) ||
            (tabControl.SelectedTab == contactTabPage && contactDGV.SelectedRows.Count == 0))
        {
            return;
        }

        // 2. Prüfen, ob die Zwischenablage Bilddaten enthält
        if (!Clipboard.ContainsImage())
        {
            Console.Beep();
            return;
        }

        // 3. Bild in den Speicher laden
        var clipboardImage = Clipboard.GetImage();
        if (clipboardImage == null) { return; }

        // ---------------------------------------------------------
        // FALL 1: Lokale Datenbank (EF Core)
        // ---------------------------------------------------------
        if (tabControl.SelectedTab == addressTabPage && addressBSource.Current is Adresse adresse)
        {
            try
            {
                topAlignZoomPictureBox.Image?.Dispose();
                topAlignZoomPictureBox.Image = null;

                Image finalImage;
                // Lokale Adressen: Wir skalieren auf 100 Pixel Breite
                if (clipboardImage.Width > 100)
                {
                    finalImage = Utils.SkaliereBildDaten(clipboardImage, 100);
                    clipboardImage.Dispose(); // GDI+ Objekt sauber freigeben!
                }
                else { finalImage = clipboardImage; }  // Bild direkt übernehmen

                topAlignZoomPictureBox.Image = finalImage;
                delPictboxToolStripButton.Enabled = true;

                // Bilddaten für DB vorbereiten (aus Clipboard immer als Jpeg sichern)
                byte[] datenZumSpeichern;
                using (var outputMs = new MemoryStream())
                {
                    finalImage.Save(outputMs, ImageFormat.Jpeg);
                    datenZumSpeichern = outputMs.ToArray();
                }

                adresse.Foto ??= new Foto();
                adresse.Foto.Fotodaten = datenZumSpeichern;

                addressBSource.ResetCurrentItem();
                UpdateSaveButton();
            }
            catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        }
        // ---------------------------------------------------------
        // FALL 2: Google Kontakte (aus Zwischenablage)
        // ---------------------------------------------------------
        else if (tabControl.SelectedTab == contactTabPage && contactBSource.Current is Contact currentContact)
        {
            // --- DER FIX: Eine echte, unabhängige Kopie des alten Bildes anlegen ---
            Image? previousImage = null;
            if (topAlignZoomPictureBox.Image != null)
            {
                previousImage = new Bitmap(topAlignZoomPictureBox.Image);
            }
            var previousDelButtonState = delPictboxToolStripButton.Enabled;

            try
            {
                // Sofortige Vorschau anzeigen
                Image? scaledPreviewImage = null;
                if (clipboardImage.Width > 250)
                {
                    scaledPreviewImage = Utils.SkaliereBildDaten(clipboardImage, 250);
                }
                else
                {
                    scaledPreviewImage = (Image)clipboardImage.Clone();
                }

                // PictureBox übernimmt ab hier die Verantwortung für scaledPreviewImage
                topAlignZoomPictureBox.Image = scaledPreviewImage;
                delPictboxToolStripButton.Enabled = true;
                topAlignZoomPictureBox.Refresh(); // Zwingt zum sofortigen Neuzeichnen vor dem Dialog

                // Aufruf der angepassten Methode mit allen 4 Parametern!
                var dialogResult = await ProcessAndUploadGoogleContactPhotoAsync(clipboardImage, ImageFormat.Jpeg, currentContact, scaledPreviewImage);

                // Auswertung des Dialogs
                if (dialogResult == TaskDialogButton.Cancel)
                {
                    // Benutzer hat abgebrochen: Altes Bild und Button-Zustand wiederherstellen.
                    // Hinweis: Durch diese Zuweisung erledigt das Custom-Control das Dispose() 
                    // für das verworfene scaledPreviewImage ganz automatisch!
                    topAlignZoomPictureBox.Image = previousImage;
                    delPictboxToolStripButton.Enabled = previousDelButtonState;
                }
                else
                {
                    // Benutzer hat zugestimmt: Das temporäre Backup-Bild aus dem Speicher werfen
                    previousImage?.Dispose();
                }
            }
            catch (Exception ex)
            {
                // Fehlerfall: Altes Bild wiederherstellen
                topAlignZoomPictureBox.Image = previousImage;
                delPictboxToolStripButton.Enabled = previousDelButtonState;
                Utils.ErrTaskDlg(Handle, ex);
            }
            finally
            {
                // Das Originalbild aus der Zwischenablage immer sicher freigeben
                clipboardImage?.Dispose();
            }
        }
    }

    private async void AddPictboxToolStripButton_Click(object sender, EventArgs e)
    {
        // Sicherheitschecks
        if ((tabControl.SelectedTab == addressTabPage && addressBSource.Current == null) ||
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
            if (addressBSource.Current is Adresse adresse)
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

                    adresse.Foto ??= new Foto(); // Neue Foto-Entity anlegen, falls noch keine existiert
                    adresse.Foto.Fotodaten = datenZumSpeichern;

                    addressBSource.ResetCurrentItem();
                    UpdateSaveButton();

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
        // FALL 2: Google Kontakte (im FileDialog)
        // ---------------------------------------------------------
        else if (tabControl.SelectedTab == contactTabPage && contactBSource.Current is Contact currentContact)
        {
            // --- DER FIX: Eine echte, unabhängige Kopie des alten Bildes anlegen ---
            Image? previousImage = null;
            if (topAlignZoomPictureBox.Image != null)
            {
                // new Bitmap() erzeugt ein komplett neues GDI-Objekt im Arbeitsspeicher,
                // das überlebt, selbst wenn die PictureBox das Original abschießt.
                previousImage = new Bitmap(topAlignZoomPictureBox.Image);
            }

            var previousDelButtonState = delPictboxToolStripButton.Enabled;

            try
            {
                Image originalImage;
                ImageFormat origImgFormat;
                long fileLength;

                // 1. Bild sicher in den RAM laden und SOFORT vom FileStream entkoppeln
                using (var fs = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                {
                    fileLength = fs.Length;
                    using var tempImage = Image.FromStream(fs);
                    origImgFormat = tempImage.RawFormat;
                    originalImage = new Bitmap(tempImage);
                }

                // 2. originalImage am Ende der Verarbeitung sauber freigeben
                using (originalImage)
                {
                    Utils.WendeExifOrientierungAn(originalImage);

                    if (fileLength > 1024 * 1024)
                    {
                        Utils.MsgTaskDlg(Handle, "Automatische Größenreduzierung",
                            $"Die Dateigröße ist größer als 1 MB ({Utils.FormatBytes(fileLength)}).\nEs erfolgt eine Skalierung.",
                            TaskDialogIcon.ShieldWarningYellowBar);
                    }

                    Image? scaledPreviewImage = null;
                    if (originalImage.Width > 250)
                    {
                        scaledPreviewImage = Utils.SkaliereBildDaten(originalImage, 250);
                    }
                    else
                    {
                        scaledPreviewImage = (Image)originalImage.Clone();
                    }

                    // Vorschau anzeigen (Hier zerstört die PictureBox das alte Original, aber wir haben ja das Backup!)
                    topAlignZoomPictureBox.Image = scaledPreviewImage;
                    delPictboxToolStripButton.Enabled = true;
                    topAlignZoomPictureBox.Refresh();

                    var dialogResult = await ProcessAndUploadGoogleContactPhotoAsync(originalImage, origImgFormat, currentContact, scaledPreviewImage);

                    if (dialogResult == TaskDialogButton.Cancel)
                    {
                        // Abbruch: Unser sicheres Backup wiederherstellen
                        topAlignZoomPictureBox.Image = previousImage;
                        delPictboxToolStripButton.Enabled = previousDelButtonState;
                        //scaledPreviewImage.Dispose();  // das macht schohn das Custom-Control
                    }
                    else
                    {
                        // Erfolg: Backup wird nicht mehr gebraucht und gelöscht
                        previousImage?.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                // Fehlerfall: Backup wiederherstellen
                topAlignZoomPictureBox.Image = previousImage;
                delPictboxToolStripButton.Enabled = previousDelButtonState;
                Utils.ErrTaskDlg(Handle, ex);
            }
        }
    }

    private async void DelPictboxToolStripButton_Click(object sender, EventArgs e)
    {
        // --- FALL A: SQL ADRESSEN ---
        if (tabControl.SelectedTab == addressTabPage && addressBSource.Current is Adresse adresse)
        {
            var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Adressen", "Möchtest du das Bild entfernen?",
                                "Das Foto wird zum Löschen vorgemerkt.", "&Entfernen", "&Behalten", false);
            if (isYes)
            {
                try
                {
                    if (adresse.Foto != null)
                    {
                        _context?.Fotos.Remove(adresse.Foto);
                        // EF Core 10 Tipp: Wir setzen die Referenz explizit auf null
                        adresse.Foto = null;

                        topAlignZoomPictureBox.Image = Resources.AddressBild100;
                        delPictboxToolStripButton.Enabled = false;

                        addressBSource.ResetCurrentItem();
                        UpdateSaveButton();
                    }
                }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
        }
        // --- FALL B: GOOGLE KONTAKTE ---
        else if (tabControl.SelectedTab == contactTabPage && contactBSource.Current is Contact googleKontakt)
        {
            // PRÜFUNG: Ist es ein neuer Kontakt (noch nicht gespeichert)?
            if (string.IsNullOrEmpty(googleKontakt.ResourceName))
            {
                // Wenn der Kontakt neu ist, existiert das Bild nur lokal in der PictureBox.
                // Wir löschen es einfach ohne API-Call und ohne "Unwiderruflich"-Warnung.

                topAlignZoomPictureBox.Image = Resources.ContactBild100; // Standard-Icon zurücksetzen
                delPictboxToolStripButton.Enabled = false;

                // Ggf. SaveButton Status aktualisieren (da Änderung am ungespeicherten Objekt zurückgenommen wurde)
                UpdateSaveButton();
                return;
            }

            // Bestehender Kontakt: API Call nötig
            var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Google Kontakte", "Möchtest du das Bild wirklich löschen?",
                    "Das Foto wird bei Google unwiderruflich entfernt.", "&Löschen", "&Belassen", false);
            if (isYes)
            {
                try
                {
                    // WICHTIG: Wir übergeben das OBJEKT googleKontakt
                    await DeleteContactPhotoAsync(googleKontakt);
                    googleKontakt.PhotoUrl = null;

                    // UI-Update
                    topAlignZoomPictureBox.Image = Resources.ContactBild100; // Spezielles Kontakt-Icon
                    delPictboxToolStripButton.Enabled = false;

                    // Da das Foto weg ist, muss die Spalte im Grid ("alle mit Bild") aktualisiert werden
                    contactBSource.ResetCurrentItem();
                }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
        }
    }

    private async Task<bool> CopyLocalToGoogleAsync(Adresse localAddress)
    {
        Contact? createdContact = null;
        var newGoogleContact = new Contact();
        try
        {
            var typeLocal = typeof(Adresse);
            var typeGoogle = typeof(Contact);
            foreach (var fieldName in dataFields)
            {
                var propLocal = typeLocal.GetProperty(fieldName);
                var value = propLocal?.GetValue(localAddress);
                var propGoogle = typeGoogle.GetProperty(fieldName);
                if (propGoogle != null && propGoogle.CanWrite) { propGoogle.SetValue(newGoogleContact, value); }
            }

            // WICHTIG: Gruppen manuell mappen, da sie nicht in dataFields stehen!
            if (localAddress.Gruppen != null)
            {
                newGoogleContact.GroupNames = [.. localAddress.Gruppen.Select(g => g.Name)];
            }
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            return false;
        }

        var success = await Utils.RunWithProgressDialogAsync(this, "Google Upload", "Kontakt wird erstellt…", async token =>
        {
            Image? imageToUpload = null;
            try
            {
                if (localAddress.Foto?.Fotodaten != null && localAddress.Foto.Fotodaten.Length > 0)
                {
                    try
                    {
                        var ms = new MemoryStream(localAddress.Foto.Fotodaten);
                        imageToUpload = new Bitmap(ms);
                    }
                    catch { }
                }
                var manager = new GooglePeopleManager(secretPath, tokenDir);
                createdContact = await manager.CreateContactAsync(newGoogleContact, imageToUpload, token);
            }
            finally { imageToUpload?.Dispose(); }
        });

        if (success && createdContact != null)
        {
            try
            {
                _allGoogleContacts?.Add(createdContact);
                if (_allGoogleContacts != null)
                {
                    Utils.SortContacts(_allGoogleContacts);
                    contactBSource.DataSource = _allGoogleContacts;
                    contactBSource.ResetBindings(false);
                    var newIndex = _allGoogleContacts.IndexOf(createdContact);
                    if (newIndex >= 0) { contactBSource.Position = newIndex; }
                }
                else
                {
                    contactBSource.Add(createdContact);
                    contactBSource.Position = contactBSource.Count - 1;
                }
                _lastActiveContact = createdContact;
                _originalContactSnapshot = (Contact)createdContact.Clone();
                return true;
            }
            catch (Exception ex)
            {
                Utils.ErrTaskDlg(Handle, ex);
                return false;
            }
        }
        return false;
    }

    private async Task<bool> CopyGoogleToLocalAsync(Contact googleKontakt)
    {
        try
        {
            var newLocalAddress = new Adresse();

            // -----------------------------------------------------------
            // 1. Automatische Zuweisung mittels Reflection und dataFields
            // -----------------------------------------------------------
            var sourceType = typeof(Contact);
            var targetType = typeof(Adresse);

            foreach (var fieldName in dataFields)
            {
                var sourceProp = sourceType.GetProperty(fieldName);
                var targetProp = targetType.GetProperty(fieldName);

                if (sourceProp != null && targetProp != null && targetProp.CanWrite)
                {
                    var value = sourceProp.GetValue(googleKontakt);
                    targetProp.SetValue(newLocalAddress, value);
                }
            }

            // -----------------------------------------------------------
            // 2. Gruppen mappen (mit EF Core Logik aus dem Import)
            // -----------------------------------------------------------
            if (googleKontakt.GroupNames != null)
            {
                foreach (var gName in googleKontakt.GroupNames)
                {
                    // Leere Namen und den Favoriten-Stern ignorieren wir für die lokale DB
                    if (string.IsNullOrWhiteSpace(gName) || gName == "★") { continue; }

                    var gruppe = _context?.Gruppen.Local.FirstOrDefault(g => g.Name.Equals(gName, StringComparison.OrdinalIgnoreCase))
                                 ?? _context?.Gruppen.FirstOrDefault(g => g.Name.Equals(gName, StringComparison.CurrentCultureIgnoreCase));

                    if (gruppe == null)
                    {
                        gruppe = new Gruppe { Name = gName };
                        _context?.Gruppen.Add(gruppe);
                    }
                    newLocalAddress.Gruppen.Add(gruppe);
                }
            }

            // -----------------------------------------------------------
            // 3. Foto separat laden (Speziallogik, nicht im Array)
            // -----------------------------------------------------------
            if (!string.IsNullOrEmpty(googleKontakt.PhotoUrl))
            {
                try
                {
                    var bytes = await HttpService.Client.GetByteArrayAsync(googleKontakt.PhotoUrl);
                    newLocalAddress.Foto = new Foto { Fotodaten = bytes };
                }
                catch { /* Foto Fehler ignorieren, Rest wird trotzdem gespeichert */ }
            }

            // -----------------------------------------------------------
            // 4. UI Update & Sortierung
            // -----------------------------------------------------------
            var insertIndex = Utils.GetAddressInsertIndex(addressBSource, newLocalAddress);
            addressBSource.Insert(insertIndex, newLocalAddress);
            addressBSource.Position = insertIndex;

            return true;
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            return false;
        }
    }

    //private void UpdateMembershipTags()
    //{
    //    var isContactTab = tabControl.SelectedTab == contactTabPage;
    //    var groupsList = isContactTab ? curContactMemberships : curAddressMemberships;
    //    flowLayoutPanel.Controls.Clear();
    //    foreach (var membership in groupsList)
    //    {
    //        var tagControl = new TagControl
    //        {
    //            Text = membership,
    //            Membership = membership
    //        };

    //        tagControl.DeleteClick += (sender, e) =>
    //        {
    //            var ctrl = sender as TagControl;
    //            var membershipToRemove = ctrl?.Membership;
    //            if (string.IsNullOrEmpty(membershipToRemove)) { return; }

    //            if (isContactTab) // --- Google Kontakte Logic ---
    //            {
    //                curContactMemberships.Remove(membershipToRemove);
    //                UpdateMembershipTags();
    //                UpdateCurrentContactMemberships();
    //            }
    //            else
    //            {
    //                if (addressBSource.Current is Adresse adresse)
    //                {
    //                    var gruppeToDelete = adresse.Gruppen.FirstOrDefault(g => g.Name.Equals(membershipToRemove, StringComparison.OrdinalIgnoreCase));
    //                    if (gruppeToDelete != null)
    //                    {
    //                        // 1. Verknüpfung entfernen (Erzeugt "Deleted" State bei der Schatten-Entität)
    //                        adresse.Gruppen.Remove(gruppeToDelete);
    //                        curAddressMemberships.Remove(membershipToRemove);

    //                        // 2. UI Aktualisieren
    //                        UpdateMembershipTags();
    //                        UpdateTagComboBoxDataSource();
    //                        UpdatePlaceholderVis();

    //                        // 3. WICHTIG: UI benachrichtigen (aktiviert Buttons, feuert Events)
    //                        //addressBindingSource.ResetCurrentItem();

    //                        // 4. Save-Button explizit prüfen
    //                        UpdateSaveButton();
    //                    }
    //                }
    //            }
    //        };
    //        flowLayoutPanel.Controls.Add(tagControl);
    //    }
    //    UpdatePlaceholderVis();
    //}

    private void UpdateMembershipTags()
    {
        var isContactTab = tabControl.SelectedTab == contactTabPage;
        var targetTags = isContactTab ? curContactMemberships : curAddressMemberships;

        // 1. Aktuell angezeigte Tags auslesen (Wir casten auf TagControl, um die Membership-Property zu lesen)
        var currentTags = flowLayoutPanel.Controls.Cast<Control>()
            .Where(c => c.Name != "lblPlaceholder")
            .Select(c => c is TagControl tc ? tc.Membership : c.Text) // Sicherer Fallback auf Text
            .ToList();

        // 2. Der Flimmer-Killer: Sind die Tags exakt gleich? Dann abbrechen!
        if (targetTags.SequenceEqual(currentTags))
        {
            return;
        }

        // 3. Es gibt Änderungen. Wir pausieren das Layout.
        flowLayoutPanel.SuspendLayout();

        // 4. Alte Controls SAUBER entfernen (Verhindert das Memory Leak!)
        while (flowLayoutPanel.Controls.Count > 0)
        {
            var ctrl = flowLayoutPanel.Controls[0];
            flowLayoutPanel.Controls.RemoveAt(0);
            ctrl.Dispose(); // Windows-Handles an das System zurückgeben
        }

        // 5. Neue Controls oder Placeholder hinzufügen
        if (targetTags.Count == 0)
        {
            UpdatePlaceholderVis();
        }
        else
        {
            foreach (var membership in targetTags)
            {
                var tagControl = new TagControl
                {
                    Text = membership,
                    Membership = membership
                };

                // Deine originale Lösch-Logik
                tagControl.DeleteClick += (sender, e) =>
                {
                    var ctrl = sender as TagControl;
                    var membershipToRemove = ctrl?.Membership;
                    if (string.IsNullOrEmpty(membershipToRemove)) { return; }

                    if (isContactTab) // --- Google Kontakte Logic ---
                    {
                        curContactMemberships.Remove(membershipToRemove);
                        UpdateMembershipTags();
                        UpdateCurrentContactMemberships();
                    }
                    else
                    {
                        if (addressBSource.Current is Adresse adresse)
                        {
                            var gruppeToDelete = adresse.Gruppen.FirstOrDefault(g => g.Name.Equals(membershipToRemove, StringComparison.OrdinalIgnoreCase));
                            if (gruppeToDelete != null)
                            {
                                // 1. Verknüpfung entfernen (Erzeugt "Deleted" State bei der Schatten-Entität)
                                adresse.Gruppen.Remove(gruppeToDelete);
                                curAddressMemberships.Remove(membershipToRemove);

                                // 2. UI Aktualisieren
                                UpdateMembershipTags();
                                UpdateTagComboBoxDataSource();

                                // (UpdatePlaceholderVis wird jetzt automatisch von UpdateMembershipTags aufgerufen, wenn leer!)

                                // 3. Save-Button explizit prüfen
                                UpdateSaveButton();
                            }
                        }
                    }
                };

                flowLayoutPanel.Controls.Add(tagControl);
            }
        }

        // 6. Alles auf einmal zeichnen lassen
        flowLayoutPanel.ResumeLayout(true);
    }

    private void TagButton_Click(object sender, EventArgs e)
    {
        var newMembershipName = tagComboBox.Text.Trim();
        if (string.IsNullOrEmpty(newMembershipName)) { return; }
        if (newMembershipName == "*") { newMembershipName = "★"; }

        if (tabControl.SelectedTab == contactTabPage)
        {
            if (curContactMemberships.Contains(newMembershipName)) { return; }
            curContactMemberships.Add(newMembershipName);
            allContactMemberships.Add(newMembershipName);

            UpdateMembershipTags();
            UpdateTagComboBoxDataSource();
            UpdateCurrentContactMemberships(); // Google nutzt weiterhin JSON/Strings
        }
        else if (tabControl.SelectedTab == addressTabPage)
        {
            if (addressBSource.Current is Adresse adresse && _context != null) // _context Prüfung hier integriert
            {
                var entry = _context.Entry(adresse);  // EF Core EntityEntry für die aktuelle Adresse holen
                if (!entry.Collection(a => a.Gruppen).IsLoaded)
                {
                    entry.Collection(a => a.Gruppen).Load();
                    LoadGroupsForCurrentAddress();
                }

                if (adresse.Gruppen.Any(g => g.Name.Equals(newMembershipName, StringComparison.OrdinalIgnoreCase)))
                {
                    tagComboBox.SelectAll();
                    tagComboBox.Focus();
                    return; // Hier brechen wir ab - aber jetzt ist die UI bereits aktuell!
                }
                // A) Zuerst im ChangeTracker (Lokal) schauen
                var gruppe = _context?.Gruppen.Local.FirstOrDefault(g => g.Name.Equals(newMembershipName, StringComparison.OrdinalIgnoreCase));
                // B) Wenn nicht lokal, dann in der Datenbank suchen
                gruppe ??= _context?.Gruppen.FirstOrDefault(g => g.Name == newMembershipName);
                if (gruppe == null)
                {
                    gruppe = new Gruppe { Name = newMembershipName };
                    _context?.Gruppen.Add(gruppe);
                    allAddressMemberships.Add(newMembershipName);  // Zur BindingList hinzufügen, damit die ComboBox es sofort kennt
                }
                adresse.Gruppen.Add(gruppe);  // Verknüpfung herstellen
                curAddressMemberships.Add(newMembershipName);
                //_context?.Entry(adresse).State = EntityState.Modified;  // Adresse als modifiziert markieren
                UpdateMembershipTags();
                UpdateTagComboBoxDataSource();
                addressBSource.ResetCurrentItem();
                UpdateSaveButton();
            }
        }
    }

    private void UpdateCurrentContactMemberships()
    {
        if (tabControl.SelectedTab == contactTabPage)
        {
            if (contactBSource.Current is Contact contact) { contact.GroupNames = [.. curContactMemberships]; }
        }
    }

    //private void UpdateTagComboBoxDataSource()
    //{
    //    var isContactTab = tabControl.SelectedTab == contactTabPage;
    //    string[] list = isContactTab  // Nur die Gruppen holen, die der aktuelle Datensatz noch NICHT hat
    //        ? [.. allContactMemberships.Except(curContactMemberships, StringComparer.OrdinalIgnoreCase)]
    //        : [.. allAddressMemberships.Except(curAddressMemberships, StringComparer.OrdinalIgnoreCase)];
    //    tagComboBox.DataSource = null;   // Keine DataSource! Wir füllen die Items direkt.
    //    tagComboBox.Items.Clear();
    //    if (list.Length > 0) { tagComboBox.Items.AddRange(list); }
    //    tagComboBox.Text = string.Empty;
    //    tagComboBox.AutoCompleteCustomSource ??= new AutoCompleteStringCollection();
    //    tagComboBox.AutoCompleteCustomSource.Clear();
    //    if (list.Length > 0) { tagComboBox.AutoCompleteCustomSource.AddRange(list); }
    //    tagComboBox.AutoCompleteMode = AutoCompleteMode.Append;
    //    tagComboBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
    //}

    private void UpdateTagComboBoxDataSource()
    {
        var isContactTab = tabControl.SelectedTab == contactTabPage;
        string[] newList = isContactTab
            ? [.. allContactMemberships.Except(curContactMemberships, StringComparer.OrdinalIgnoreCase)]
            : [.. allAddressMemberships.Except(curAddressMemberships, StringComparer.OrdinalIgnoreCase)];

        // 1. Prüfen, ob sich die Items überhaupt geändert haben
        var currentItems = tagComboBox.Items.Cast<string>().ToArray();
        if (newList.SequenceEqual(currentItems))
        {
            // Keine Änderung an der Liste nötig. Wir leeren nur den Text.
            if (tagComboBox.Text != string.Empty)
            {
                tagComboBox.Text = string.Empty;
            }
            return;
        }

        // 2. Es gibt Änderungen! Wir nutzen den WM_SETREDRAW Trick, um das Flackern zu stoppen.
        const int WM_SETREDRAW = 0x000B;

        // Zeichnen verbieten
        NativeMethods.SendMessage(tagComboBox.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);

        try
        {
            // 3. Items austauschen
            tagComboBox.BeginUpdate(); // BeginUpdate ist wichtig für die Dropdown-Items
            tagComboBox.Items.Clear();
            if (newList.Length > 0)
            {
                tagComboBox.Items.AddRange(newList);
            }
            tagComboBox.EndUpdate();

            // 4. AutoComplete austauschen (Atomar, wie bei der Grussformel)
            tagComboBox.AutoCompleteMode = AutoCompleteMode.None; // Kurz ausschalten, um COM-Fehler zu vermeiden

            var newAutoComplete = new AutoCompleteStringCollection();
            if (newList.Length > 0)
            {
                newAutoComplete.AddRange(newList);
            }
            tagComboBox.AutoCompleteCustomSource = newAutoComplete;

            tagComboBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
            tagComboBox.AutoCompleteMode = AutoCompleteMode.Append;

            // 5. Textfeld leeren
            tagComboBox.Text = string.Empty;
        }
        finally
        {
            // Zeichnen wieder erlauben und EINMAL neu zeichnen
            NativeMethods.SendMessage(tagComboBox.Handle, WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
            tagComboBox.Invalidate();
        }
    }

    private void UpdatePlaceholderVis()
    {
        if (flowLayoutPanel.Controls.Count == 0)
        {
            var lblPlaceholder = new Label
            {
                Text = "Label",
                AutoSize = true,
                ForeColor = Color.Gray,
                BackColor = Color.Transparent,
                Name = "lblPlaceholder",
                Location = new Point(0, 0)
            };
            flowLayoutPanel.Controls.Add(lblPlaceholder);
        }
    }

    private void TagComboBox_Enter(object sender, EventArgs e)
    {
        tagComboBox.BackColor = _isDarkMode ? Color.FromArgb(80, 80, 0) : Color.LightYellow;
        tagComboBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
    }

    private void TagComboBox_Leave(object sender, EventArgs e)
    {
        tagComboBox.BackColor = _isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
        tagComboBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
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
        // Wenn die Standard-Liste offen ist und der User eine echte Taste drückt (Buchstaben/Zahlen),
        // schließen wir die Liste, BEVOR der Buchstabe getippt wird.
        if (tagComboBox.DroppedDown && e.KeyCode >= Keys.A && e.KeyCode <= Keys.Z) { tagComboBox.DroppedDown = false; }
        if (e.KeyCode == Keys.Enter)
        {
            if (tagButton.Enabled) { TagButton_Click(tagButton, EventArgs.Empty); }
            else { tbNotizen.Focus(); }
            e.SuppressKeyPress = true;
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
            dialogGroups = new SortedSet<string>(_context.Gruppen.Local.Select(g => g.Name), StringComparer.OrdinalIgnoreCase);
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
                addressBSource,
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
                contactBSource,
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
                var activeAddressAffected = false;

                foreach (var kvp in changes)
                {
                    var oldName = kvp.Key;
                    var newName = kvp.Value;

                    if (oldName == "★") { continue; }  // Favoriten schützen

                    var groupEntity = _context.Gruppen.Local.FirstOrDefault(g => g.Name.Equals(oldName, StringComparison.OrdinalIgnoreCase));

                    if (groupEntity == null) { continue; }
                    if (string.IsNullOrWhiteSpace(newName))
                    {
                        _context.Gruppen.Remove(groupEntity); // 1. Aus dem EF ChangeTracker entfernen
                        allAddressMemberships.Remove(oldName); // 2. Lokale UI-Listen aktualisieren
                        var curMembershipToRemove = curAddressMemberships.FirstOrDefault(g => g.Equals(oldName, StringComparison.OrdinalIgnoreCase));
                        if (curMembershipToRemove != null)
                        {
                            curAddressMemberships.Remove(curMembershipToRemove);
                            activeAddressAffected = true;
                        }
                        needsSave = true;
                    }
                    else
                    {
                        groupEntity.Name = newName; // 1. In der Entität umbenennen (EF Core merkt sich das)
                        allAddressMemberships.Remove(oldName); // 2. Lokale UI-Listen aktualisieren
                        allAddressMemberships.Add(newName);
                        var curMembershipToRename = curAddressMemberships.FirstOrDefault(g => g.Equals(oldName, StringComparison.OrdinalIgnoreCase));
                        if (curMembershipToRename != null)
                        {
                            curAddressMemberships.Remove(curMembershipToRename);
                            curAddressMemberships.Add(newName);
                            activeAddressAffected = true;
                        }
                        needsSave = true;
                    }
                }
                if (needsSave)
                {
                    UpdateSaveButton();  // await SaveSQLDatabaseAsync(); // Änderungen in die Datenbank schreiben
                    addressBSource.ResetBindings(false); // UI über Änderungen informieren
                    UpdateTagComboBoxDataSource();
                    if (activeAddressAffected) { UpdateMembershipTags(); } // Tag-Panel neu zeichnen, falls die aktuelle Adresse betroffen war
                }
            }
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            var groupDict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var gName in allContactMemberships) { groupDict[gName] = 0; } // 1. Zuerst alle bekannten Gruppen mit 0 initialisieren, damit auch leere Gruppen auftauchen
            if (_allGoogleContacts != null)  // 2. Jetzt die Kontakte durchgehen und die Zähler erhöhen
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
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                // Hier wird der Dialog direkt wie in deinen anderen Methoden aufgerufen
                await Utils.RunWithProgressDialogAsync(this, "Google-Gruppen", "Änderungen werden synchronisiert…", async token =>
                {
                    await ProcessGoogleGroupChangesAsync(frm.groupNameMap, token);
                });
            }
        }
    }

    private async Task ProcessGoogleGroupChangesAsync(Dictionary<string, string> groupChanges, CancellationToken token)
    {
        var realChanges = groupChanges.Where(kvp => kvp.Key != kvp.Value || string.IsNullOrEmpty(kvp.Value)).ToDictionary(k => k.Key, k => k.Value);

        if (realChanges.Count == 0)
        {
            return;
        }

        var contactsNeedRefresh = false;
        var activeContactAffected = false;

        foreach (var kvp in realChanges)
        {
            token.ThrowIfCancellationRequested(); // 1. Token prüfen, bevor der nächste API-Call oder die Verarbeitung startet

            var oldName = kvp.Key;
            var newName = kvp.Value;

            var resourceEntry = contactGroupsDict.FirstOrDefault(x => x.Value.Equals(oldName, StringComparison.OrdinalIgnoreCase));
            var resourceName = resourceEntry.Key;

            if (string.IsNullOrEmpty(resourceName))
            {
                continue;
            }

            var manager = new GooglePeopleManager(secretPath, tokenDir);

            if (string.IsNullOrEmpty(newName))
            {
                // 1. Bei Google löschen (Token durchreichen)
                await manager.DeleteContactGroupAsync(resourceName, token);

                // 2. Lokale Dictionarys & alle Gruppen aktualisieren
                contactGroupsDict.Remove(resourceName);
                allContactMemberships.Remove(oldName);

                // 3. Wenn die Gruppe beim aktuellen Kontakt im UI zu sehen ist: Entfernen
                var curMembershipToRemove = curContactMemberships.FirstOrDefault(g => g.Equals(oldName, StringComparison.OrdinalIgnoreCase));

                if (curMembershipToRemove != null)
                {
                    curContactMemberships.Remove(curMembershipToRemove);
                    activeContactAffected = true;
                }

                // 4. Aus allen geladenen Kontakten entfernen (Inklusive RawGooglePerson!)
                foreach (var contact in _allGoogleContacts)
                {
                    var hasChanged = false;

                    // A) String-Liste bereinigen
                    if (contact.GroupNames.Contains(oldName, StringComparer.OrdinalIgnoreCase))
                    {
                        var updatedGroups = contact.GroupNames.ToList();
                        updatedGroups.RemoveAll(g => g.Equals(oldName, StringComparison.OrdinalIgnoreCase));
                        contact.GroupNames = [.. updatedGroups];
                        hasChanged = true;
                    }

                    // B) RawGooglePerson Memberships bereinigen (WICHTIG!)
                    if (contact.RawGooglePerson?.Memberships != null)
                    {
                        var membershipToRemove = contact.RawGooglePerson.Memberships
                            .FirstOrDefault(m => m.ContactGroupMembership?.ContactGroupResourceName == resourceName);

                        if (membershipToRemove != null)
                        {
                            contact.RawGooglePerson.Memberships.Remove(membershipToRemove);
                            hasChanged = true;
                        }
                    }

                    if (hasChanged) { contactsNeedRefresh = true; }
                }

                // 5. Snapshot bereinigen (verhindert falschen Speicher-Dialog und API-Fehler)
                if (_originalContactSnapshot != null)
                {
                    // A) String-Liste bereinigen
                    if (_originalContactSnapshot.GroupNames.Contains(oldName, StringComparer.OrdinalIgnoreCase))
                    {
                        var snapGroups = _originalContactSnapshot.GroupNames.ToList();
                        snapGroups.RemoveAll(g => g.Equals(oldName, StringComparison.OrdinalIgnoreCase));
                        _originalContactSnapshot.GroupNames = [.. snapGroups];
                    }

                    // B) RawGooglePerson Memberships bereinigen (WICHTIG!)
                    if (_originalContactSnapshot.RawGooglePerson?.Memberships != null)
                    {
                        var membershipToRemove = _originalContactSnapshot.RawGooglePerson.Memberships
                            .FirstOrDefault(m => m.ContactGroupMembership?.ContactGroupResourceName == resourceName);

                        if (membershipToRemove != null)
                        {
                            _originalContactSnapshot.RawGooglePerson.Memberships.Remove(membershipToRemove);
                        }
                    }
                }
            }
        }

        // UI-Updates durchführen
        if (contactsNeedRefresh)
        {
            contactBSource.ResetBindings(false);
        }

        UpdateTagComboBoxDataSource();

        if (activeContactAffected)
        {
            UpdateMembershipTags();
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
        if (m.Msg == NativeMethods.WM_TRAY_RESTORE)  // Wenn AutoHotkey unseren geheimen Weckruf sendet
        {
            if (!Visible) { RestoreFromTray(); }
            return; // Nachricht exklusiv verarbeitet, Basisklasse wird komplett übersprungen!
        }
        else if (m.Msg == NativeMethods.WM_TRAY_MINIMIZE)
        {
            var activeDialog = OwnedForms.FirstOrDefault(f => f.Visible);  // Das erste sichtbare Kind-Fenster (Dialog) suchen
            if (activeDialog != null)
            {

                System.Media.SystemSounds.Exclamation.Play();  // Den nativen Windows-Hinweiston abspielen (blockiert die GUI nicht!)
                activeDialog.Activate();  // Da AHK das Hauptfenster aktiviert hat, holen wir den Dialog zurück in den Vordergrund
                return; // Minimieren komplett abbrechen!
            }
            if (min2TrayTSButton.Visible) { HideToTray(); }  // Kein Dialog offen -> Normale Logik anwenden
            else { WindowState = FormWindowState.Minimized; }
            return;
        }
        else if (m.Msg == NativeMethods.WM_SETTINGCHANGE)  // 2. Bei System-Nachrichten nur "mithören"
        {
            var area = Marshal.PtrToStringUni(m.LParam);
            if (string.IsNullOrEmpty(area) || area == "ImmersiveColorSet")
            {
                Application.SetColorMode(SystemColorMode.System);  // .NET 10 Dark Mode nativ neu evaluieren
                UpdateAppearanceStatus();
                Refresh();
                ToolStripManager.VisualStylesEnabled = true;
            }
        }
        base.WndProc(ref m);  // Alles, was oben nicht mit 'return' abgebrochen wurde, wird nun sauber vom Standard-Windows-System verarbeitet.
    }

    private void UpdateAppearanceStatus()
    {
        _isDarkMode = Application.SystemColorMode == SystemColorMode.Dark;
        if (Application.SystemColorMode == SystemColorMode.System) { _isDarkMode = DefaultBackColor.R < 128; } //falls die Automatik hakt
        ConfigureDgvAppearance(addressDGV, Color.FromArgb(176, 125, 71)); // Dein Braun
        ConfigureDgvAppearance(contactDGV, Color.FromArgb(0, 102, 204));  // Blau (z.B. Windows Default Blue)
        foreach (var c in Utils.GetAllControls(this))
        {
            if (c is PaddedTextBox || c is PaddedMaskedTextBox)  //  || c is ComboBox
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

    //private async void ContactDGV_RowValidating(object sender, DataGridViewCellCancelEventArgs e) => await OldAskSaveContactChangesAsync();
    private async void ContactDGV_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
    {
        // WICHTIG: Wenn wir gerade programmgesteuert (z.B. über "Neu"-Button)
        // die Zeile wechseln, darf dieses Event nicht dazwischenfunken.
        if (isSelectionChanging) { return; }

        // Wenn der User selbst klickt/navigiert:
        // Falls Check false zurückgibt (Abbruch), brechen wir den Zeilenwechsel ab.
        if (!await ContactChanges_Check())
        {
            e.Cancel = true;
        }
    }

    //private void AddressDGV_SelectionChanged(object sender, EventArgs e) => scrollTimer.Start();
    private void AddressDGV_SelectionChanged(object sender, EventArgs e)
    {
        if (addressDGV.FirstDisplayedScrollingRowIndex >= 0) { _savedAddressScrollIndex = addressDGV.FirstDisplayedScrollingRowIndex; }
    }

    private void ContactDGV_DataError(object sender, DataGridViewDataErrorEventArgs e)
    {
        if (e.Exception is IndexOutOfRangeException || e.Exception is ArgumentException)
        {
            e.Cancel = true;
            e.ThrowException = false;
        }
    }

    private async void UpdateCheckToolStripMenuItem_Click(object sender, EventArgs e)
    {
        // 1. RadioButtons vorbereiten
        var rbn0 = new TaskDialogRadioButton("Jeden Tag");
        var rbn1 = new TaskDialogRadioButton("Jede Woche");
        var rbn2 = new TaskDialogRadioButton("Jeden Monat");
        var rbn3 = new TaskDialogRadioButton("Niemals");

        // 2. Beide Pages ZUERST deklarieren, damit sie gegenseitig bekannt sind
        var progressPage = new TaskDialogPage();
        var updatePage = new TaskDialogPage();

        var cts = new CancellationTokenSource();
        // 3. Den "Überspringen"-Button konfigurieren
        var btnSkip = new TaskDialogButton("Überspringen")
        {
            AllowCloseDialog = false // Verhindert das Schließen des Dialogs
        };

        // Flag um doppelte Navigation zu verhindern (User klickt Skip UND Task wird fertig)
        var hasNavigated = false;

        btnSkip.Click += (s, args) =>
        {
            cts.Cancel(); // sofort den Download abbrechen
            // Wir nutzen hier direkt die Variable 'progressPage' aus dem Scope (Closure), 
            // statt zu versuchen sie aus 'args' zu lesen.
            if (!hasNavigated && progressPage.BoundDialog != null)
            {
                hasNavigated = true;
                progressPage.Navigate(updatePage);
            }
        };

        // 4. ProgressPage Eigenschaften setzen
        progressPage.Caption = appName;
        progressPage.Heading = "Update-Prüfung";
        progressPage.Text = "Suche nach Updates...";
        progressPage.ProgressBar = new TaskDialogProgressBar(TaskDialogProgressBarState.Marquee);
        progressPage.SizeToContent = true;
        progressPage.AllowCancel = true;
        progressPage.Buttons.Add(btnSkip); // Button hinzufügen

        // 5. UpdatePage Eigenschaften setzen
        updatePage.Caption = appName;
        updatePage.Heading = "Automatische Updatesuche";
        updatePage.Text = "Wie häufig soll nach einem Update gesucht werden?";
        updatePage.Buttons.Add(TaskDialogButton.OK);
        updatePage.Buttons.Add(TaskDialogButton.Cancel);
        updatePage.AllowCancel = true;
        updatePage.SizeToContent = true;

        updatePage.RadioButtons.Add(rbn0);
        updatePage.RadioButtons.Add(rbn1);
        updatePage.RadioButtons.Add(rbn2);
        updatePage.RadioButtons.Add(rbn3);

        // Initialisierung der Settings (RadioButtons auswählen)
        if (_settings.UpdateIndex == 1) { rbn1.Checked = true; }
        else if (_settings.UpdateIndex == 2) { rbn2.Checked = true; }
        else if (_settings.UpdateIndex == 3) { rbn3.Checked = true; }
        else { rbn0.Checked = true; }

        // 6. Die asynchrone Logik
        progressPage.Created += async (s, args) =>
        {
            try
            {
                // Version abrufen
                var (latestVersion, releaseDate) = await Utils.GetLatestVersionInfoAsync();

                // Wenn wir hier ankommen, wurde NICHT abgebrochen.
                // Trotzdem zur Sicherheit prüfen (falls Cancel genau zwischen await und hier passierte)
                if (hasNavigated || cts.IsCancellationRequested) { return; }

                RefreshUpdateUI(latestVersion, releaseDate);

                var footText = "";
                if (latestVersion != null)
                {
                    var currentVersion = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0);

                    // Formatierung für Fußnote
                    if (latestVersion > currentVersion) { footText = $"Update verfügbar: v{latestVersion} vom {releaseDate}\nBeachte den Download-Button in der Statuszeile rechts unten!"; }
                    else { footText = $"Status: Aktuell\nInstalliert: {currentVersion.ToString(3)}\nVerfügbar: {latestVersion}\nDatum: {releaseDate}"; }
                }
                else { footText = "Der Update-Server konnte nicht erreicht werden."; }
                updatePage.Footnote = new TaskDialogFootnote(footText);

                // Navigation zur UpdatePage, falls noch nicht geschehen
                if (!hasNavigated && progressPage.BoundDialog != null)
                {
                    hasNavigated = true;
                    progressPage.Navigate(updatePage);
                }
            }
            catch (OperationCanceledException) { }  // Alles gut, der User wollte abbrechen. Nichts tun.
        };

        // 7. Dialog anzeigen
        // Da wir zur updatePage navigieren, kommt das Resultat von dort (OK oder Cancel)
        var resultButton = TaskDialog.ShowDialog(this, progressPage);

        if (resultButton == TaskDialogButton.OK)
        {
            var newIndex = rbn1.Checked ? 1 : rbn2.Checked ? 2 : rbn3.Checked ? 3 : 0;
            _settings.UpdateIndex = newIndex;
            SettingsManager.Save(_settings, _settingsPath);
        }
    }

    private void RefreshUpdateUI(Version? latestVersion, string? releaseDate)
    {
        var currentVersion = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0);

        if (latestVersion != null)
        {
            if (latestVersion > currentVersion)
            {
                btnUpdateAvailable.Visible = true;
                btnUpdateAvailable.ToolTipText = $"Update verfügbar: v{latestVersion} vom {releaseDate}";
            }
            else
            {
                btnUpdateAvailable.Visible = false;
                _settings.LastUpdateCheck = DateTime.Now;  // aktualisieren nur wenn kein Update verfügbar ist
                SettingsManager.Save(_settings, _settingsPath);
            }
        }
        else
        {
            // Fehlerfall: Update-Prüfung deaktivieren, um ständige Fehlversuche zu vermeiden
            _settings.UpdateIndex = 3;
            SettingsManager.Save(_settings, _settingsPath);
        }
    }

    private void BtnUpdateAvailable_ButtonClick(object sender, EventArgs e)
    {
        var url = "https://www.netradio.info/address/";  //var url = btnUpdateAvailable.Tag?.ToString();
        if (!string.IsNullOrEmpty(url)) { Utils.StartLink(Handle, url); }
    }

    private void TbNotizen_KeyDown(object sender, KeyEventArgs e)
    {  // Strg+F wird in ProcessCmdKey behandelt
        if (e.KeyCode == Keys.F3)  // Weitersuchen
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            if (tbNotizen.TextLength == 0) { return; }
            if (!string.IsNullOrEmpty(_lastSearchTerm)) { FindInTextBox(_lastSearchTerm, _lastMatchCase); }
            else { OpenSearchDialog(); }
        }
    }

    private void OpenSearchDialog()
    {
        if (tbNotizen.TextLength == 0) { return; }
        var selectedText = tbNotizen.SelectedText;
        var initialSearch = string.IsNullOrEmpty(selectedText) ? _lastSearchTerm : selectedText;
        using var searchDialog = new TextBoxSearch(initialSearch, _lastMatchCase);
        if (searchDialog.ShowDialog(this) == DialogResult.OK)
        {
            _lastSearchTerm = searchDialog.SearchText;
            _lastMatchCase = searchDialog.MatchCase;
            Utils.AddToSearchHistory(_lastSearchTerm);  // Historie aktualisieren (auf z.B. 10 Einträge begrenzt, Neueste oben)
            FindInTextBox(_lastSearchTerm, _lastMatchCase);
        }
    }

    private void FindInTextBox(string searchText, bool matchCase)
    {
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var startIndex = tbNotizen.SelectionStart + tbNotizen.SelectionLength;  // Suche ab der aktuellen Cursorposition oder nach der letzten Markierung starten
        if (startIndex >= tbNotizen.TextLength) { startIndex = 0; }
        var index = tbNotizen.Text.IndexOf(searchText, startIndex, comparison);
        if (index != -1)
        {
            tbNotizen.Select(index, searchText.Length);  // Text gefunden -> Markieren und ins Bild scrollen
            tbNotizen.ScrollToCaret();
            tbNotizen.Focus();
        }
        else
        {
            if (startIndex > 0)  // Falls wir nicht am Anfang gestartet sind, Suche von oben wiederholen
            {
                index = tbNotizen.Text.IndexOf(searchText, 0, comparison);
                if (index != -1)
                {
                    tbNotizen.Select(index, searchText.Length);
                    tbNotizen.ScrollToCaret();
                    tbNotizen.Focus();
                    return;
                }
            }
            Utils.MsgTaskDlg(Handle, "Suche in Notizen", "Der Suchtext wurde nicht gefunden.", TaskDialogIcon.Information);
        }
    }

    private void AddressDGV_Scroll(object sender, ScrollEventArgs e)
    {
        if (e.ScrollOrientation == ScrollOrientation.VerticalScroll) { _savedAddressScrollIndex = addressDGV.FirstDisplayedScrollingRowIndex; } // Position speichern, wenn der User scrollt
    }

    private void AddressDGV_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
    {
        if (_savedAddressScrollIndex >= 0 && _savedAddressScrollIndex < addressDGV.RowCount)
        {
            _ = addressDGV.InvokeAsync(() =>  // InvokeAsync stellt sicher, dass das Layout-Rendering abgeschlossen ist, bevor wir den Scrollbalken verschieben.
            {
                try
                {
                    if (_savedAddressScrollIndex < addressDGV.RowCount) { addressDGV.FirstDisplayedScrollingRowIndex = _savedAddressScrollIndex; }
                }
                catch { } // Stille Korrektur, falls die Ansicht z.B. gefiltert wurde
            });
        }
    }

    private void ContactDGV_Scroll(object sender, ScrollEventArgs e)
    {
        if (e.ScrollOrientation == ScrollOrientation.VerticalScroll) { _savedContactScrollIndex = contactDGV.FirstDisplayedScrollingRowIndex; }
    }

    private void ContactDGV_SelectionChanged(object sender, EventArgs e)
    {
        if (contactDGV.FirstDisplayedScrollingRowIndex >= 0)
        {
            _savedContactScrollIndex = contactDGV.FirstDisplayedScrollingRowIndex;
        }
    }

    private void ContactDGV_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
    {
        if (_savedContactScrollIndex >= 0 && _savedContactScrollIndex < contactDGV.RowCount)
        {
            _ = contactDGV.InvokeAsync(() =>
            {
                try
                {
                    if (_savedContactScrollIndex < contactDGV.RowCount) { contactDGV.FirstDisplayedScrollingRowIndex = _savedContactScrollIndex; }
                }
                catch { }
            });
        }
    }

    private void TopAlignZoomPictureBox_MouseDoubleClick(object sender, EventArgs e) => AddPictboxToolStripButton_Click(topAlignZoomPictureBox, EventArgs.Empty);

    //private void FindDuplicatesToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //    // Lokale Hilfsfunktion: Findet alle Werte, die mehr als einmal vorkommen
    //    static HashSet<string> GetDuplicateKeys(IEnumerable<string?> items)
    //    {
    //        return items
    //            .Where(x => !string.IsNullOrWhiteSpace(x) && x != "|") // "|" ist unser Platzhalter für leere Namen
    //            .GroupBy(x => x!.Trim(), StringComparer.OrdinalIgnoreCase)
    //            .Where(g => g.Count() > 1)
    //            .Select(g => g.Key)
    //            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    //    }

    //    if (tabControl.SelectedTab == addressTabPage && _context != null)
    //    {
    //        var source = _context.Adressen.Local;

    //        // 1. Alle mehrfach vorkommenden E-Mails und Namen ermitteln
    //        var duplicateMails = GetDuplicateKeys(source.Select(a => a.Mail1));
    //        var duplicateNames = GetDuplicateKeys(source.Select(a => $"{a.Vorname?.Trim()}|{a.Nachname?.Trim()}"));

    //        // 2. Wenn alles sauber ist, eine Erfolgsmeldung ausgeben
    //        if (duplicateMails.Count == 0 && duplicateNames.Count == 0)
    //        {
    //            Utils.MsgTaskDlg(Handle, "Alles sauber", "Es wurden keine Duplikate in den lokalen Adressen gefunden.", TaskDialogIcon.ShieldSuccessGreenBar);
    //            return;
    //        }

    //        // 3. Grid filtern: Zeige nur die Datensätze, deren Mail oder Name in der Duplikat-Liste steht
    //        ExecuteFilter(source, addressBSource, addressDGV, a =>
    //        {
    //            var mail = a.Mail1?.Trim() ?? string.Empty;
    //            var fullName = $"{a.Vorname?.Trim()}|{a.Nachname?.Trim()}";

    //            return duplicateMails.Contains(mail) || duplicateNames.Contains(fullName);

    //        }, "… mögliche Duplikate", "Adressen");
    //    }
    //    else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
    //    {
    //        var source = _allGoogleContacts;

    //        // Gleiches Spiel für die Google Kontakte
    //        var duplicateMails = GetDuplicateKeys(source.Select(c => c.Mail1));
    //        var duplicateNames = GetDuplicateKeys(source.Select(c => $"{c.Vorname?.Trim()}|{c.Nachname?.Trim()}"));

    //        if (duplicateMails.Count == 0 && duplicateNames.Count == 0)
    //        {
    //            Utils.MsgTaskDlg(Handle, "Alles sauber", "Es wurden keine Duplikate in den Google Kontakten gefunden.", TaskDialogIcon.ShieldSuccessGreenBar);
    //            return;
    //        }

    //        ExecuteFilter(source, contactBSource, contactDGV, c =>
    //        {
    //            var mail = c.Mail1?.Trim() ?? string.Empty;
    //            var fullName = $"{c.Vorname?.Trim()}|{c.Nachname?.Trim()}";

    //            return duplicateMails.Contains(mail) || duplicateNames.Contains(fullName);

    //        }, "… mögliche Duplikate", "Google Kontakte");
    //    }
    //}

    private void FindDuplicatesToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ProcessDuplicateSearch(_context.Adressen.Local, addressBSource, addressDGV, "lokalen Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ProcessDuplicateSearch(_allGoogleContacts, contactBSource, contactDGV, "Google-Kontakte");
        }
    }

    private void ProcessDuplicateSearch<T>(IEnumerable<T> source, BindingSource bs, DataGridView dgv, string sourceName) where T : class, IContactEntity
    {
        // Lokale Hilfsfunktion bleibt intern
        static HashSet<string> GetDuplicateKeys(IEnumerable<string?> items)
        {
            return items
                .Where(x => !string.IsNullOrWhiteSpace(x) && x != "|")
                .GroupBy(x => x!.Trim(), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        // 1. Kriterien auswerten
        var duplicateMails = GetDuplicateKeys(source.Select(a => a.Mail1));
        var duplicateNames = GetDuplicateKeys(source.Select(a => $"{a.Vorname?.Trim()}|{a.Nachname?.Trim()}"));

        // 2. Erfolgskontrolle
        if (duplicateMails.Count == 0 && duplicateNames.Count == 0)
        {
            Utils.MsgTaskDlg(Handle, "Gratulation!", $"Die {sourceName} wurden durchsucht.\nEs wurden keine Duplikate gefunden.", TaskDialogIcon.ShieldSuccessGreenBar);
            return;
        }

        // 3. Filter ausführen
        ExecuteFilter(source, bs, dgv, a =>
        {
            var mail = a.Mail1?.Trim() ?? string.Empty;
            var fullName = $"{a.Vorname?.Trim()}|{a.Nachname?.Trim()}";
            return duplicateMails.Contains(mail) || duplicateNames.Contains(fullName);
        }, "… mögliche Duplikate", sourceName);

        // 4. Transparente Ergebnismeldung (Zusammenfassung)
        var detailInfo = $"Die {sourceName} wurden durchsucht.\n\n" +
                         $"Kriterien:\n" +
                         $"• Identische E-Mail-Adresse (Feld: Mail 1)\n" +
                         $"• Identische Kombination aus Vorname und Nachname\n\n" +
                         $"Ergebnis:\n" +
                         $"• {duplicateNames.Count} Namen kommen mehrfach vor.\n" +
                         $"• {duplicateMails.Count} E-Mail-Adressen kommen mehrfach vor.";

        Utils.MsgTaskDlg(Handle, "Duplikate identifiziert", detailInfo, TaskDialogIcon.ShieldWarningYellowBar);
    }

    private void ClipboardTSButton_Click(object sender, EventArgs e) => ClipboardTSMenuItem_Click(sender, e);

    private void NotifyIcon_MouseDoubleClick(object sender, MouseEventArgs e) => RestoreFromTray();

    private void HideToTray()
    {
        notifyIcon.Visible = true;
        Hide(); // Versteckt das Fenster komplett
        ShowInTaskbar = false; // Entfernt es aus der Taskleiste

        // Optional: Einmaliger Hinweis beim ersten Mal
        notifyIcon.ShowBalloonTip(3000, "Adressen & Kontakte",
            "Das Programm läuft im Hintergrund weiter.", ToolTipIcon.Info);
    }

    private async void RestoreFromTray()
    {
        var quickSplash = new FrmSplashScreen
        {
            StartPosition = FormStartPosition.CenterScreen,
            TopMost = true
        };
        quickSplash.Show();
        Opacity = 0;
        Show();
        WindowState = FormWindowState.Normal;
        ShowInTaskbar = true;
        await Task.Delay(50);  // 50ms ist ein "sicherer" Wert, 20ms ist oft das Minimum für einen Effekt
        min2TrayTSButton.Enabled = false;  // Trick, damit die Hintergrundfarbe des Buttons zurückgesetzt wird (sonst bleibt er nach dem Klick dunkel)
        min2TrayTSButton.Enabled = true;
        min2TrayTSButton.BackColor = tableLayoutPanel.BackColor;
        Opacity = 1;
        notifyIcon.Visible = false;
        BringToFront();
        Activate();
        quickSplash.Close();
        quickSplash.Dispose();
    }

    private void Min2TrayTSButton_Click(object sender, EventArgs e) => HideToTray();

    private void OpenTrayMenuItem_Click(object sender, EventArgs e) => RestoreFromTray();

    private void ExitTrayMenuItem_Click(object sender, EventArgs e) => Application.Exit();

    private void TrayMenu_Opened(object? sender, EventArgs e)
    {
        var rect = openTrayMenuItem.Bounds;
        var centerX = rect.Left + (rect.Width / 2);
        var centerY = rect.Top + (rect.Height / 2);
        var screenPoint = trayMenu.PointToScreen(new Point(centerX, centerY));
        NativeMethods.SetCursorPos(screenPoint.X, screenPoint.Y);
        //openTrayMenuItem.Select();  // visuell als gewählt markieren
    }

}

