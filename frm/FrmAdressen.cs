using System.ComponentModel;
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
    //private int contactNewRowIndex = -1;
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
    //private string userEmail = string.Empty;
    private bool _isClosing = false; // Flag, um Endlosschleife zu verhindern
    private bool _isFiltering = false; // Verhindert Speichern während der Suche
    private BindingList<Contact> _allGoogleContacts = []; // Klassenvariable
    private bool _isDarkMode;
    private CancellationTokenSource? _googleCts; // Wenn der User die Kontakte lädt und kurz darauf erneut klickt, soll der erste Ladevorgang abgebrochen werden: deshalb global!!
    private int _currentDbVersion;
    private bool _isTabSwitchingProgrammatically = false; // Verhindert unerwünschte Event-Auslösung bei Tab-Wechseln durch Code
    private TabPage? _previousTab;  // innerhalb des Selecting-Events kann man sich nicht auf tabControl.SelectedTab verlassen

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
        ApplySettingsToUI();
    }

    private void ApplySettingsToUI()
    {
        FormStateManager.RestoreWindowBounds(this, _settings.WindowPosition, _settings.WindowMaximized);
        _settings.SplitterPosition = _settings.SplitterPosition > 0 ? _settings.SplitterPosition : splitContainer.SplitterDistance;
        searchTSTextBox.TextBox.SetInnerMargins(4, 4);
        tbNotizen.SetInnerMargins(4, 4);
        maskedTextBox.SetInnerMargins(4, 4);
        maskedTextBox.SetPlaceholder("TT.MM.JJJJ");
        SetColorScheme();
        tsClearLabel.Visible = false;
    }

    private void FrmAdressen_Load(object sender, EventArgs e) => ApplyFileWatcherSettings();

    private async void FrmAdressen_Shown(object sender, EventArgs e)
    {
        Update();  // sicherer als DoEvents(), da es nur Painting betrifft; soll weiße Flächen verhindern
        Cursor.Current = Cursors.WaitCursor;
        try
        {
            splitContainer.SplitterDistance = _settings.SplitterPosition;
            flexiTSStatusLabel.Width = 244 + splitContainer.SplitterDistance - 536;
            if (Utils.IsUpdateCheckDue(_settings.UpdateIndex, _settings.LastUpdateCheck))
            {
                var (version, date) = await Utils.GetLatestVersionInfoAsync();
                RefreshUpdateUI(version, date);
            }
            if (!argsPath) { _databaseFilePath = _settings.RecentFiles.Count > 0 ? _settings.RecentFiles[0] : string.Empty; }
            if ((_settings.ReloadRecent || argsPath) && !string.IsNullOrEmpty(_databaseFilePath)) { await ConnectSQLDatabaseAsync(_databaseFilePath); }
            else if (!_settings.ReloadRecent && !_settings.NoAutoload && !string.IsNullOrEmpty(_settings.StandardFile)) { await ConnectSQLDatabaseAsync(_settings.StandardFile); }
            if (_settings.ContactsAutoload) { await LoadAndDisplayGoogleContactsAsync(); }

        }
        finally
        {
            Cursor.Current = Cursors.Default;
            if (_splashScreen != null)
            {
                _splashScreen.Close();
                _splashScreen.Dispose();
            }
            searchTSTextBox.TextBox.Focus();
        }
    }

    private void SaveConfiguration()
    {
        _settings.WindowMaximized = WindowState == FormWindowState.Maximized;
        var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        _settings.WindowPosition = new WindowPlacement { X = bounds.X, Y = bounds.Y, Width = bounds.Width, Height = bounds.Height };
        _settings.SplitterPosition = splitContainer.SplitterDistance;
        var activeDGV = tabControl.SelectedTab == contactTabPage ? contactDGV : addressDGV;
        if (activeDGV.Columns.Count > 0)
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
                    "Bitte aktualisieren Sie Ihr Programm.", TaskDialogIcon.ShieldErrorRedBar);
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
                statusStrip.Update(); // Text sofort malen

                var ownerHandle = Handle;
                migrationDone = await Task.Run(() => DatabaseMigrator.MigrateLegacyData(_context, ownerHandle));
                if (migrationDone) { _currentDbVersion = AppSettings.DatabaseSchemaVersion; }
            }


            // SCHRITT C: Laden (Der längste Teil)
            // Wir setzen ihn auf 50%, wohl wissend, dass er hier kurz "hängt"
            toolStripProgressBar.Value = 50;
            toolStripStatusLabel.Text = "Lade Datensätze...";
            statusStrip.Update();

            // 1. NEU: Alle Gruppen vorab in den Cache laden
            // Das löst das Problem, dass im Filter-Dialog nur benutzte Gruppen auftauchen.
            // Jetzt kennt .Local sofort alle verfügbaren Gruppen.
            await _context.Gruppen.LoadAsync();

            // 2. Adressen laden (mit Optimierung)
            await _context.Adressen
                // NEU: Eager Loading für die Gruppen-Beziehung.
                // Lädt die Verknüpfungen sofort mit. Das verhindert das "Nachploppen"
                // und stellt sicher, dass ChangeTracker die Beziehungen sofort kennt.
                .Include(a => a.Gruppen)

                // Sortierung (wie gehabt)
                .OrderBy(a => EF.Functions.Collate(a.Nachname, "GERMAN"))
                .ThenBy(a => EF.Functions.Collate(a.Vorname, "GERMAN"))
                .LoadAsync();

            // SCHRITT D: UI Aufbau (Binding)
            // Daten sind da, jetzt geht es ans Anzeigen
            toolStripProgressBar.Value = 80;
            toolStripStatusLabel.Text = "Erstelle Ansicht...";
            statusStrip.Update();

            addressBindingSource.DataSource = _context.Adressen.Local.ToBindingList();
            addressDGV.DataSource = addressBindingSource;
            AutoValidate = AutoValidate.EnableAllowFocusChange; // Fehler im Validating-Event anzeigen, aber Fokuswechsel erlauben; Standard = EnablePreventFocusChange
            ApplyColumnSettings(addressDGV);
            foreach (DataGridViewColumn column in addressDGV.Columns) { column.SortMode = DataGridViewColumnSortMode.NotSortable; }

            PopulateMemberships();
            SwitchDataBinding(addressBindingSource);

            if (_context != null)
            {
                _settings.RecentFiles.Remove(_databaseFilePath);
                _settings.RecentFiles.Insert(0, _databaseFilePath);
                if (_settings.RecentFiles.Count > AppSettings.MaxRecentFiles) { _settings.RecentFiles = [.. _settings.RecentFiles.Take(AppSettings.MaxRecentFiles)]; }

                newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = deleteTSButton.Enabled = newTSButton.Enabled = duplicateToolStripMenuItem.Enabled = copyTSButton.Enabled = wordTSButton.Enabled = envelopeTSButton.Enabled = true;
                copyToOtherDGVTSMenuItem.Enabled = false;

                tabControl.SelectTab(0);

                _context.ChangeTracker.StateChanged += OnStateChanged;
                addressBindingSource.CurrentChanged += AddressBindingSource_CurrentChanged;

                if (addressBindingSource.Count > 0) { AddressBindingSource_CurrentChanged(this, EventArgs.Empty); }

                if (!migrationDone && _settings.BirthdayAddressShow)
                {
                    BeginInvoke(new Action(() => { BirthdayReminder(addressDGV); }));
                }

                _ = Task.Run(() => Utils.StartSearchCacheWarmup(_context.Adressen.Local));

                // SCHRITT E: Fertig
                AddressBindingSource_CurrentChanged(addressBindingSource, EventArgs.Empty);  // Einmalig feuern für den ersten Datensatz
                toolStripProgressBar.Value = 100; // Voller Balken
                toolStripStatusLabel.Text = $"{addressBindingSource.Count} Adressen geladen.";
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
            ActiveControl = null;  // Zwingt das aktuelle Control (z.B. eine TextBox), den Wert zu pushen.
            isInputValid = ValidateChildren(ValidationConstraints.Enabled);  // nur die relevanten Controls (Edit-Controls) validieren
            addressBindingSource?.EndEdit();  // damit ein durch 'AddNew' erzeugter Datensatz final in die Liste übernommen und vom ChangeTracker erkannt wird
        }
        finally
        {
            addressDGV.CausesValidation = true;
            contactDGV.CausesValidation = true;
        }
        var analysis = DbChangeAnalyzer.AnalyzeChanges(_context);
        if (_context == null || !analysis.HasChanges)
        {
            if (closeDB) { CloseDatabaseConnection(); }
            return DialogResult.None;
        }
        if (!askNever && _settings.AskBeforeSaveSQL)
        {
            if (tabControl.SelectedTab != addressTabPage) { tabControl.SelectTab(addressTabPage); }
            TaskDialogButton saveButton = new("&Speichern");
            TaskDialogButton dontSaveButton = new("&Nicht speichern");
            var cancelButton = TaskDialogButton.Cancel;
            TaskDialogPage page = new()
            {
                Caption = $"{appName} - {Path.GetFileName(_databaseFilePath)}",
                Heading = analysis.DialogHeading,
                Text = analysis.DialogText,
                Icon = TaskDialogIcon.ShieldWarningYellowBar,
                AllowCancel = true,
                SizeToContent = true,
                Buttons = { saveButton, dontSaveButton, cancelButton }
            };

            // NEU: Expander hinzufügen
            if (!string.IsNullOrWhiteSpace(analysis.ExpanderText))
            {
                page.Expander = new TaskDialogExpander()
                {
                    Text = analysis.ExpanderText,
                    Position = TaskDialogExpanderPosition.AfterText,
                    Expanded = false // Standardmäßig eingeklappt
                };
            }

            var result = TaskDialog.ShowDialog(this, page);
            if (result == cancelButton) { return DialogResult.Cancel; }
            if (result == dontSaveButton)
            {
                _isFiltering = true;
                try
                {
                    await DbChangeAnalyzer.RevertChangesAsync(analysis.RealChanges);
                    _context.ChangeTracker.Entries().Where(e => e.State != EntityState.Unchanged).ToList().ForEach(e => e.State = EntityState.Unchanged);  // "Nachbeben" beseitigen
                }
                finally { _isFiltering = false; }
                if (closeDB) { CloseDatabaseConnection(); }
                return DialogResult.No;
            }
        }
        if (!isInputValid)
        {
            Utils.MsgTaskDlg(Handle, "Speichern nicht möglich", "Einige Eingaben sind ungültig oder unvollständig.", TaskDialogIcon.ShieldErrorRedBar);
            return DialogResult.Cancel;
        }

        try
        {
            await _context.SaveChangesAsync();
            await _context.Database.ExecuteSqlRawAsync("PRAGMA wal_checkpoint(TRUNCATE);");  // WAL-Checkpoint erzwingen, damit die .db-Datei für das Backup vollständig ist!
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
                if (isFormClosing) { await Utils.DailyBackupAsync(_databaseFilePath, _settings.BackupDirectory); }  // Beim Schließen geht Sicherheit vor! Kein Beenden bevor die Sicherung fertig ist!
                else { _ = Utils.DailyBackupAsync(_databaseFilePath, _settings.BackupDirectory); }  // "Fire-and-Forget", Programm soll sofort wieder bedienbar sein, ohne auf die Sicherung zu warten.
            }
            if (_settings.AddZipBackup && File.Exists(_databaseFilePath) && !string.IsNullOrWhiteSpace(_settings.AddZipDirectory))
            {
                if (isFormClosing) { await Utils.UpdateZipBackupAsync(_databaseFilePath, _settings.AddZipDirectory); }
                else { _ = Utils.UpdateZipBackupAsync(_databaseFilePath, _settings.AddZipDirectory); }
            }
            return DialogResult.Yes;
        }
        catch (DbUpdateConcurrencyException dbEx)
        {
            Utils.MsgTaskDlg(Handle, "Konflikt beim Speichern", $"Details: {dbEx.Message}\nIhre lokalen Änderungen werden verworfen.");
            foreach (var entry in dbEx.Entries) { await entry.ReloadAsync(); }
            saveTSButton.Enabled = false;
            return DialogResult.Abort;
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            return DialogResult.Abort;
        }
        finally { if (closeDB) { CloseDatabaseConnection(); } }
    }

    private void CloseDatabaseConnection()
    {
        // 1. Events abklemmen, damit keine Logik mehr getriggert wird
        addressBindingSource.CurrentChanged -= AddressBindingSource_CurrentChanged;
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
        addressBindingSource.DataSource = typeof(Adresse);
        contactBindingSource.DataSource = typeof(Contact);

        // 5. Context entsorgen
        _context?.Dispose();
        _context = null;

        Debug.WriteLine("Datenbankverbindung sicher getrennt.");
    }

    private async void OpenToolStripMenuItem_Click(object? sender, EventArgs? e)
    {
        await CheckContactChanges(async () =>
        {
            openFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";

            var fullPath = _databaseFilePath;
            var fileName = Path.GetFileName(fullPath) ?? "Adressen.adb";
            var dirName = Path.GetDirectoryName(fullPath);

            openFileDialog.FileName = fileName;
            //openFileDialog.InitialDirectory = !string.IsNullOrEmpty( sDatabaseFolder) && Directory.Exists(sDatabaseFolder) ? sDatabaseFolder : dirName ?? string.Empty;
            openFileDialog.InitialDirectory = (_settings.DatabaseFolder is { Length: > 0 } dbDir && Directory.Exists(dbDir)) ? dbDir : dirName ?? string.Empty;
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                // Falls schon eine DB offen ist, sauber schließen und speichern
                if (_context != null)
                {
                    // WICHTIG: Hier speichern wir und schließen die Verbindung.
                    // CloseDatabaseConnection wird durch 'true' bereits aufgerufen.
                    await SaveSQLDatabaseAsync(true);
                }

                // Jetzt die neue Datenbank laden
                await ConnectSQLDatabaseAsync(openFileDialog.FileName);

                // UI-Reset nach dem Laden
                ignoreSearchChange = true;
                searchTSTextBox.Text = string.Empty;
                ApplyGlobalSearch(string.Empty);
                ignoreSearchChange = false;
            }

            //if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            //{
            //    if (addressBindingSource != null && _context != null) { await SaveSQLDatabaseAsync(true); }
            //    //ConnectSQLDatabase(openFileDialog.FileName);
            //    await ConnectSQLDatabaseAsync(openFileDialog.FileName);
            //    ignoreSearchChange = true;
            //    searchTSTextBox.Text = string.Empty;
            //    ApplyGlobalSearch(string.Empty); // Filter komplett zurücksetzen
            //    ignoreSearchChange = false;
            //}
        });
    }

    private async void ExitToolStripMenuItem_Click(object? sender, EventArgs? e)
    {
        if (addressBindingSource != null) { await SaveSQLDatabaseAsync(true); }
        Close();
    }

    private async void AddressDGV_CellClick(object sender, DataGridViewCellEventArgs e)
    {
        // 1. Validitätsprüfung (Header-Klicks ausschließen)
        if (e.RowIndex < 0 || e.ColumnIndex < 0)
        {
            return;
        }

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

    private async void AddressBindingSource_ListChanged(object? sender, ListChangedEventArgs e)
    {
        if (addressBindingSource.Current is Adresse currentAdresse && _context != null)
        {
            // Wir prüfen, ob die Relationen schon geladen sind. Wenn nicht, laden wir sie nach.
            // Das passiert asynchron im Hintergrund, während die UI schon da ist.

            var entry = _context.Entry(currentAdresse);

            if (!entry.Collection(a => a.Gruppen).IsLoaded)
            {
                await entry.Collection(a => a.Gruppen).LoadAsync();
                // Ggf. UI updaten, die Gruppen anzeigt
            }

            if (!entry.Collection(a => a.Dokumente).IsLoaded)
            {
                await entry.Collection(a => a.Dokumente).LoadAsync();
                // Ggf. UI updaten, die Dokumente anzeigt
            }

            // Nachfolgender Code wird in ShowPhotoInPictureBoxy aufgerufen, wenn die Adresse gewechselt wird. Dort prüfen wir dann, ob das Foto geladen ist, und laden es bei Bedarf nach.
            //// Foto ist 1:1, das laden wir auch explizit bei Bedarf  
            //if (!entry.Reference(a => a.Foto).IsLoaded)
            //{
            //    await entry.Reference(a => a.Foto).LoadAsync();
            //}
        }
        UpdateSaveButton();
    }

    private void AddressBindingSource_CurrentChanged(object? sender, EventArgs e)
    {
        if (_isFiltering) { return; } // Konflikte mit Suchfilter vermeiden
        try
        {
            ignoreTextChange = true;
            if (addressBindingSource?.Current is Adresse currentAdresse)
            {
                ErzeugeGrussformeln();
                ShowPhotoInPictureBoxy(currentAdresse);
                UpdateMembershipCBox();
                LoadGroupsForCurrentAddress();
                UpdateDocumentListView(currentAdresse);
                if (currentAdresse.Geburtstag.HasValue) { AgeLabel_MaskedTB_Set(currentAdresse.Geburtstag.Value); }
                else { AgeLabel_MaskedTB_Clear(); }
            }
            else
            {
                topAlignZoomPictureBox.Image = Resources.AddressBild100;
                delPictboxToolStripButton.Enabled = false;
                flowLayoutPanel.Controls.Clear();
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
        if (colCount == 0) { return; } // Nichts zu tun
        for (var i = 0; i < colCount; i++)
        {
            if (i < _settings.HideColumnArr.Length) { dgv.Columns[i].Visible = !_settings.HideColumnArr[i]; }
            if (i < _settings.ColumnWidths.Length) { dgv.Columns[i].Width = Math.Max(20, _settings.ColumnWidths[i]); }
        }
    }

    private void OpenTSButton_Click(object sender, EventArgs e) => OpenToolStripMenuItem_Click(sender, e);

    private void FrmAdressen_Resize(object sender, EventArgs e)
    {
        flexiTSStatusLabel.Width = 244 + splitContainer.SplitterDistance - 536;
        searchTSTextBox.Width = 202 + splitContainer.SplitterDistance - 536 - (tsClearLabel.Visible ? tsClearLabel.Width : 0);
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
        var term = searchText.Trim().ToLower();  // ToLower für case-insensitive Suche ist evtl. nicht nötig wg. COLLATE NOCASE in AdressenDbContext
        var isSearchEmpty = string.IsNullOrWhiteSpace(term);
        _isFiltering = true;

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
        //currencyManager?.SuspendBinding(); // Nicht verwenden wenn DataSource der BindingSource getauscht wird!

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
                if (addressBindingSource.Current != null) { ShowPhotoInPictureBoxy(addressBindingSource.Current); }
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
                if (contactBindingSource.Current is Contact selectedContact) { ShowPhotoInPictureBoxy(selectedContact); }
            }
        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
        finally
        {
            //currencyManager?.ResumeBinding(); // Nicht verwenden wenn DataSource der BindingSource getauscht wird!
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
        // 1. Fall: Adressen-Tab
        if (tabControl.SelectedTab == addressTabPage && addressBindingSource?.Current is Adresse)
        {
            var result = await SaveSQLDatabaseAsync(false, true);
            if (result == DialogResult.Yes || result == DialogResult.None)
            {
                saveTSButton.Enabled = false;
            }
            return;
        }

        // 2. Fall: Kontakt-Tab (Guard Clause: Wenn falscher Tab oder falscher Datensatz -> Abbruch)
        if (tabControl.SelectedTab != contactTabPage || contactBindingSource.Current != _lastActiveContact)
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

        contactBindingSource.EndEdit();
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
            contactBindingSource.ResetBindings(false);
        }
    }

    private void TbNotizen_SizeChanged(object sender, EventArgs e) => NativeMethods.ShowScrollBar(tbNotizen.Handle, 1, TextRenderer.MeasureText(tbNotizen.Text, tbNotizen.Font,
        new Size(tbNotizen.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height > tbNotizen.Height);

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
                contactBindingSource.Add(newContact); // Löst evtl. Events aus -> werden durch Lock ignoriert
                contactBindingSource.ResetBindings(false);

                var realIndex = contactBindingSource.IndexOf(newContact);

                if (realIndex >= 0)
                {
                    // 4. Position wechseln
                    contactBindingSource.Position = realIndex; // Löst RowValidating aus -> wird durch Lock ignoriert

                    if (contactDGV.RowCount > realIndex)
                    {
                        contactDGV.FirstDisplayedScrollingRowIndex = realIndex;
                        contactDGV.Rows[realIndex].Selected = true;
                    }
                }

                // 5. Interne Referenzen auf den NEUEN Kontakt biegen
                _lastActiveContact = newContact;
                _originalContactSnapshot = (Contact)newContact.Clone();

                // UI Updates...
                ShowPhotoInPictureBoxy(newContact);
                UpdateMembershipTags();
                UpdateSaveButton();
            }
            finally { isSelectionChanging = false; }

            if (cbAnrede.CanFocus) { cbAnrede.Focus(); }
        }
        else if (tabControl.SelectedTab == addressTabPage && addressBindingSource != null)
        {
            addressBindingSource.AddNew();  // noch nicht fest in die zugrunde liegende BindingList "committed".
            addressBindingSource.EndEdit(); // dadurch wird der Status im EF ChangeTracker zuverlässig auf 'Added' gesetzt
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
                if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact originalContact)
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
                    contactBindingSource.ResetBindings(false);

                    // Position finden und ansteuern
                    var newIndex = _allGoogleContacts.IndexOf(clone);
                    if (newIndex >= 0)
                    {
                        contactBindingSource.Position = newIndex;

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
                else if (tabControl.SelectedTab == addressTabPage && addressBindingSource?.Current is Adresse originalAdresse && _context != null)
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

                    duplikat.Id = 0;
                    duplikat.Foto?.Id = 0;

                    // Einfügeposition bestimmen
                    var insertIndex = Utils.GetAddressInsertIndex(addressBindingSource, duplikat);

                    // In BindingSource einfügen
                    addressBindingSource.Insert(insertIndex, duplikat);
                    addressBindingSource.Position = insertIndex;

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
        if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact selectedGoogleContact)
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
                    var currentIdx = addressBindingSource.Position;
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
        else if (tabControl.SelectedTab == addressTabPage && addressBindingSource.Current is Adresse selectedLocalAddress)
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
                    var currentIdx = contactBindingSource.Position;
                    if (currentIdx >= 0 && currentIdx < contactDGV.RowCount)
                    {
                        // 1. Scrollen
                        contactDGV.FirstDisplayedScrollingRowIndex = currentIdx;

                        // 2. Zeile markieren
                        contactDGV.Rows[currentIdx].Selected = true;

                        // 3. Fokus auf erste SICHTBARE Zelle setzen (Fix für den Absturz)
                        var firstVisibleCol = contactDGV.Columns.GetFirstColumn(DataGridViewElementStates.Visible);
                        if (firstVisibleCol != null)
                        {
                            contactDGV.CurrentCell = contactDGV.Rows[currentIdx].Cells[firstVisibleCol.Index];
                        }
                    }
                }

                cbAnrede.Focus();
                saveTSButton.Enabled = false;
                flexiTSStatusLabel.Text = "Kontakt erfolgreich zu Google kopiert.";
            }
            else
            {
                tabControl.SelectedTab = addressTabPage;
            }
        }
        else
        {
            Console.Beep();
        }
    }

    private async void DeleteTSButton_Click(object sender, EventArgs e)
    {
        // 1. Der Gatekeeper: Prüft auf ungespeicherte Änderungen
        await CheckContactChanges(async () =>
        {
            if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact googleKontakt)
            {
                var (askBefore, deleteNow) = Utils.AskBeforeDeleteContact(Handle, googleKontakt, _settings.AskBeforeDelete, false);
                _settings.AskBeforeDelete = askBefore;

                if (!deleteNow) { return; }

                // Lock setzen
                isSelectionChanging = true;
                try
                {
                    // Aufruf über den TaskDialog aus der Utils-Klasse
                    var success = await Utils.RunWithProgressDialogAsync(
                        this,
                        "Kontakt löschen",
                        "Der Kontakt wird bei Google gelöscht...",
                        async token =>
                        {
                            await DeleteGoogleContactAsync(googleKontakt, token);
                        });

                    // Nur bei Erfolg aus der Liste entfernen
                    if (success)
                    {
                        _allGoogleContacts?.Remove(googleKontakt);
                        contactBindingSource.RemoveCurrent();

                        // State-Reset
                        _lastActiveContact = null;
                        _originalContactSnapshot = null;

                        UpdateContactStatusBar();
                    }
                }
                // Ein explizites Catch ist hier nicht zwingend nötig, da RunWithProgressDialogAsync 
                // Exceptions bereits abfängt und per TaskDialog anzeigt.
                finally
                {
                    isSelectionChanging = false;
                }
            }
            else if (tabControl.SelectedTab == addressTabPage && addressBindingSource.Current is Adresse adresseZumLoeschen && _context != null)
            {
                if (addressBindingSource.IsBindingSuspended || adresseZumLoeschen == null) { return; }
                if (addressDGV.CurrentRow?.IsNewRow == true) { return; }
                addressBindingSource.EndEdit();
                var deleteFinal = true;
                if (_settings.AskBeforeDelete)
                {
                    var (askBefore, deleteNow) = Utils.AskBeforeDeleteAddress(Handle, adresseZumLoeschen, _settings.AskBeforeDelete);
                    _settings.AskBeforeDelete = askBefore;
                    deleteFinal = deleteNow;
                }
                if (!deleteFinal) { return; }
                try
                {
                    var entry = _context.Entry(adresseZumLoeschen);
                    var isNewRecord = entry.State == EntityState.Added || adresseZumLoeschen.Id == 0;

                    if (isNewRecord)
                    {
                        if (adresseZumLoeschen.Foto is not null)
                        {
                            var fotoEntry = _context.Entry(adresseZumLoeschen.Foto);
                            if (fotoEntry.State == EntityState.Added || adresseZumLoeschen.Foto.Id == 0) { fotoEntry.State = EntityState.Detached; }
                        }
                        entry.State = EntityState.Detached;
                    }
                    else { _context.Adressen.Remove(adresseZumLoeschen); }

                    if (addressBindingSource.Contains(adresseZumLoeschen)) { addressBindingSource.Remove(adresseZumLoeschen); }
                    UpdateSaveButton();
                    UpdateAddressStatusBar();

                    if (addressBindingSource.Count > 0) { addressDGV.Rows[addressBindingSource.Position].Selected = true; }
                }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
            else { Console.Beep(); }
        });
    }

    private void ExecuteAndPreserveSelection<T>(BindingSource bindingSource, DataGridView grid, Action dataUpdateAction) where T : class
    {
        T? currentItem = null;  // aktuelles Objekt merken
        if (bindingSource.Current is not null) { currentItem = bindingSource.Current as T; }
        var currencyManager = BindingContext?[bindingSource] as CurrencyManager;
        currencyManager?.SuspendBinding();
        try
        {
            grid?.CurrentCell = null;
            dataUpdateAction();
        }
        finally { currencyManager?.ResumeBinding(); }
        if (currentItem != null)  // Selektion wiederherstellen
        {
            var newIndex = bindingSource.IndexOf(currentItem);
            if (newIndex >= 0)
            {
                bindingSource.Position = newIndex;

                if (grid != null && grid.RowCount > newIndex)
                {
                    grid.BeginInvoke(new Action(() =>
                    {
                        if (newIndex >= grid.RowCount || newIndex < 0) { return; }
                        var row = grid.Rows[newIndex];
                        if (!FormStateManager.RowIsVisible(grid, row)) { grid.FirstDisplayedScrollingRowIndex = newIndex; }  // Scrollen (nur wenn nötig)
                        grid.ClearSelection(); // Alte Selektionen entfernen
                        row.Selected = true;   // Diese Zeile markieren
                        //var firstVisibleCol = grid.Columns.Cast<DataGridViewColumn>().FirstOrDefault(c => c.Visible);
                        //if (firstVisibleCol != null) { grid.CurrentCell = row.Cells[firstVisibleCol.Index]; }
                    }));
                }
            }
        }
    }

    private async void FrmAdressen_FormClosing(object sender, FormClosingEventArgs e)
    {
        // 1. Laufende Google-Requests sofort abbrechen
        _googleCts?.Cancel();

        // 2. Rekursions-Check: Wenn wir am Ende der Methode Close() rufen, springen wir hier raus
        if (_isClosing) { return; }

        // -------------------------------------------------------------
        // SCHRITT A: Prüfungen durchführen (Abbruch ermöglichen)
        // -------------------------------------------------------------

        // Fall 1: SQL Datenbank (EF Core 10)
        if (_context != null)
        {
            var result = await SaveSQLDatabaseAsync(false, false, true);
            if (result == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }
        }

        // Fall 2: Google Kontakte (Zentraler Gatekeeper)
        // Wir übergeben 'isClosing: true', damit die Methode weiß, 
        // dass sie keine UI-Resets (wie Focus) mehr machen muss.
        var readyToCloseGoogle = await ContactChanges_Check(isClosing: true);
        if (!readyToCloseGoogle)
        {
            e.Cancel = true;
            return;
        }

        // -------------------------------------------------------------
        // SCHRITT B: Aufräumen und Endgültig Schließen
        // -------------------------------------------------------------

        // Ab hier gibt es kein Zurück mehr: Wir brechen das aktuelle (synchrone) Schließen ab,
        // um den asynchronen Cleanup-Prozess vollständig zu durchlaufen.
        e.Cancel = true;
        AutoValidate = AutoValidate.Disable;
        Enabled = false;
        Cursor = Cursors.WaitCursor;

        try
        {
            SaveConfiguration();

            // Ressourcen sauber freigeben
            _googleCts?.Dispose();
            CloseDatabaseConnection();

            addressBindingSource?.Dispose();
            contactBindingSource?.Dispose();

            // Timer stoppen und entsorgen
            searchTimer?.Dispose();
            debounceTimer?.Dispose();
            scrollTimer?.Dispose();
        }
        catch (Exception ex) { Debug.WriteLine($"Fehler beim Cleanup: {ex.Message}"); }
        finally
        {
            // 3. Finales Flag setzen, Cursor zurücksetzen und Schließen neu triggern
            _isClosing = true;
            Cursor = Cursors.Default;
            Close();
        }
    }

    private void AboutToolStripMenuItem_Click(object sender, EventArgs e) => Utils.HelpMsgTaskDlg(Handle, appLong, Icon, _currentDbVersion);

    private void AddressDGV_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e) => toolStripStatusLabel.Text = addressDGV.RowCount.ToString() + " Adressen";

    private void AddressDGV_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e) => toolStripStatusLabel.Text = addressDGV.RowCount.ToString() + " Adressen";

    private void ErzeugeGrussformeln()
    {
        // Bereinigen der bestehenden Vorschläge in der TextBox
        cbGrussformel.AutoCompleteCustomSource.Clear();

        // Mapping erstellen
        var pt = new List<(string Key, string Value)> { ("#vorname", tbVorname.Text), ("#nickname", tbNickname.Text), ("#nachname", tbNachname.Text), ("#titel", cbPraefix.Text) };

        // Die Logik bleibt identisch, nur das Ziel ist nun die AutoCompleteCustomSource
        cbGrussformel.AutoCompleteCustomSource.AddRange([.. grussformelList
        .Select(s =>
        {
            var result = s;
            foreach (var (key, value) in pt.Where(p => !string.IsNullOrWhiteSpace(p.Value))) { result = result.Replace(key, value); }
            return result;
        })
        .Where(text => !text.Contains('#')) // Nur fertige Strings ohne Platzhalter
        .Distinct()]);
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
                await ConnectSQLDatabaseAsync(_databaseFilePath);
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
            writer.WriteLine("Herr;;Mustermann;Max;;;;Musterfirma;Hausmeister;Musterstraße 1;12345;Musterstadt;Deutschland;;;;12.05.1985;max@muster.de;;030123456;;0170123456;;;Notiztext;Freunde,Wichtig");

            Utils.MsgTaskDlg(Handle, "Vorlage erstellt", $"Die Datei 'adress_vorlage.csv' wurde auf Ihrem Desktop gespeichert.");
        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
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
                if (addressBindingSource != null) { await SaveSQLDatabaseAsync(true); }
                //ConnectSQLDatabase(file);  // Erst wenn das Speichern fertig ist, geht es hier weiter:
                await ConnectSQLDatabaseAsync(file);
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
            case Keys.Enter | Keys.Control:
            case Keys.Tab | Keys.Control:   // funktioniert nicht
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
                    Utils.MsgTaskDlg(Handle, "Word fehlt", "Microsoft Word wurde nicht gefunden. Bitte installieren Sie es.");
                    return;
                }
                WordProcess();
            }
            else
            {
                if (!WordManager.IsLibreOfficeInstalled)
                {
                    Utils.MsgTaskDlg(Handle, "LibreOffice fehlt", "LibreOffice Writer wurde nicht gefunden. Bitte installieren Sie es.");
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

        bookmarkTextDictionary["Praefix_Zwischenname_Nachname"] =
            string.Join(" ", new[] { praefix, zwischen, nachname }.Where(s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Vorname_Zwischenname_Nachname"] =
            string.Join(" ", new[] { vorname, zwischen, nachname }.Where(s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Vorname_Zwischenname_initial_Nachname"] =
            string.Join(" ", new[] { vorname, zwischenInitial, nachname }.Where(s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Praefix_Vorname_Zwischenname_Nachname"] =
            string.Join(" ", new[] { praefix, vorname, zwischen, nachname }.Where(s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Anrede_Praefix_Vorname_Zwischenname_Nachname"] =
            string.Join(" ", new[] { anrede, praefix, vorname, zwischen, nachname }.Where(s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Praefix_Vorname_Zwischenname_initial_Nachname"] =
            string.Join(" ", new[] { praefix, vorname, zwischenInitial, nachname }.Where(s => !string.IsNullOrWhiteSpace(s)));

        bookmarkTextDictionary["Anrede_Praefix_Vorname_Zwischenname_initial_Nachname"] =
            string.Join(" ", new[] { anrede, praefix, vorname, zwischenInitial, nachname }.Where(s => !string.IsNullOrWhiteSpace(s)));

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
        if (targetSource == null || (targetSource.DataSource == null && targetSource == contactBindingSource)) { return; }
        var useNullConversion = targetSource == addressBindingSource;  // Unterscheidung: Lokale DB (null erlaubt) vs. Google (leerer String bevorzugt)
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
        // Bereinigen der bestehenden Vorschlagslisten
        cbAnrede.AutoCompleteCustomSource.Clear();
        cbPraefix.AutoCompleteCustomSource.Clear();
        cbPLZ.AutoCompleteCustomSource.Clear();
        cbOrt.AutoCompleteCustomSource.Clear();
        cbLand.AutoCompleteCustomSource.Clear();
        cbSchlussformel.AutoCompleteCustomSource.Clear();
        cbGrussformel.AutoCompleteCustomSource.Clear();

        if (targetSource == addressBindingSource && _context != null)
        {
            var localData = _context.Adressen.Local;

            // AddRange erwartet ein String-Array. Durch den Spread-Operator [.. ] 
            // wird das Ergebnis der LINQ-Abfrage direkt in das passende Array konvertiert.
            cbAnrede.AutoCompleteCustomSource.AddRange([.. localData.Select(a => a.Anrede ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbPraefix.AutoCompleteCustomSource.AddRange([.. localData.Select(a => a.Praefix ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbPLZ.AutoCompleteCustomSource.AddRange([.. localData.Select(a => a.PLZ ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbOrt.AutoCompleteCustomSource.AddRange([.. localData.Select(a => a.Ort ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbLand.AutoCompleteCustomSource.AddRange([.. localData.Select(a => a.Land ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbSchlussformel.AutoCompleteCustomSource.AddRange([.. localData.Select(a => a.Schlussformel ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
        }
        else if (targetSource == contactBindingSource && contactBindingSource.DataSource is BindingList<Contact> contactList)
        {
            cbAnrede.AutoCompleteCustomSource.AddRange([.. contactList.Select(c => c.Anrede ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbPraefix.AutoCompleteCustomSource.AddRange([.. contactList.Select(c => c.Praefix ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbPLZ.AutoCompleteCustomSource.AddRange([.. contactList.Select(c => c.PLZ ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbOrt.AutoCompleteCustomSource.AddRange([.. contactList.Select(c => c.Ort ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbLand.AutoCompleteCustomSource.AddRange([.. contactList.Select(c => c.Land ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            cbSchlussformel.AutoCompleteCustomSource.AddRange([.. contactList.Select(c => c.Schlussformel ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
            // cbGrussformel.AutoCompleteCustomSource.AddRange([.. contactList.Select(c => c.Grussformel ?? "").Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().Order()]);
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
                Debug.WriteLine("Fehler beim Laden des Fotos: " + ex.Message);
            }
        }
    }


    #region Google Logic (Refactored)

    private async Task LoadAndDisplayGoogleContactsAsync()
    {
        // A. UI Status Checks & Datensicherheit
        if (tabControl.SelectedTab == addressTabPage && addressBindingSource != null)
        {
            if (filterRemoveToolStripMenuItem.Visible)
            {
                FilterRemoveToolStripMenuItem_Click(null!, EventArgs.Empty);
            }

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
            // DER NEUE GATEKEEPER: Erst prüfen, ob noch Änderungen offen sind.
            // Wenn der User "Abbrechen" klickt, laden wir gar nicht erst neu.
            if (!await ContactChanges_Check())
            {
                return;
            }

            lastContactSearch = searchTSTextBox.TextBox.Text;
            ignoreSearchChange = true;
            searchTSTextBox.TextBox.Clear();
            ignoreSearchChange = false;
        }

        // B. Netzwerk-Check
        if (!Utils.GoogleConnectionCheck(Handle, secretPath)) { return; }
        else { topAlignZoomPictureBox.Image = Resources.ContactBild100; }

        // 1. Alten Prozess abbrechen UND aufräumen
        if (_googleCts != null)
        {
            _googleCts.Cancel();
            _googleCts.Dispose();
        }

        // 2. Neues Token erstellen
        _googleCts = new CancellationTokenSource();
        var ct = _googleCts.Token; // Token in lokale Var, um Zugriff im Manager sicherzustellen

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

            // D. Auth-Logik (Browser-Erkennung)
            if (isNewLogin || stopwatch.ElapsedMilliseconds > 2000)
            {
                contactBirthdayFlag = false;
            }

            //userEmail = result.UserEmail;

            // E. Gruppen verarbeiten
            contactGroupsDict = result.GroupMap;
            allContactMemberships.Clear();
            foreach (var kvp in contactGroupsDict)
            {
                var gName = kvp.Value;

                // Kein Exclude Check mehr nötig!
                gName = gName.Equals("starred", StringComparison.OrdinalIgnoreCase) ? "★" : gName;
                allContactMemberships.Add(gName);
            }
            allContactMemberships.Add("★");
            toolStripProgressBar.Value = 50;

            // F. Datenbindung mit LOCK
            // Wir setzen isSelectionChanging auf true, damit das Grid beim Zuweisen 
            // der neuen Liste nicht Amok läuft (Validierung, SelectionChanged etc.)
            isSelectionChanging = true;
            try
            {
                var contactList = new BindingList<Contact>([.. result.Contacts]);

                if (contactList.Count == 0)
                {
                    toolStripStatusLabel.Text = "Keine Kontakte gefunden.";
                    contactDGV.DataSource = null;
                    return;
                }

                _allGoogleContacts = contactList;
                toolStripStatusLabel.Text = $"{contactList.Count} Kontakte geladen.";

                contactBindingSource.DataSource = contactList;
                contactDGV.DataSource = contactBindingSource;

                ApplyColumnSettings(contactDGV);
                toolStripProgressBar.Value = 80;

                SwitchDataBinding(contactBindingSource);

                tabControl.SelectedIndex = 1;
                Text = $"Kontakte - Google Kontakte";  // $"Kontakte - {userEmail}"
            }
            finally { isSelectionChanging = false; }

            // G. UI Finalisierung
            var hasRows = contactDGV.Rows.Count > 0;
            copyTSButton.Enabled = copyToOtherDGVTSMenuItem.Enabled = wordToolStripMenuItem.Enabled =
                envelopeToolStripMenuItem.Enabled = wordTSButton.Enabled = envelopeTSButton.Enabled = hasRows;

            duplicateToolStripMenuItem.Enabled = false;
            btnEditContact.Visible = true;

            if (hasRows)
            {
                contactDGV.Rows[0].Selected = true;
            }

            // Doku-Tab wegräumen falls nötig
            if (tabulation.TabPages.Contains(tabPageDoku))
            {
                deactivatedPage = tabPageDoku;
                tabulation.TabPages.Remove(tabPageDoku);
            }

            // Geburtstagserinnerung
            if (contactBirthdayFlag && _settings.BirthdayContactShow)
            {
                toolStripProgressBar.Visible = false;
                BirthdayReminder(contactDGV);
            }
            contactBirthdayFlag = true;
            toolStripProgressBar.Value = 100;

            // Background Warmup
            Utils.StartSearchCacheWarmup(_allGoogleContacts);
            UpdateMembershipCBox();
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
        // 1. Nur auf null prüfen
        if (contact == null) { return; }
        // 2. Wenn keine ResourceName da ist (Kontakt war nie bei Google), ist nichts zu tun
        if (string.IsNullOrEmpty(contact.ResourceName)) { return; }

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

                var index = contactBindingSource.IndexOf(contact);
                if (index >= 0) { contactBindingSource.ResetItem(index); }
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
            ShowPhotoInPictureBoxy(contact);
        }
        catch (Exception ex)
        {
            if (ex.Message.Contains("NotFound")) // Einfacher Check statt using Google...
            {
                Utils.MsgTaskDlg(Handle, "Kein Foto", "Es konnte online kein Foto gefunden werden.", TaskDialogIcon.Information);
                contact.PhotoUrl = null;
                ShowPhotoInPictureBoxy(contact);
            }
            else
            {
                Utils.ErrTaskDlg(Handle, ex);
            }
        }
    }

    #endregion
    private async void GoogleTSButton_Click(object sender, EventArgs e) => await CheckContactChanges(LoadAndDisplayGoogleContactsAsync);

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
        if (e.RowIndex < 0 || e.ColumnIndex < 0)
        {
            return;
        }

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
            ErzeugeGrussformeln();

            // --- D: Geburtstag & Alter ---
            if (contact.Geburtstag.HasValue) { AgeLabel_MaskedTB_Set(contact.Geburtstag.Value); }
            else { AgeLabel_MaskedTB_Clear(); }

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
            if (contactBindingSource.DataSource == null || contactBindingSource.Count == 0)
            {
                // Hinweis: Um "async void" Probleme zu minimieren, lagern wir das Laden oft aus.
                // Hier ist es okay, aber der Dialog blockiert kurz den Tab-Wechsel visuell.
                var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Google Kontakte", "Keine Kontakte vorhanden", "Möchten Sie Ihre Kontakte jetzt laden?");

                if (isYes)
                {
                    await LoadAndDisplayGoogleContactsAsync();
                }
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
            SwitchDataBinding(addressBindingSource);

            if (addressBindingSource.Current != null)
            {
                ShowPhotoInPictureBoxy(addressBindingSource.Current);
            }

            // UI Status
            if (addressBindingSource?.Count > 0)
            {
                Text = $"{appName} – {(string.IsNullOrEmpty(_databaseFilePath) ? "unbenannt" : _databaseFilePath)}";
                btnEditContact.Visible = false;
                UpdateSaveButton();

                // Buttons aktivieren
                SetCommonButtonState(true);
                copyToOtherDGVTSMenuItem.Enabled = false;

                // Statuszeile
                var rowCount = _context?.Adressen.Local.Count ?? 0;
                var visibleRowCount = addressBindingSource.Count;
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
            if (contactBindingSource.Current is Contact current)
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
            if (contactBindingSource.DataSource != null)
            {
                SwitchDataBinding(contactBindingSource);
                if (contactBindingSource.Current != null)
                {
                    ShowPhotoInPictureBoxy(contactBindingSource.Current);
                }
            }

            // UI Status
            if (contactBindingSource.Count > 0)
            {
                //Text = !string.IsNullOrWhiteSpace(userEmail) ? $"Kontakte - {userEmail}" : "Google-Kontakte";
                Text = "Kontakte - Google Kontakte";
                btnEditContact.Visible = true;

                // Menü Items gemäß Logik (Google-Tab hat andere Regeln für Neu/Löschen im Menü)
                newToolStripMenuItem.Enabled = duplicateToolStripMenuItem.Enabled = deleteToolStripMenuItem.Enabled = false;

                // Toolbar Buttons aktivieren
                SetCommonButtonState(true);
                copyToOtherDGVTSMenuItem.Enabled = true;

                toolStripStatusLabel.Text = $"{contactBindingSource.Count} Kontakte";
            }
        }

        // Common Cleanup
        flexiTSStatusLabel.Text = string.Empty;
        searchTSTextBox.TextBox.Focus();
    }

    private void SetCommonButtonState(bool enabled)
    {
        newTSButton.Enabled = duplicateToolStripMenuItem.Enabled =
        deleteToolStripMenuItem.Enabled = deleteTSButton.Enabled =
        copyTSButton.Enabled = wordTSButton.Enabled = envelopeTSButton.Enabled = enabled;
    }

    private void HandleSearchTransition(ref string sourceStorage, ref string targetStorage)
    {
        if (searchTSTextBox.TextBox.TextLength > 0)
        {
            sourceStorage = searchTSTextBox.Text;
            ignoreSearchChange = true;
            searchTSTextBox.TextBox.Clear();
            ignoreSearchChange = false;
        }

        if (!string.IsNullOrEmpty(targetStorage))
        {
            ignoreSearchChange = true;
            searchTSTextBox.TextBox.Text = targetStorage;
            ignoreSearchChange = false;
            targetStorage = string.Empty;
        }
    }

    private void AuthentMenuItem_Click(object sender, EventArgs e)
    {
        using TaskDialogIcon questionDialogIcon = new(Resources.question32);
        TaskDialogPage page = new()
        {
            Caption = appCont,
            Heading = "Möchten Sie die Zugangsdaten löschen?",
            Text = "Wenn Sie den Request-Token löschen, können Sie\nnur nach erneuter Autorisierung Google-Kontakte\nherunterladen. Hierzu öffnet sich beim nächsten\nVersuch automatisch die Goolge-Anmeldeseite.",
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
        Cursor = Cursors.WaitCursor;
        FillDictionary();
        using var frm = new FrmPrintSetting(_settings, bookmarkTextDictionary);
        FormStateManager.RestoreWindowBounds(frm, _settings.PrintWindowPosition);
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
        fileSystemWatcher.IncludeSubdirectories = true;
        fileSystemWatcher.Filters.Clear();
        foreach (var pattern in documentTypes) { fileSystemWatcher.Filters.Add(pattern); }

        // Pfad setzen und nur aktivieren, wenn alles passt
        if (_settings.WatchFolder && !string.IsNullOrEmpty(docPath) && Directory.Exists(docPath))
        {
            fileSystemWatcher.Path = docPath;
            fileSystemWatcher.EnableRaisingEvents = true;
        }
        else { fileSystemWatcher.EnableRaisingEvents = false; }
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

    //private void TsClearLabel_Click(object sender, EventArgs e) => Clear_SearchTextBox();
    private async void TsClearLabel_Click(object sender, EventArgs e) => await CheckContactChanges(Clear_SearchTextBox);

    private void TsClearLabel_VisibleChanged(object sender, EventArgs e) => searchTSTextBox.Width = 202 + splitContainer.SplitterDistance - 536 - (tsClearLabel.Visible ? tsClearLabel.Width : 0);

    private void TsClearLabel_Paint(object sender, PaintEventArgs e) => BeginInvoke(new Action(() => Graphics.FromHwnd(toolStrip.Handle).DrawRectangle(Pens.Black, tsClearLabel.Bounds.Location.X - 2, tsClearLabel.Bounds.Location.Y + 2, tsClearLabel.Width + 1, tsClearLabel.Height - 4)));

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
        maskedTextBox.Mask = @"00\.00\.0000";
        maskedTextBox.BackColor = _isDarkMode ? Color.FromArgb(80, 80, 0) : Color.LightYellow;
        maskedTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
        if (string.IsNullOrWhiteSpace(maskedTextBox.Text.Replace(".", "").Replace("_", "").Trim())) // falls leer, Cursor ganz links
        {
            maskedTextBox.SelectionStart = 0;
            maskedTextBox.SelectionLength = 0;
        }
        else { maskedTextBox.SelectAll(); } // falls schon was drin steht, alles markieren    
        ignoreTextChange = false;
    }

    private void FormatAndSetDate()
    {
        // Nur Ziffern extrahieren
        var digits = new string([.. maskedTextBox.Text.Where(char.IsDigit)]);

        if (string.IsNullOrEmpty(digits))
        {
            maskedTextBox.Mask = "";
            maskedTextBox.Text = "";
            return;
        }

        // Aktuelles Datum als Basis
        var today = DateTime.Today;
        string d = "01", m = "01";
        var y = today.Year;

        switch (digits.Length)
        {
            case <= 2: // Nur Tag (z.B. "05")
                {
                    d = digits.PadLeft(2, '0');
                    m = today.Month.ToString("00");
                    break;
                }
            case 3:
            case 4: // Tag und Monat (z.B. "0512")
                {
                    d = digits[..2];
                    m = digits[2..].PadLeft(2, '0');
                    break;
                }
            case 5:
            case 6: // Tag, Monat, kurzes Jahr (z.B. "051224")
                {
                    d = digits[..2];
                    m = digits.Substring(2, 2);
                    var yearPart = digits[4..];

                    if (yearPart.Length == 1)
                    {
                        // Einstellige Jahre (selten) -> 200x
                        y = int.Parse("200" + yearPart);
                    }
                    else
                    {
                        // Zweistellige Jahre: Rolling Century Logik
                        // Wenn eingegebenes Jahr > aktuelles Jahr (z.B. 90 > 26), dann 19xx, sonst 20xx
                        var shortY = int.Parse(yearPart);
                        var currentShort = today.Year % 100;
                        var century = (shortY > currentShort) ? 1900 : 2000;
                        y = century + shortY;
                    }
                    break;
                }
            case 8: // Komplettes Datum (z.B. "05121990")
                {
                    d = digits[..2];
                    m = digits.Substring(2, 2);
                    if (int.TryParse(digits.AsSpan(4, 4), out var fullYear)) { y = fullYear; }
                    break;
                }
        }

        ignoreTextChange = true;
        try
        {
            // Versuch, das Datum zu bauen
            var dateString = $"{d}.{m}.{y}";

            if (DateTime.TryParseExact(dateString, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var resultDate))
            {
                // Logik: Datum muss in der Vergangenheit liegen
                if (resultDate > today)
                {
                    if (digits.Length <= 4)
                    {
                        // Fall 1: Jahr wurde automatisch ergänzt (Input war nur Tag/Monat)
                        // Beispiel: Eingabe "20.12" am 10.01.2026 -> Ergab 20.12.2026 (Zukunft) -> Korrektur auf 20.12.2025
                        resultDate = resultDate.AddYears(-1);
                    }
                    else if (digits.Length is 5 or 6)
                    {
                        // Fall 2: Jahr wurde zweistellig eingegeben
                        // Beispiel: Eingabe "050126" am 01.01.2026 -> Ergab 05.01.2026 (Zukunft) -> Korrektur auf 05.01.1926
                        resultDate = resultDate.AddYears(-100);
                    }
                    // Fall 3 (Länge 8): Wenn User explizit 4-stelliges Zukunftsjahr tippt, lassen wir es (oder validieren es als Fehler),
                    // hier wird es der Einfachheit halber akzeptiert, da keine automatische Annahme getroffen wurde.
                }

                maskedTextBox.Text = resultDate.ToString("dd.MM.yyyy");
            }
            else
            {
                // Ungültiges Datum (z.B. 30.02.)
                maskedTextBox.Mask = "";
                maskedTextBox.Text = "";
            }

            maskedTextBox.DataBindings["Text"]?.WriteValue();
        }
        finally { ignoreTextChange = false; }
    }

    private void MaskedTextBox_Leave(object sender, EventArgs e)
    {
        ignoreTextChange = true;
        maskedTextBox.BackColor = _isDarkMode ? Color.FromArgb(45, 45, 45) : Color.White;
        maskedTextBox.ForeColor = _isDarkMode ? Color.White : Color.Black;
        try
        {
            var digits = new string([.. maskedTextBox.Text.Where(char.IsDigit)]); // nur die Ziffern behalten
            if (digits.Length == 0) { AgeLabel_MaskedTB_Clear(); } // nichts eingegeben -> alles löschen
            else if (digits.Length > 0 && digits.Length < 8)
            {
                FormatAndSetDate();
                if (DateOnly.TryParseExact(maskedTextBox.Text, "dd.MM.yyyy", out var geburtsdatum)) { AgeLabel_MaskedTB_Set(geburtsdatum); }
                else { AgeLabel_MaskedTB_Clear(); }
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
        ignoreTextChange = false;
        maskedTextBox.Focus(); // Fokus setzen, damit TextChanged-Event ausgelöst wird
        maskedTextBox.Clear();
        UpdateSaveButton(); // Status aktualisieren, da das TextChanged-Event unterdrückt wurde
    }

    private void TextBox_TextChanged(object sender, EventArgs e)
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
            //saveFileDialog.InitialDirectory = string.IsNullOrEmpty(sDatabaseFolder) || !Directory.Exists(sDatabaseFolder) ? null : sDatabaseFolder;
            saveFileDialog.InitialDirectory = string.IsNullOrEmpty(_settings.DatabaseFolder) || !Directory.Exists(_settings.DatabaseFolder) ? null : _settings.DatabaseFolder;
            saveFileDialog.DefaultExt = "adb";
            saveFileDialog.Filter = "Adressen-Datenbank (*.adb)|*.adb|Alle Dateien (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (addressBindingSource != null) { await SaveSQLDatabaseAsync(true); }
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
        // 1. Initialisierung
        // Wir übergeben die statischen Standardwerte an das Formular (für den "Standard"-Button darin).
        using var frm = new FrmColumns(AppSettings.DefaultHideColumns);

        var columnList = frm.GetColumnList();
        var itemCount = columnList.Items.Count;

        // 2. Aktuellen Status aus _settings in die Checkboxen laden
        // Wir prüfen sicherheitshalber die Länge, um IndexOutOfRange zu vermeiden
        var limit = Math.Min(itemCount, _settings.HideColumnArr.Length);
        for (var i = 0; i < limit; i++)
        {
            // Checked bedeutet "Sichtbar", also das Gegenteil von "Hide"
            columnList.Items[i].Checked = !_settings.HideColumnArr[i];
        }

        // 3. Auswertung bei OK
        if (frm.ShowDialog() == DialogResult.OK)
        {
            var newHideArr = new bool[itemCount];

            // GUI-Status in bool-Array wandeln
            for (var i = 0; i < itemCount; i++)
            {
                newHideArr[i] = !columnList.Items[i].Checked; // !Checked = Hidden
            }

            // Settings aktualisieren
            _settings.HideColumnArr = newHideArr;

            // Auf beide Grids anwenden (Helper nutzen!)
            ApplyColumnSettings(addressDGV);
            ApplyColumnSettings(contactDGV);

            // Speichern
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

        // 1. Klon erstellen (für sauberes Abbrechen)
        var tempSettings = _settings.DeepClone();

        // 2. Form mit Settings-Objekt initialisieren
        // Hinweis: FrmCopyScheme muss angepasst werden (siehe unten)
        using var frm = new FrmCopyScheme(tempSettings, bookmarkTextDictionary);

        if (frm.ShowDialog() == DialogResult.OK)
        {
            // 3. Bei OK: Die geänderten Settings übernehmen und speichern
            // Die Konvertierung der Listen in Arrays ist bereits im Dialog passiert.
            _settings = tempSettings;
            SettingsManager.Save(_settings, _settingsPath);
        }
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
            if (addressDGV.CurrentRow != null && !FormStateManager.RowIsVisible(addressDGV, addressDGV.CurrentRow))
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
            if (contactDGV.CurrentRow != null && !FormStateManager.RowIsVisible(contactDGV, contactDGV.CurrentRow))
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
        // 1. Google Kontakte Logik (Snapshot) - bleibt wie sie ist
        if (addressBindingSource.Current is Contact currentContact)
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

            addressBindingSource.ResetBindings(false);
        }
        // 2. Lokale EF Core Adressen (Verbessert)
        else if (_context is not null)
        {
            var analysis = DbChangeAnalyzer.AnalyzeChanges(_context);
            if (!analysis.HasChanges) { return; }

            // --- NEU: Sicherheitsabfrage mit TaskDialog ---
            // Wir nutzen analysis.DialogText für die Liste der Namen, definieren aber eine eigene Heading für das Verwerfen.
            var confirmHeading = "Möchten Sie diese Änderungen wirklich unwiderruflich verwerfen?";

            var (isYes, _, _) = Utils.YesNo_TaskDialog(
                this,                                // Owner
                "Änderungen rückgängig machen",      // Caption
                confirmHeading,                      // Heading
                analysis.DialogText,                 // Text (Liste der Namen aus dem Analyzer)
                "Änderungen verwerfen",              // Custom "Ja"-Button
                "Abbrechen",                         // Custom "Nein"-Button
                false                                // Default Button auf "Nein" (false) für Sicherheit
            );

            if (!isYes) { return; }
            // ---------------------------------------------


            // 1. Merker setzen
            var topRowIndex = addressDGV.FirstDisplayedScrollingRowIndex;
            // Wir merken uns die ID, da das Objekt-Handle nach dem Revert/Reload manchmal instabil ist
            var currentId = (addressBindingSource.Current as Adresse)?.Id;

            // 2. Änderungen in EF rückgängig machen
            await DbChangeAnalyzer.RevertChangesAsync(analysis.RealChanges);

            // EXTRA: Den State-Tracker beruhigen
            // Wir erzwingen, dass EF Core die Einträge als "Unchanged" ansieht, 
            // damit beim Beenden kein falscher Alarm kommt.
            foreach (var entry in _context.ChangeTracker.Entries().Where(x => x.State != EntityState.Unchanged))
            {
                entry.State = EntityState.Unchanged;
            }

            // 3. UI-Refresh & Re-Sort
            SuspendLayout();
            addressBindingSource.RaiseListChangedEvents = false;
            addressDGV.DataSource = null;

            //var sortedLocalList = _context.Adressen.Local
            //    .OrderBy(a => a.Nachname)
            //    .ThenBy(a => a.Vorname)
            //    .ToList();

            //addressBindingSource.DataSource = new BindingList<Adresse>(sortedLocalList); // FALSCH: Das erstellt eine Kopie der Liste und trennt die Verbindung zu EF Core Local

            // RICHTIG:
            addressBindingSource.DataSource = _context.Adressen.Local.ToBindingList();
            // Falls Sie sortieren müssen, sortieren Sie Local vorher nicht neu (das geht nicht), 
            // sondern nutzen Sie die Sortierfunktion des DataGridViews oder eine BindingSource-Sortierung, 
            // aber EF Core Local ist per Definition unsortiert (Einfügereihenfolge).

            addressDGV.DataSource = addressBindingSource;
            addressBindingSource.RaiseListChangedEvents = true;
            addressBindingSource.ResetBindings(true);
            ResumeLayout();

            // 4. Selektion und Scroll-Position wiederherstellen
            //if (currentId.HasValue)
            //{
            //    var item = sortedLocalList.FirstOrDefault(a => a.Id == currentId.Value);
            //    if (item != null)
            //    {
            //        var newIndex = addressBindingSource.IndexOf(item);
            //        if (newIndex >= 0 && newIndex < addressDGV.RowCount)
            //        {
            //            addressBindingSource.Position = newIndex;

            //            // SICHERHEITS-CHECK für die CurrentCell
            //            // Wir suchen die erste sichtbare Spalte, um den "unsichtbare Zelle" Fehler zu vermeiden
            //            var firstVisibleCol = addressDGV.Columns
            //                .Cast<DataGridViewColumn>()
            //                .FirstOrDefault(c => c.Visible);

            //            if (firstVisibleCol != null)
            //            {
            //                try
            //                {
            //                    addressDGV.CurrentCell = addressDGV.Rows[newIndex].Cells[firstVisibleCol.Index];
            //                    addressDGV.Rows[newIndex].Selected = true;
            //                }
            //                catch (InvalidOperationException)
            //                {
            //                    // Falls das Grid noch im "Reset" Modus ist, ignorieren wir das Setzen der Zelle
            //                }
            //            }
            //        }
            //    }
            //}
            //if (topRowIndex >= 0 && topRowIndex < addressDGV.RowCount)
            //{
            //    try { addressDGV.FirstDisplayedScrollingRowIndex = topRowIndex; } catch { }
            //}
        }
        UpdateSaveButton();
        UpdateAddressStatusBar();
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
        saveFileDialog.Title = "Wählen Sie einen Speicherort";
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
                        //ConnectSQLDatabase(backupPath);
                        await ConnectSQLDatabaseAsync(backupPath);
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
                if (addressBindingSource != null)
                {
                    // Jetzt funktioniert await, weil das Lambda async ist
                    await SaveSQLDatabaseAsync(true);
                }

                // ConnectSQLDatabase wird erst ausgeführt, wenn SaveSQLDatabaseAsync fertig ist
                //ConnectSQLDatabase(file);
                await ConnectSQLDatabaseAsync(file);
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

        foreach (ListViewItem item in dokuListView.Items)
        {
            if (!existingDbPaths.Contains(item.Text))
            {
                selectedAddress.Dokumente.Add(new Dokument
                {
                    Dateipfad = item.Text,
                    AdressId = selectedAddress.Id,
                    Adresse = selectedAddress
                });
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
        //NativeMethods.SendMessage(searchTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, 8 << 16);
        //NativeMethods.SendMessage(searchTextBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, 8);
        searchTSTextBox.TextBox.SetInnerMargins(8, 8);

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
        if (_isFiltering || IsDisposed) { return; }
        if (InvokeRequired)
        {
            BeginInvoke(UpdateSaveButton);
            return;
        }

        // Statische Helfer-Methode für die Titel-Logik
        static void UpdateTabTitle(TabPage tab, string baseTitle, bool hasChanges)
        {
            var targetText = hasChanges ? $"{baseTitle}*" : baseTitle;
            if (tab.Text != targetText) { tab.Text = targetText; }
        }

        // 1. Wir berechnen die Zustände für beide Datenquellen
        var hasSqlChanges = HasRealEFChanges();
        var hasGoogleChanges = HasRealContactChanges(_lastActiveContact, _originalContactSnapshot);

        // 2. Wir aktualisieren IMMER beide Tab-Titel (UX-Bonus!)
        // Nutze hier exakt die Texte, die du auch im Designer vergeben hast
        UpdateTabTitle(addressTabPage, " Lokale Adressen", hasSqlChanges);
        UpdateTabTitle(contactTabPage, " Google Kontakte", hasGoogleChanges);

        // 3. Den Save-Button in der Toolbar schalten wir je nach aktivem Tab
        if (tabControl.SelectedTab == addressTabPage)
        {
            saveTSButton.Enabled = hasSqlChanges;
            saveTSButton.ToolTipText = hasSqlChanges ? "Lokale Adressen speichern (Strg+S)" : "Keine Änderungen";
        }
        else if (tabControl.SelectedTab == contactTabPage)
        {
            saveTSButton.Enabled = hasGoogleChanges;
            saveTSButton.ToolTipText = hasGoogleChanges ? "Google-Kontakt hochladen (Strg+S)" : "Keine Änderungen";
        }
    }

    private bool HasRealEFChanges()
    {
        if (_context == null) { return false; }

        // WICHTIG: DetectChanges erkennt Added/Deleted bei Beziehungen automatisch
        _context.ChangeTracker.DetectChanges();

        foreach (var entry in _context.ChangeTracker.Entries())
        {
            // Fall 1: Hinzugefügte oder Gelöschte Objekte
            // Das deckt ab:
            // - Neue Adressen
            // - Gelöschte Adressen
            // - Neue Gruppen-Verknüpfungen (Schatten-Entitäten) -> Das ist Ihr Fall beim Hinzufügen!
            // - Gelöschte Gruppen-Verknüpfungen -> Das ist Ihr Fall beim Löschen!
            if (entry.State == EntityState.Added || entry.State == EntityState.Deleted)
            {
                return true;
            }

            // Fall 2: Modifizierte Objekte (z.B. Tippfehler-Korrektur)
            if (entry.State == EntityState.Modified)
            {
                // Wenn es eine Schatten-Entität ist (keine C#-Klasse), ist Modified eher selten, aber möglich.
                // Wir behandeln es wie eine echte Änderung.
                if (entry.Metadata.ClrType == null) { return true; }

                // Bei echten Klassen prüfen wir die Properties auf relevante Änderungen
                foreach (var prop in entry.Properties)
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
        if (_lastActiveContact == null || _originalContactSnapshot == null) { return true; }  // kein Kontakt aktiv oder kein Snapshot -> alles okay


        contactDGV.CausesValidation = false;
        addressDGV.CausesValidation = false;
        try
        {
            if (!ValidateChildren(ValidationConstraints.Enabled)) { return false; }  // Validierung der Textboxen/Comboboxen fehlgeschlagen
            contactBindingSource.EndEdit();  // Daten vom UI in die BindingSource schreiben
        }
        finally
        {
            contactDGV.CausesValidation = true;
            addressDGV.CausesValidation = true;
        }

        var currentContact = _lastActiveContact;  // Lokale Referenz für den Compiler (wegen Nullable-Check)
        var isNewContact = string.IsNullOrEmpty(currentContact.ResourceName);
        if (!HasRealContactChanges(currentContact, _originalContactSnapshot))
        {
            if (isNewContact) { RemoveContactFromList(currentContact); }  // Der User hat "Neu" geklickt, aber NICHTS getippt und wechselt jetzt die Zeile.
            return true; // Er darf wechseln
        }
        var result = await AskSaveContactChangesAsync(isClosing);  // Bei DialogResult.Yes wurde bereits in AskSaveContactChangesAsync gespeichert.
        if (result == DialogResult.Cancel) { return false; }  // Bleiben (Abbruch)
        if (result == DialogResult.No)
        {
            if (isNewContact) { RemoveContactFromList(currentContact); }  // Er ist neu UND der User will ihn nicht speichern -> Weg damit.
            else
            {
                currentContact.CopyFrom(_originalContactSnapshot);
                if (!isClosing) { contactBindingSource.ResetCurrentItem(); }  // Er existierte schon -> Änderungen rückgängig machen
            }
        }
        return true;
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
        var headingText = string.IsNullOrWhiteSpace(fullName)
            ? "Möchten Sie die Änderungen speichern?"
            : $"Möchten Sie die Änderungen an {fullName} speichern?";
        var fieldList = string.Join("\n", changedFields.Select(f => "• " + char.ToUpper(f[0]) + f[1..]));
        //if (photoChanged) { fieldList += "\n• Foto"; }
        var shortSummary = $"{changedFields.Count} Bereich(e) wurden geändert.\n{fieldList}";
        var detailedDiff = Utils.GenerateDetailedDiff(_lastActiveContact, _originalContactSnapshot, dataFields);
        var btnSave = new TaskDialogButton("&Hochladen") { AllowCloseDialog = false }; // Wichtig: Schließt nicht sofort
        var btnDiscard = new TaskDialogButton("&Verwerfen");
        var btnCancel = TaskDialogButton.Cancel;
        var pageMain = new TaskDialogPage()
        {
            Caption = "Google Kontakte",
            Heading = headingText,
            Text = shortSummary + Environment.NewLine + detailedDiff,
            Icon = TaskDialogIcon.ShieldBlueBar,
            AllowCancel = true,
            Buttons = { btnSave, btnDiscard, btnCancel },
            DefaultButton = btnSave
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
                pageProgress.Heading = "Erfolgreich gespeichert.";
                pageProgress.Text = "Die Daten wurden synchronisiert.";

                // Kurze Pause für UX, dann schließen
                await Task.Delay(500);
                pageProgress.Buttons[0].Enabled = true;
                pageProgress.Buttons[0].PerformClick();
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
            // Wenn wir erfolgreich waren, müssen wir ggf. Buttons deaktivieren (außer wir schließen eh)
            if (!isClosing)
            {
                saveTSButton.Enabled = false;
                // Grid Fotozelle aktualisieren, falls Foto neu (per BindingSource Reset oder direkt)
                contactBindingSource.ResetBindings(false);
            }
            return DialogResult.Yes;
        }
        if (clickedButton == btnDiscard) { return DialogResult.No; }
        return DialogResult.Cancel;
    }

    private async Task ExecuteGoogleSaveAsync(Contact contactToSave, List<string> changedFields, Image? currentImage, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();  // 1. Initialisierung
        var manager = new GooglePeopleManager(secretPath, tokenDir);
        if (string.IsNullOrEmpty(contactToSave.ResourceName))  // 2. Weiche: Erstellen oder Aktualisieren
        {
            var createdContact = await manager.CreateContactAsync(contactToSave, currentImage, token);  // übernimmt Kontakt + Bild in einem Rutsch 
            contactToSave.ResourceName = createdContact.ResourceName;
            contactToSave.ETag = createdContact.ETag;
            contactToSave.PhotoUrl = createdContact.PhotoUrl;
        }
        else  // === FALL B: UPDATE ===
        {
            if (changedFields.Count > 0 || changedFields.Contains("memberships"))
            {
                var updatedPerson = await manager.UpdateContactAsync(contactToSave, changedFields, contactGroupsDict, _originalContactSnapshot, checkEmptyGroups: true, token: token);
                contactToSave.ETag = updatedPerson.ETag;
                contactToSave.ResourceName = updatedPerson.ResourceName;
            }
            token.ThrowIfCancellationRequested();  // Token prüfen, bevor das teure Bild-Update startet
        }
        _originalContactSnapshot = (Contact)contactToSave.Clone();  // Lokalen Status konsolidieren
        contactToSave.ResetSearchCache();  // gehört streng genommen eigentlich in den Aufrufer (Click-Event), aber wenn es hier steht, ist sichergestellt, dass es nur bei Erfolg passiert.
    }

    private void RemoveContactFromList(Contact contact)
    {
        isSelectionChanging = true;  // Verhindert Events während des Löschens
        try
        {
            _allGoogleContacts.Remove(contact); // contactBindingSource.Remove(contact);  // Falls Sie die BindingList direkt nutzen, reicht das Remove oben.
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
                if (activeDgv.Rows.Count > newIndex) { activeDgv.FirstDisplayedScrollingRowIndex = newIndex; }
            }
        }
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
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV, a => a.Dokumente.Count != 0, "… mit Briefverweis", "Adressen");
        }
    }

    private void PhotoPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // 1. Wir fragen die DB: Welche IDs haben ein Foto? ("SELECT Id FROM Adressen WHERE FotoId IS NOT NULL")
            var idsWithPhoto = _context.Adressen.Where(a => a.Foto != null).Select(a => a.Id).ToHashSet(); // HashSet für extrem schnelle Suche
            // 2. Wir filtern die lokale Liste anhand dieser IDs
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV, a => idsWithPhoto.Contains(a.Id), "… mit Bild", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV, c => !string.IsNullOrWhiteSpace(c.PhotoUrl), "… mit Bild", "Google Kontakte");
        }
    }

    private void PhotoMinusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            // 1. Gleiches Spiel: IDs holen
            var idsWithPhoto = _context.Adressen.Where(a => a.Foto != null).Select(a => a.Id).ToHashSet();
            // 2. Filter umdrehen: Zeige alle, deren ID NICHT in der Liste ist
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV, a => !idsWithPhoto.Contains(a.Id), "… ohne Bild", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV, c => string.IsNullOrWhiteSpace(c.PhotoUrl), "… ohne Bild", "Google Kontakte");
        }
    }

    private void MailPlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV, a => !string.IsNullOrWhiteSpace(a.Mail1) || !string.IsNullOrWhiteSpace(a.Mail2), "… mit E-Mail", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV, c => !string.IsNullOrWhiteSpace(c.Mail1) || !string.IsNullOrWhiteSpace(c.Mail2), "… mit E-Mail", "Google Kontakte");
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

    private void TelephonePlusFilterToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (tabControl.SelectedTab == addressTabPage && _context != null)
        {
            ExecuteFilter(_context.Adressen.Local, addressBindingSource, addressDGV,
                a => !string.IsNullOrWhiteSpace(a.Telefon1) || !string.IsNullOrWhiteSpace(a.Telefon2), "… mit Telefonnummer", "Adressen");
        }
        else if (tabControl.SelectedTab == contactTabPage && _allGoogleContacts != null)
        {
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
            ExecuteFilter(_allGoogleContacts, contactBindingSource, contactDGV,
                c => string.IsNullOrWhiteSpace(c.Telefon1) && string.IsNullOrWhiteSpace(c.Telefon2), "… ohne Telefonnummer", "Google Kontakte");
        }
    }

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

    private void FilterlToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
    {
        var isAddressTab = tabControl.SelectedTab == addressTabPage && addressDGV.Rows.Count > 0;
        var isContactTab = tabControl.SelectedTab == contactTabPage && contactDGV.Rows.Count > 0;
        var enableCommon = isAddressTab || isContactTab;
        foreach (ToolStripItem item in filterlToolStripMenuItem.DropDownItems)
        {
            if (item == adressenMitBriefToolStripMenuItem) { item.Enabled = isAddressTab; }
            else if (item is ToolStripMenuItem) { item.Enabled = enableCommon; }
        }
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

    private void TopAlignZoomPictureBox_DoubleClick(object sender, EventArgs e) => AddPictboxToolStripButton_Click(topAlignZoomPictureBox, EventArgs.Empty);

    private void FilterRemoveToolStripMenuItem_Click(object sender, EventArgs e)
    {
        ignoreSearchChange = true;
        searchTSTextBox.TextBox.Clear();
        tsClearLabel.Visible = false;
        ignoreSearchChange = false;

        if (tabControl.SelectedTab == addressTabPage)
        {
            if (_context == null) { return; }
            ExecuteAndPreserveSelection<Adresse>(addressBindingSource, addressDGV, () => { addressBindingSource.DataSource = _context.Adressen.Local.ToBindingList(); });
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

    private async void AddPictboxToolStripButton_Click(object sender, EventArgs e)
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

        if (openFileDialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        // ---------------------------------------------------------
        // FALL 1: Lokale Datenbank (EF Core)
        // ---------------------------------------------------------
        if (tabControl.SelectedTab == addressTabPage)
        {
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

                    if (loadedImage.Width > 100)
                    {
                        scaledImage = Utils.SkaliereBildDaten(loadedImage, 100);
                        finalImage = scaledImage;
                    }
                    else
                    {
                        finalImage = loadedImage;
                    }

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

                    addressBindingSource.ResetCurrentItem();
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
        // FALL 2: Google Kontakte
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

                if (contactBindingSource.Current is not Contact currentContact)
                {
                    return;
                }

                // PRÜFUNG: Ist es ein neuer Kontakt?
                var isNewContact = string.IsNullOrEmpty(currentContact.ResourceName);

                // Button-Text anpassen: "Hochladen" für existierende, "Übernehmen" für neue
                var buttonText = isNewContact ? "Übernehmen" : "Hochladen";
                var initialButtonYes = new TaskDialogButton(buttonText) { AllowCloseDialog = isNewContact }; // Bei 'Neu' darf der Dialog schließen, bei 'Update' geht es zur ProgressPage
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

                var replaceWarning = string.Empty;
                if (!string.IsNullOrEmpty(currentContact.PhotoUrl))
                {
                    replaceWarning += $"Das vorhandene Foto wird überschrieben!\n\n";
                }

                // Text für "Neuer Kontakt" leicht anpassen, da noch kein Upload erfolgt
                var infoText = isNewContact
                    ? $"Information: Abmessung {workingImage.Width}×{workingImage.Height} Pixel.\nDas Bild wird erst beim Speichern des Kontakts hochgeladen.{caveText}"
                    : $"{replaceWarning}Information: Abmessung {workingImage.Width}×{workingImage.Height} Pixel.{caveText}";

                var initialPage = new TaskDialogPage()
                {
                    Caption = "Google Kontakte",
                    Heading = isNewContact ? "Foto übernehmen?" : "Möchten Sie die Änderung speichern?",
                    Text = infoText,
                    Icon = new(Resources.question32),
                    AllowCancel = true,
                    SizeToContent = true,
                    Buttons = { initialButtonNo, initialButtonYes }
                };

                foreach (var rb in radioButtons)
                {
                    initialPage.RadioButtons.Add(rb);
                }

                // Progress Page nur nötig, wenn wir direkt hochladen (kein neuer Kontakt)
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
                        intermediateImageToDispose = workingImage;
                        workingImage = Utils.BeschneideZuQuadrat(workingImage, false);
                        finalImageToUpload = workingImage;
                        finalImageForDisplay = (Image)workingImage.Clone();
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
                        finalImageToUpload = workingImage;
                        finalImageForDisplay = Utils.ReduziereWieGoogle(workingImage, 100);
                    }
                    else  // Fall: Keine RadioButtons (Bild war nicht hochkant)
                    {
                        finalImageToUpload = workingImage;
                        finalImageForDisplay = (Image)workingImage.Clone();
                    }

                    // Bild in UI setzen
                    topAlignZoomPictureBox.Image = finalImageForDisplay;
                    intermediateImageToDispose?.Dispose();

                    // UNTERSCHEIDUNG LOGIK
                    if (isNewContact) { delPictboxToolStripButton.Enabled = true; }  // Das Bild ist jetzt in der PictureBox und wird später von ExecuteGoogleSaveAsync abgeholt.
                    else { initialPage.Navigate(progressPage); }  // Existierender Kontakt: Weiterleitung zum Upload
                };
                if (!isNewContact)  // Event nur abonnieren, wenn wir wirklich hochladen wollen
                {
                    progressPage.Created += async (s, e) =>
                    {
                        try
                        {
                            await UpdateContactPhotoAsync(currentContact, finalImageToUpload!, origImgFormat, () => progressPage.Buttons.First().PerformClick());
                        }
                        finally
                        {
                            workingImage?.Dispose();
                        }
                    };
                }

                TaskDialog.ShowDialog(Handle, initialPage);

                // Nach dem Dialog (synchroner Call für ShowDialog):
                if (isNewContact && initialButtonYes.Enabled == true) // Wenn wir "Ja" geklickt haben (Button ist enabled)
                {
                    // Ggf. cleanup hier, falls nötig. workingImage wird im 'finally' unten disposed.
                }
            }
            catch (Exception ex)
            {
                Utils.MsgTaskDlg(Handle, $"Fehler beim Laden: {ex.GetType()}", $"Bild konnte nicht geladen werden: {ex.Message}", TaskDialogIcon.Error);
                finalImageForDisplay?.Dispose();
            }
            finally { workingImage?.Dispose(); }  // Bei neuem Kontakt wurde workingImage noch nicht im ProgressPage-Event disposed
        }
    }

    private async void DelPictboxToolStripButton_Click(object sender, EventArgs e)
    {
        // --- FALL A: SQL ADRESSEN ---
        if (tabControl.SelectedTab == addressTabPage && addressBindingSource.Current is Adresse adresse)
        {
            var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Adressen", "Möchten Sie das Bild entfernen?",
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

                        addressBindingSource.ResetCurrentItem();
                        UpdateSaveButton();
                    }
                }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
        }
        // --- FALL B: GOOGLE KONTAKTE ---
        else if (tabControl.SelectedTab == contactTabPage && contactBindingSource.Current is Contact googleKontakt)
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
            var (isYes, _, _) = Utils.YesNo_TaskDialog(this, "Google Kontakte", "Möchten Sie das Bild wirklich löschen?",
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
                    contactBindingSource.ResetCurrentItem();
                }
                catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
            }
        }
    }

    private async Task<bool> CopyLocalToGoogleAsync(Adresse localAddress)
    {
        Contact? createdContact = null;
        var newGoogleContact = new Contact();
        try  // Datenmapping (findet noch auf dem UI Thread statt, da sehr schnell)
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
        }
        catch (Exception ex)  // Fehler beim Mapping fangen wir hier direkt ab
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
                        var ms = new MemoryStream(localAddress.Foto.Fotodaten);  // MemoryStream darf nicht disposed werden, solange das Bitmap lebt.
                        imageToUpload = new Bitmap(ms);
                    }
                    catch { }  // Bild defekt - Upload läuft ohne Bild weiter
                }
                var manager = new GooglePeopleManager(secretPath, tokenDir);
                createdContact = await manager.CreateContactAsync(newGoogleContact, imageToUpload, token);
            }
            finally { imageToUpload?.Dispose(); }
        });
        if (success && createdContact != null)  // UI Update (nur bei Erfolg und vorhandenem Ergebnis)
        {
            try
            {
                _allGoogleContacts?.Add(createdContact);  // zur Hauptliste hinzufügen
                if (_allGoogleContacts != null)
                {
                    Utils.SortContacts(_allGoogleContacts);  // Sortierung anwenden
                    contactBindingSource.DataSource = _allGoogleContacts;  // BindingSource aktualisieren 
                    contactBindingSource.ResetBindings(false);
                    var newIndex = _allGoogleContacts.IndexOf(createdContact);
                    if (newIndex >= 0) { contactBindingSource.Position = newIndex; }
                }
                else
                {
                    contactBindingSource.Add(createdContact);  // Fallback
                    contactBindingSource.Position = contactBindingSource.Count - 1;
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
                // Property-Infos beider Klassen holen
                var sourceProp = sourceType.GetProperty(fieldName);
                var targetProp = targetType.GetProperty(fieldName);

                // Prüfung: Existiert das Feld in beiden Klassen und ist es beschreibbar?
                if (sourceProp != null && targetProp != null && targetProp.CanWrite)
                {
                    var value = sourceProp.GetValue(googleKontakt);
                    targetProp.SetValue(newLocalAddress, value);
                }
            }

            // -----------------------------------------------------------
            // 2. Foto separat laden (Speziallogik, nicht im Array)
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
            // 3. UI Update & Sortierung
            // -----------------------------------------------------------
            var insertIndex = Utils.GetAddressInsertIndex(addressBindingSource, newLocalAddress);
            addressBindingSource.Insert(insertIndex, newLocalAddress);
            addressBindingSource.Position = insertIndex;

            return true;
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            return false;
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
                            // 1. Verknüpfung entfernen (Erzeugt "Deleted" State bei der Schatten-Entität)
                            adresse.Gruppen.Remove(gruppeToDelete);
                            curAddressMemberships.Remove(membershipToRemove);

                            // 2. UI Aktualisieren
                            UpdateMembershipTags();
                            UpdateMembershipCBox();
                            UpdatePlaceholderVis();

                            // 3. WICHTIG: UI benachrichtigen (aktiviert Buttons, feuert Events)
                            addressBindingSource.ResetCurrentItem();

                            // 4. Save-Button explizit prüfen
                            UpdateSaveButton();
                        }
                    }
                }
            };
            flowLayoutPanel.Controls.Add(tagControl);
        }
        UpdatePlaceholderVis();
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
            UpdateMembershipCBox();
            UpdateMembershipJson(); // Google nutzt weiterhin JSON/Strings
        }
        else if (tabControl.SelectedTab == addressTabPage)
        {
            if (addressBindingSource.Current is Adresse adresse && _context != null) // _context Prüfung hier integriert
            {
                // 1. Sicherstellen, dass die Gruppen geladen sind
                var entry = _context.Entry(adresse);
                if (!entry.Collection(a => a.Gruppen).IsLoaded)
                {
                    entry.Collection(a => a.Gruppen).Load();

                    // NEU: WICHTIG! 
                    // Da wir Gruppen nachgeladen haben, müssen wir die UI-Listen synchronisieren.
                    // Sonst bricht der nächste Check ab, obwohl die Gruppe im FlowPanel noch fehlt.
                    LoadGroupsForCurrentAddress();
                }

                // 2. Prüfen, ob die Adresse die Gruppe schon hat
                // (Jetzt ist sichergestellt, dass auch die UI aktuell ist)
                if (adresse.Gruppen.Any(g => g.Name.Equals(newMembershipName, StringComparison.OrdinalIgnoreCase)))
                {
                    tagComboBox.SelectAll();
                    tagComboBox.Focus();
                    return; // Hier brechen wir ab - aber jetzt ist die UI bereits aktuell!
                }
                // 2. Gruppe in der DB suchen oder neu erstellen

                // A) Zuerst im ChangeTracker (Lokal) schauen
                // HIER ist StringComparison auch ERLAUBT (In-Memory)
                var gruppe = _context?.Gruppen.Local
                    .FirstOrDefault(g => g.Name.Equals(newMembershipName, StringComparison.OrdinalIgnoreCase));

                // B) Wenn nicht lokal, dann in der Datenbank suchen
                // HIER WAR DER FEHLER: EF Core kann StringComparison nicht nach SQL übersetzen.

                // Variante A (Beste Performance): Verlässt sich auf die DB-Einstellung (meist case-insensitive)
                gruppe ??= _context?.Gruppen.FirstOrDefault(g => g.Name == newMembershipName);

                if (gruppe == null)
                {
                    gruppe = new Gruppe { Name = newMembershipName };
                    _context?.Gruppen.Add(gruppe);
                    // Wichtig: Zur BindingList hinzufügen, damit die ComboBox es sofort kennt
                    allAddressMemberships.Add(newMembershipName);
                }

                adresse.Gruppen.Add(gruppe);  // Verknüpfung herstellen
                curAddressMemberships.Add(newMembershipName);

                //_context?.Entry(adresse).State = EntityState.Modified;  // Adresse als modifiziert markieren
                UpdateMembershipTags();
                UpdateMembershipCBox();
                addressBindingSource.ResetCurrentItem();
                UpdateSaveButton();
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
            else { tbNotizen.Focus(); }  // SelectNextControl((Control)sender, true, true, true, true); 
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
        if (Application.SystemColorMode == SystemColorMode.System) { _isDarkMode = DefaultBackColor.R < 128; } //falls die Automatik hakt
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

    private void AddressDGV_SelectionChanged(object sender, EventArgs e) => scrollTimer.Start();

    private void ScrollTimer_Tick(object sender, EventArgs e) => scrollTimer.Stop();

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
                    if (latestVersion > currentVersion) { footText = $"Update verfügbar: v{latestVersion} vom {releaseDate}"; }
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

}
