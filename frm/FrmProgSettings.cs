//using System.Diagnostics;
//using System.Drawing.Drawing2D;
//using Adressen.cls; // Namespace der AppSettings

//namespace Adressen;

//public partial class FrmProgSettings : Form
//{
//    private readonly AppSettings _settings;

//    // Konstruktor nimmt jetzt direkt die Settings entgegen
//    internal FrmProgSettings(AppSettings settings)
//    {
//        InitializeComponent();
//        _settings = settings;

//        // Binding initialisieren
//        InitializeDataBindings();

//        // Initiale UI-Status-Updates (Enabling/Disabling)
//        UpdateUiState();
//    }

//    private void InitializeDataBindings()
//    {
//        // --- CheckBoxen (Bool) ---
//        // DataSourceUpdateMode.OnPropertyChanged sorgt für sofortiges Schreiben in das Objekt
//        Bind(ckbAskBeforeDelete, "Checked", nameof(AppSettings.AskBeforeDelete));
//        Bind(ckbContactsAutoload, "Checked", nameof(AppSettings.ContactsAutoload));
//        Bind(ckbBackup, "Checked", nameof(AppSettings.DailyBackup));
//        Bind(ckbWatchFolder, "Checked", nameof(AppSettings.WatchFolder));
//        Bind(ckbAskBeforeSaveSQL, "Checked", nameof(AppSettings.AskBeforeSaveSQL));

//        // --- TextBoxen (String) ---
//        Bind(tbStandard, "Text", nameof(AppSettings.StandardFile));
//        Bind(tbBackupFolder, "Text", nameof(AppSettings.BackupDirectory));
//        Bind(tbDatabaseFolder, "Text", nameof(AppSettings.DatabaseFolder));
//        Bind(tbWatchFolder, "Text", nameof(AppSettings.DocumentFolder)); // Achtung: Property hieß im alten Code "LetterDirectory"

//        // --- RadioButtons (Komplexe Logik) ---

//        // 1. Start-Verhalten (Mapping: Bool Property -> 3 RadioButtons)
//        // Logik: 
//        // ReloadRecent == true -> rbRecent
//        // NoAutoload == true -> rbEmpty
//        // StandardFile gesetzt -> rbStandard
//        // Das ist etwas tricky zu binden, da es Abhängigkeiten gibt. 
//        // Hier ist eine Hybrid-Lösung oft stabiler als reines Binding:
//        if (_settings.ReloadRecent) { rbRecent.Checked = true; }
//        else if (_settings.NoAutoload) { rbEmpty.Checked = true; }
//        else { rbStandard.Checked = true; }

//        // Events für manuelle Rückschreibung der Start-Logik
//        rbRecent.CheckedChanged += (s, e) => { if (rbRecent.Checked) { _settings.ReloadRecent = true; _settings.NoAutoload = false; } };
//        rbEmpty.CheckedChanged += (s, e) => { if (rbEmpty.Checked) { _settings.ReloadRecent = false; _settings.NoAutoload = true; } };
//        rbStandard.CheckedChanged += (s, e) => { if (rbStandard.Checked) { _settings.ReloadRecent = false; _settings.NoAutoload = false; } };


//        // 2. Farbschema (Mapping: String -> RadioButtons)
//        BindRadio(rbtnBlue, nameof(AppSettings.ColorScheme), "blue");
//        BindRadio(rbtnDark, nameof(AppSettings.ColorScheme), "dark");
//        BindRadio(rbtnPale, nameof(AppSettings.ColorScheme), "pale");
//        // Fallback für "grey" oder unbekannt könnte man optional behandeln, wird aber durch Default selection abgedeckt.

//        // 3. Textverarbeitung (Mapping: Bool? -> RadioButtons)
//        BindRadio(rbMSWord, nameof(AppSettings.WordProcessorProgram), true);
//        BindRadio(rbLibreOffice, nameof(AppSettings.WordProcessorProgram), false);
//        BindRadio(rbManualSelect, nameof(AppSettings.WordProcessorProgram), null);
//    }

//    /// <summary>
//    /// Helfer für einfaches Property-Binding
//    /// </summary>
//    private void Bind(Control control, string propertyName, string dataMember) => control.DataBindings.Add(propertyName, _settings, dataMember, false, DataSourceUpdateMode.OnPropertyChanged);

//    /// <summary>
//    /// Bindet einen RadioButton an einen bestimmten Wert einer Property.
//    /// </summary>
//    /// <param name="rb">Der RadioButton</param>
//    /// <param name="dataMember">Name der Property in AppSettings</param>
//    /// <param name="targetValue">Der Wert, den die Property haben muss, damit dieser RB checked ist (z.B. "blue" oder true)</param>
//    private void BindRadio(RadioButton rb, string dataMember, object? targetValue)
//    {
//        var binding = new Binding("Checked", _settings, dataMember, true, DataSourceUpdateMode.OnPropertyChanged);

//        // Format: Von Datenquelle (string/bool?) zum UI (bool Checked)
//        binding.Format += (s, e) =>
//        {
//            if (e.Value == null && targetValue == null) { e.Value = true; }
//            else if (e.Value != null && e.Value.Equals(targetValue)) { e.Value = true; }
//            else { e.Value = false; }
//        };

//        // Parse: Vom UI (bool Checked) zur Datenquelle (string/bool?)
//        binding.Parse += (s, e) =>
//        {
//            if ((bool)e.Value!) { e.Value = targetValue; }
//        };

//        rb.DataBindings.Add(binding);
//    }

//    // UI-Status aktualisieren (Enabled/Disabled Logik)
//    private void UpdateUiState()
//    {
//        tbStandard.Enabled = btnStandardFile.Enabled = rbStandard.Checked;

//        var backupActive = ckbBackup.Checked;
//        tbBackupFolder.Enabled = btnBackupFolder.Enabled = backupActive;
//        btnExplorer.Enabled = backupActive && !string.IsNullOrEmpty(tbBackupFolder.Text);

//        var watchActive = ckbWatchFolder.Checked;
//        tbWatchFolder.Enabled = btnWatchFolder.Enabled = lblWatchFolder.Enabled = watchActive;
//    }

//    // --- Event Handler (nur noch für UI-Logik, nicht Datentransfer) ---

//    private void FrmProgSettings_Load(object sender, EventArgs e)
//    {
//        // Initialer Check, falls StandardFile leer ist aber ausgewählt wurde
//        if (rbStandard.Checked && string.IsNullOrEmpty(tbStandard.Text)) { rbEmpty.Checked = true; }
//    }

//    private void RbStandard_CheckedChanged(object sender, EventArgs e) => UpdateUiState();
//    private void CkbBackup_CheckedChanged(object sender, EventArgs e) => UpdateUiState();
//    private void TbBackupFolder_TextChanged(object sender, EventArgs e) => UpdateUiState(); // Explorer Button aktivieren
//    private void CkbWatchFolder_CheckedChanged(object sender, EventArgs e) => UpdateUiState();


//    // --- File Dialog Buttons ---
//    // Da die Textboxen gebunden sind, reicht es, die Text-Eigenschaft zu setzen. 
//    // Das Binding aktualisiert automatisch das Settings-Objekt.

//    private void BtnStandardFile_Click(object sender, EventArgs e)
//    {
//        openFileDialog.InitialDirectory = !string.IsNullOrEmpty(tbStandard.Text) ? Path.GetDirectoryName(tbStandard.Text) : null;
//        if (openFileDialog.ShowDialog() == DialogResult.OK) { tbStandard.Text = openFileDialog.FileName; }
//    }

//    private void BtnDatabaseFolder_Click(object sender, EventArgs e)
//    {
//        if (folderBrowserDialog.ShowDialog() == DialogResult.OK) { tbDatabaseFolder.Text = folderBrowserDialog.SelectedPath; }
//    }

//    private void BtnBackupFolder_Click(object sender, EventArgs e)
//    {
//        folderBrowserDialog.InitialDirectory = Directory.Exists(tbBackupFolder.Text) ? tbBackupFolder.Text : string.Empty;
//        if (folderBrowserDialog.ShowDialog() == DialogResult.OK) { tbBackupFolder.Text = folderBrowserDialog.SelectedPath; }
//    }

//    private void BtnWatchFolder_Click(object sender, EventArgs e)
//    {
//        if (folderBrowserDialog.ShowDialog() == DialogResult.OK) { tbWatchFolder.Text = folderBrowserDialog.SelectedPath; }
//    }

//    private void BtnExplorer_Click(object sender, EventArgs e)
//    {
//        if (Directory.Exists(tbBackupFolder.Text))
//        {
//            using var process = new Process();
//            process.StartInfo.FileName = tbBackupFolder.Text;
//            process.StartInfo.UseShellExecute = true;
//            process.Start();
//        }
//        else { Console.Beep(); }
//    }

//    // --- Standard Form Kram ---

//    private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
//    {
//        // ... (Ihr bestehender DrawItem Code, unverändert) ...
//        var g = e.Graphics;
//        g.SmoothingMode = SmoothingMode.HighQuality;
//        var tabPage = tabControl.TabPages[e.Index];
//        var tabBounds = tabControl.GetTabRect(e.Index);
//        var backBrush = e.State == DrawItemState.Selected ? SystemBrushes.GradientActiveCaption : SystemBrushes.GradientInactiveCaption;
//        var textBrush = e.State == DrawItemState.Selected ? SystemBrushes.HighlightText : SystemBrushes.ControlText;
//        g.FillRectangle(backBrush, e.Bounds);
//        using var tabFont = new Font("Segoe UI", 10f);
//        using var stringFlags = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };
//        g.DrawString(tabPage.Text, tabFont, textBrush, tabBounds, stringFlags);
//        if (e.Index == tabControl.TabCount - 1)
//        {
//            var totalTabHeight = tabBounds.Height * tabControl.TabCount;
//            var remainingRect = new Rectangle(0, totalTabHeight, tabBounds.Width + 2, tabControl.Height - totalTabHeight);
//            g.FillRectangle(SystemBrushes.Control, remainingRect);
//        }
//    }

//    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
//    {
//        if (keyData == Keys.Escape) { Close(); return true; }
//        if (keyData == Keys.Tab) { tabControl.SelectedIndex = (tabControl.SelectedIndex + 1) % tabControl.TabCount; return true; }
//        return base.ProcessCmdKey(ref msg, keyData);
//    }
//}

using System.Diagnostics;
using System.Drawing.Drawing2D;
using Adressen.cls;

namespace Adressen;

public partial class FrmProgSettings : Form
{
    private readonly AppSettings _settings;

    // Konstruktor nimmt direkt den Klon der Settings entgegen
    internal FrmProgSettings(AppSettings settings)
    {
        InitializeComponent();
        _settings = settings;

        // 1. Werte aus dem Objekt in die Maske laden
        MapSettingsToUi();

        // 2. Initiale UI-Status-Updates (Enabling/Disabling von Feldern)
        UpdateUiState();
    }

    // --- MAPPING LOGIK (Manuell statt DataBinding) ---

    private void MapSettingsToUi()
    {
        // CheckBoxen
        ckbAskBeforeDelete.Checked = _settings.AskBeforeDelete;
        ckbContactsAutoload.Checked = _settings.ContactsAutoload;
        ckbBackup.Checked = _settings.DailyBackup;
        ckbWatchFolder.Checked = _settings.WatchFolder;
        ckbAskBeforeSaveSQL.Checked = _settings.AskBeforeSaveSQL;

        // TextBoxen
        tbStandard.Text = _settings.StandardFile;
        tbBackupFolder.Text = _settings.BackupDirectory;
        tbDatabaseFolder.Text = _settings.DatabaseFolder;
        tbWatchFolder.Text = _settings.DocumentFolder;

        // RadioButtons: Start-Verhalten
        if (_settings.ReloadRecent)
        {
            rbRecent.Checked = true;
        }
        else if (_settings.NoAutoload)
        {
            rbEmpty.Checked = true;
        }
        else
        {
            rbStandard.Checked = true;
        }

        // RadioButtons: Farbschema
        switch (_settings.ColorScheme)
        {
            case "dark": rbtnDark.Checked = true; break;
            case "pale": rbtnPale.Checked = true; break;
            default: rbtnBlue.Checked = true; break; // Fallback & "blue"
        }

        // RadioButtons: Textverarbeitung (bool?)
        if (_settings.WordProcessorProgram == true) { rbMSWord.Checked = true; }
        else if (_settings.WordProcessorProgram == false) { rbLibreOffice.Checked = true; }
        else { rbManualSelect.Checked = true; } // null
    }

    private void MapUiToSettings()
    {
        // CheckBoxen
        _settings.AskBeforeDelete = ckbAskBeforeDelete.Checked;
        _settings.ContactsAutoload = ckbContactsAutoload.Checked;
        _settings.DailyBackup = ckbBackup.Checked;
        _settings.WatchFolder = ckbWatchFolder.Checked;
        _settings.AskBeforeSaveSQL = ckbAskBeforeSaveSQL.Checked;

        // TextBoxen
        _settings.StandardFile = tbStandard.Text.Trim();
        _settings.BackupDirectory = tbBackupFolder.Text.Trim();
        _settings.DatabaseFolder = tbDatabaseFolder.Text.Trim();
        _settings.DocumentFolder = tbWatchFolder.Text.Trim();

        // Start-Verhalten Logik
        if (rbRecent.Checked)
        {
            _settings.ReloadRecent = true;
            _settings.NoAutoload = false;
        }
        else if (rbEmpty.Checked)
        {
            _settings.ReloadRecent = false;
            _settings.NoAutoload = true;
        }
        else // rbStandard.Checked
        {
            _settings.ReloadRecent = false;
            _settings.NoAutoload = false;
        }

        // Farbschema
        if (rbtnDark.Checked) { _settings.ColorScheme = "dark"; }
        else if (rbtnPale.Checked) { _settings.ColorScheme = "pale"; }
        else { _settings.ColorScheme = "blue"; }

        // Textverarbeitung
        if (rbMSWord.Checked) { _settings.WordProcessorProgram = true; }
        else if (rbLibreOffice.Checked) { _settings.WordProcessorProgram = false; }
        else { _settings.WordProcessorProgram = null; }
    }

    // --- UI LOGIK & EVENTS ---

    // Wird automatisch aufgerufen, wenn das Formular geschlossen wird.
    // Wir speichern nur, wenn der User "OK" gedrückt hat.
    private void FrmProgSettings_FormClosing(object sender, FormClosingEventArgs e)
    {
        if (DialogResult == DialogResult.OK)
        {
            // Validierung (optional): Prüfen ob Standard-Pfad gesetzt ist, wenn rbStandard aktiv
            if (rbStandard.Checked && string.IsNullOrWhiteSpace(tbStandard.Text))
            {
                // Fallback, damit keine ungültige Config entsteht
                rbEmpty.Checked = true;
                // Alternativ: MessageBox anzeigen und e.Cancel = true setzen
            }

            MapUiToSettings();
        }
    }

    // UI-Status aktualisieren (Enabled/Disabled Logik)
    // HIER FINDET KEINE DATENÄNDERUNG STATT, NUR OPTIK
    private void UpdateUiState()
    {
        tbStandard.Enabled = btnStandardFile.Enabled = rbStandard.Checked;

        var backupActive = ckbBackup.Checked;
        tbBackupFolder.Enabled = btnBackupFolder.Enabled = backupActive;
        btnExplorer.Enabled = backupActive && !string.IsNullOrEmpty(tbBackupFolder.Text);

        var watchActive = ckbWatchFolder.Checked;
        tbWatchFolder.Enabled = btnWatchFolder.Enabled = lblWatchFolder.Enabled = watchActive;
    }

    private void FrmProgSettings_Load(object sender, EventArgs e)
    {
        // Ggf. Fokus setzen oder letzte Anpassungen
    }

    // Events, die die UI beeinflussen
    private void RbStandard_CheckedChanged(object sender, EventArgs e) => UpdateUiState();
    private void CkbBackup_CheckedChanged(object sender, EventArgs e) => UpdateUiState();
    private void TbBackupFolder_TextChanged(object sender, EventArgs e) => UpdateUiState();
    private void CkbWatchFolder_CheckedChanged(object sender, EventArgs e) => UpdateUiState();


    // --- File Dialog Buttons ---

    private void BtnStandardFile_Click(object sender, EventArgs e)
    {
        openFileDialog.InitialDirectory = !string.IsNullOrEmpty(tbStandard.Text) ? Path.GetDirectoryName(tbStandard.Text) : null;
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            tbStandard.Text = openFileDialog.FileName;
        }
    }

    private void BtnDatabaseFolder_Click(object sender, EventArgs e)
    {
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
        {
            tbDatabaseFolder.Text = folderBrowserDialog.SelectedPath;
        }
    }

    private void BtnBackupFolder_Click(object sender, EventArgs e)
    {
        folderBrowserDialog.InitialDirectory = Directory.Exists(tbBackupFolder.Text) ? tbBackupFolder.Text : string.Empty;
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
        {
            tbBackupFolder.Text = folderBrowserDialog.SelectedPath;
        }
    }

    private void BtnWatchFolder_Click(object sender, EventArgs e)
    {
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
        {
            tbWatchFolder.Text = folderBrowserDialog.SelectedPath;
        }
    }

    private void BtnExplorer_Click(object sender, EventArgs e)
    {
        if (Directory.Exists(tbBackupFolder.Text))
        {
            using var process = new Process();
            process.StartInfo.FileName = tbBackupFolder.Text;
            process.StartInfo.UseShellExecute = true;
            process.Start();
        }
        else { Console.Beep(); }
    }

    // --- Standard Form Kram (Tabs, Keys) ---

    private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.HighQuality;
        var tabPage = tabControl.TabPages[e.Index];
        var tabBounds = tabControl.GetTabRect(e.Index);
        var backBrush = e.State == DrawItemState.Selected ? SystemBrushes.GradientActiveCaption : SystemBrushes.GradientInactiveCaption;
        var textBrush = e.State == DrawItemState.Selected ? SystemBrushes.HighlightText : SystemBrushes.ControlText;
        g.FillRectangle(backBrush, e.Bounds);
        using var tabFont = new Font("Segoe UI", 10f);
        using var stringFlags = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };
        g.DrawString(tabPage.Text, tabFont, textBrush, tabBounds, stringFlags);
        if (e.Index == tabControl.TabCount - 1)
        {
            var totalTabHeight = tabBounds.Height * tabControl.TabCount;
            var remainingRect = new Rectangle(0, totalTabHeight, tabBounds.Width + 2, tabControl.Height - totalTabHeight);
            g.FillRectangle(SystemBrushes.Control, remainingRect);
        }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape) { Close(); return true; }
        // Kleiner Fix: Tabulatortaste sollte im TabControl navigieren, wenn Fokus nicht auf Controls liegt
        // Wenn du willst, dass TAB immer durch die Tabs wechselt:
        if (keyData == Keys.Tab && !msg.HWnd.Equals(tbStandard.Handle) && !msg.HWnd.Equals(tbBackupFolder.Handle))
        {
            // Hier Logik optional anpassen, oft reicht Standard-Windows Verhalten
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }
}