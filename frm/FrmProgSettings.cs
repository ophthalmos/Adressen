using System.Diagnostics;
using System.Drawing.Drawing2D;
using Adressen.cls;

namespace Adressen;

public partial class FrmProgSettings : Form
{
    private readonly AppSettings _settings;
    private readonly Font _tabFont = new("Segoe UI", 10f);
    private readonly StringFormat _tabStringFormat = new() { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

    internal FrmProgSettings(AppSettings settings)
    {
        InitializeComponent();
        _settings = settings;
        cbxFontName.BeginUpdate();
        cbxFontName.Items.Clear();
        cbxFontName.Items.AddRange([.. FontManager.GetValidFonts()]);  // Schriften direkt aus dem Manager abrufen (geht in ms)
        Utils.AdjustComboBoxDropDownWidth(cbxFontName);
        cbxFontName.EndUpdate();
        MapSettingsToUi();  // Dann die gespeicherte Schriftart auswählen
        UpdateUiState();
        var (autoStart, min2Tray) = Utils.IsAutoStartEnabled(Application.ProductName, Application.ExecutablePath);
        ckbAutostart.Checked = autoStart;
        ckbMin2Tray.Checked = min2Tray;
    }

    internal void SelectAnrufMonTab() => tabControl.SelectedTab = tpAnrufMon;

    private void MapSettingsToUi()  // Mapping statt DataBinding, da wir hier auch Logik für die RadioButtons haben und die Schriftart-Auswahl…
    {
        // CheckBoxen
        ckbAskBeforeDelete.Checked = _settings.AskBeforeDelete;
        ckbContactsAutoload.Checked = _settings.ContactsAutoload;
        ckbBackup.Checked = _settings.DailyBackup;
        ckbZipArchive.Checked = _settings.AddZipBackup;
        ckbWatchFolder.Checked = _settings.WatchFolder;
        ckbAskBeforeSaveSQL.Checked = _settings.AskBeforeSaveSQL;
        ckbAskBeforeSaveSQLExpander.Checked = _settings.AskBeforeSaveSQLExpander;
        ckbAskPrintEnvelope.Checked = _settings.AskPrintEnvelope;
        ckbPlaceholderText.Checked = _settings.ShowPlaceholderText;
        ckbFritzMonitorEnabled.Checked = _settings.FritzMonitorEnabled;
        ckbMonitorContactsFirst.Checked = _settings.FritzContactsFirst;
        ckbFritzPlaySound.Checked = _settings.FritzMonitorPlaySound;

        // TextBoxen
        tbStandard.Text = _settings.StandardFile;
        tbBackupFolder.Text = _settings.BackupDirectory;
        tbZipArchive.Text = _settings.AddZipDirectory;
        tbDatabaseFolder.Text = _settings.DatabaseFolder;
        tbWatchFolder.Text = _settings.DocumentFolder;
        iPv4AddressControl.Address = _settings.FritzBoxHost;

        // RadioButtons: Start-Verhalten
        if (_settings.ReloadRecent) { rbRecent.Checked = true; }
        else if (_settings.NoAutoload) { rbEmpty.Checked = true; }
        else { rbStandard.Checked = true; }

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

        // Schriftart setzen
        if (cbxFontName.Items.Contains(_settings.AppFontName)) { cbxFontName.SelectedItem = _settings.AppFontName; }
        else if (cbxFontName.Items.Count > 0) { cbxFontName.SelectedIndex = 0; }  // Fallback
        // Schriftgröße setzen (absichern gegen Werte außerhalb von 9-12)
        var size = (decimal)_settings.AppFontSize;
        if (size < nudFontSize.Minimum) { size = nudFontSize.Minimum; }
        if (size > nudFontSize.Maximum) { size = nudFontSize.Maximum; }
        nudFontSize.Value = size;
        UpdateFontResetButtonState();
    }

    private void MapUiToSettings()
    {
        // CheckBoxen
        _settings.AskBeforeDelete = ckbAskBeforeDelete.Checked;
        _settings.ContactsAutoload = ckbContactsAutoload.Checked;
        _settings.DailyBackup = ckbBackup.Checked;
        _settings.AddZipBackup = ckbZipArchive.Checked;
        _settings.WatchFolder = ckbWatchFolder.Checked;
        _settings.AskBeforeSaveSQL = ckbAskBeforeSaveSQL.Checked;
        _settings.AskBeforeSaveSQLExpander = ckbAskBeforeSaveSQLExpander.Checked;
        _settings.AskPrintEnvelope = ckbAskPrintEnvelope.Checked;
        _settings.ShowPlaceholderText = ckbPlaceholderText.Checked;
        _settings.FritzMonitorEnabled = ckbFritzMonitorEnabled.Checked;
        _settings.FritzContactsFirst = ckbMonitorContactsFirst.Checked;
        // TextBoxen
        _settings.StandardFile = Utils.CorrectUNC(tbStandard.Text.Trim());
        _settings.BackupDirectory = Utils.CorrectUNC(tbBackupFolder.Text.Trim());
        _settings.AddZipDirectory = Utils.CorrectUNC(tbZipArchive.Text.Trim());
        _settings.DatabaseFolder = Utils.CorrectUNC(tbDatabaseFolder.Text.Trim());
        _settings.DocumentFolder = Utils.CorrectUNC(tbWatchFolder.Text.Trim());
        _settings.FritzBoxHost = iPv4AddressControl.Address;
        _settings.FritzMonitorPlaySound = ckbFritzPlaySound.Checked;

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

        _settings.AppFontName = cbxFontName.SelectedItem?.ToString() ?? "Segoe UI";
        _settings.AppFontSize = (float)nudFontSize.Value;

        // Textverarbeitung
        if (rbMSWord.Checked) { _settings.WordProcessorProgram = true; }
        else if (rbLibreOffice.Checked) { _settings.WordProcessorProgram = false; }
        else { _settings.WordProcessorProgram = null; }
    }

    // --- UI LOGIK & EVENTS ---

    private void FrmProgSettings_FormClosing(object sender, FormClosingEventArgs e)
    {
        if (DialogResult == DialogResult.OK)
        {
            if (rbStandard.Checked && string.IsNullOrWhiteSpace(tbStandard.Text))
            {
                Utils.MsgTaskDlg(Handle, "Eingabe unvollständig", "Bitte wähle eine Standard-Datei aus oder ändere das Start-Verhalten.", TaskDialogIcon.ShieldWarningYellowBar);
                tbStandard.Focus();
                e.Cancel = true; // Verhindert das Schließen des Fensters
                return;
            }
            MapUiToSettings();
        }
    }

    private void TbStandard_Validating(object sender, System.ComponentModel.CancelEventArgs e)
    {
        var path = tbStandard.Text.Trim();

        // 1. Leere Eingaben zulassen (wird erst beim Klick auf OK endgültig geprüft, falls rbStandard aktiv ist)
        if (string.IsNullOrEmpty(path)) { return; }

        try
        {
            // 2. Ungültige Zeichen abfangen
            var invalidChars = Path.GetInvalidPathChars();
            if (path.IndexOfAny(invalidChars) >= 0)
            {
                Utils.MsgTaskDlg(Handle, "Ungültiger Pfad", "Der Dateipfad enthält ungültige Zeichen.", TaskDialogIcon.ShieldErrorRedBar);
                e.Cancel = true;
                return;
            }

            // 3. Prüfen, ob das Verzeichnis existiert
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                if (!Directory.Exists(directory))
                {
                    Utils.MsgTaskDlg(Handle, "Verzeichnis nicht gefunden", $"Das Verzeichnis '{directory}' existiert nicht.", TaskDialogIcon.ShieldWarningYellowBar);
                    e.Cancel = true;
                    return;
                }
            }

            // 4. Dateiendung prüfen (optional korrigieren, falls vergessen)
            var extension = Path.GetExtension(path);
            if (string.IsNullOrEmpty(extension))
            {
                // Angenommen, deine Standard-Datenbanken enden auf .db
                path = Path.ChangeExtension(path, ".db");
                tbStandard.Text = path;
            }

            // 5. Prüfen, ob die Datei tatsächlich existiert
            if (!File.Exists(path))
            {
                Utils.MsgTaskDlg(Handle, "Datei nicht gefunden", $"Die Datenbank-Datei '{Path.GetFileName(path)}' konnte nicht gefunden werden.\nBitte wähle eine existierende Datei aus.", TaskDialogIcon.ShieldWarningYellowBar);
                e.Cancel = true;
            }
        }
        catch (Exception ex)
        {
            // Fängt Fälle ab, in denen der Pfad völlig unlesbar formatiert ist
            Utils.ErrTaskDlg(Handle, ex);
            e.Cancel = true;
        }
    }

    private void TbZipArchive_Validating(object sender, System.ComponentModel.CancelEventArgs e)
    {
        var path = tbZipArchive.Text.Trim();

        // 1. Wenn das Feld leer ist, ist das in Ordnung (Feature wird damit deaktiviert)
        if (string.IsNullOrEmpty(path)) { return; }

        try
        {
            // 2. Ungültige Zeichen abfangen
            var invalidChars = Path.GetInvalidPathChars();
            if (path.IndexOfAny(invalidChars) >= 0)
            {
                Utils.MsgTaskDlg(Handle, "Ungültiger Pfad", "Der Dateipfad enthält ungültige Zeichen.", TaskDialogIcon.ShieldErrorRedBar);
                e.Cancel = true;
                return;
            }

            // 3. Dateiendung prüfen und automatisch korrigieren
            var extension = Path.GetExtension(path);
            if (!string.Equals(extension, ".zip", StringComparison.OrdinalIgnoreCase))
            {
                // Ersetzt eine falsche Endung durch .zip oder hängt sie an, falls keine vorhanden ist
                tbZipArchive.Text = Path.ChangeExtension(path, ".zip");

                // Den aktualisierten Pfad für die nächste Prüfung übernehmen
                path = tbZipArchive.Text;
            }

            // 4. Verzeichnis prüfen (existiert der Zielordner?)
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                if (!Directory.Exists(directory))
                {
                    Utils.MsgTaskDlg(Handle, "Verzeichnis nicht gefunden", $"Das Verzeichnis '{directory}' existiert nicht.\nBitte wähle einen bestehenden Ordner.", TaskDialogIcon.ShieldWarningYellowBar);
                    e.Cancel = true;
                }
            }
        }
        catch (Exception ex)
        {
            // Fängt Fälle ab, in denen der Pfad völlig unlesbar formatiert ist
            Utils.ErrTaskDlg(Handle, ex);
            e.Cancel = true;
        }
    }

    // UI-Status aktualisieren (Enabled/Disabled Logik)
    // HIER FINDET KEINE DATENÄNDERUNG STATT, NUR OPTIK
    private void UpdateUiState()
    {
        tbStandard.Enabled = btnStandardFile.Enabled = rbStandard.Checked;
        ckbAskBeforeSaveSQLExpander.Enabled = ckbAskBeforeSaveSQL.Checked;
        ckbAskBeforeSaveSQLExpander.Enabled = ckbAskBeforeSaveSQL.Checked;
        if (ckbAskBeforeSaveSQLExpander.Checked)
        {
            ckbAskBeforeSaveSQL.Checked = true;
            ckbAskBeforeSaveSQL.Enabled = false;
        }
        else { ckbAskBeforeSaveSQL.Enabled = true; }
        var backupActive = ckbBackup.Checked;
        tbBackupFolder.Enabled = btnBackupFolder.Enabled = backupActive;
        btnExplorer.Enabled = backupActive && !string.IsNullOrEmpty(tbBackupFolder.Text);
        tbZipArchive.Enabled = btnZipArchive.Enabled = ckbZipArchive.Checked;
        var watchActive = ckbWatchFolder.Checked;
        tbWatchFolder.Enabled = btnWatchFolder.Enabled = lblWatchFolder.Enabled = watchActive;
        var isMonitorActive = ckbFritzMonitorEnabled.Checked;
        iPv4AddressControl.Enabled = isMonitorActive;
        lblFritzBoxHost.Enabled = isMonitorActive;
        ckbMonitorContactsFirst.Enabled = isMonitorActive;
        ckbFritzPlaySound.Enabled = isMonitorActive;
    }

    // Events, die die UI beeinflussen
    private void RbStandard_CheckedChanged(object sender, EventArgs e) => UpdateUiState();
    private void CkbBackup_CheckedChanged(object sender, EventArgs e) => UpdateUiState();
    private void TbBackupFolder_TextChanged(object sender, EventArgs e) => UpdateUiState();
    private void CkbWatchFolder_CheckedChanged(object sender, EventArgs e) => UpdateUiState();
    private void CkbZipArchive_CheckedChanged(object sender, EventArgs e) => UpdateUiState();
    private void TbZipArchive_TextChanged(object sender, EventArgs e) => UpdateUiState();
    private void CkbAskBeforeSaveSQL_CheckedChanged(object sender, EventArgs e) => UpdateUiState();
    private void CkbAskBeforeSaveSQLExpander_CheckedChanged(object sender, EventArgs e) => UpdateUiState();

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
        folderBrowserDialog.Description = "Wähle den Datenbankordner:";
        if (Directory.Exists(tbDatabaseFolder.Text))
        {
            folderBrowserDialog.InitialDirectory = tbDatabaseFolder.Text;
        }
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
        {
            tbDatabaseFolder.Text = folderBrowserDialog.SelectedPath;
        }
    }

    private void BtnBackupFolder_Click(object sender, EventArgs e)
    {
        folderBrowserDialog.Description = "Wähle den Sicherungsordner:";
        folderBrowserDialog.InitialDirectory = Directory.Exists(tbBackupFolder.Text) ? tbBackupFolder.Text : string.Empty;
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
        {
            tbBackupFolder.Text = folderBrowserDialog.SelectedPath;
        }
    }

    private void BtnWatchFolder_Click(object sender, EventArgs e)
    {
        folderBrowserDialog.Description = "Wähle den zu überwachenden Ordner:";
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

    private void BtnZipArchive_Click(object sender, EventArgs e)
    {
        var currentPath = tbZipArchive.Text.Trim();
        var initialDir = string.Empty;
        if (!string.IsNullOrEmpty(currentPath))
        {
            var dir = Path.GetDirectoryName(currentPath);
            if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir)) { initialDir = dir; }
        }
        folderBrowserDialog.Description = "Wähle den Zielordner oder klicke auf ein bestehendes ZIP-Archiv:";
        if (!string.IsNullOrEmpty(initialDir)) { folderBrowserDialog.InitialDirectory = initialDir; }
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
        {
            var selectedPath = folderBrowserDialog.SelectedPath;
            if (selectedPath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)) { tbZipArchive.Text = selectedPath; }  // Windows behandelt ZIP-Datei als Ordner
            else { tbZipArchive.Text = Path.Combine(selectedPath, "Adressen.zip"); }  // regulärer Ordner: -> Dateinamen anhängen
        }
    }

    // --- Standard Form Kram (Tabs, Keys) ---
    private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.HighQuality;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

        var tabPage = tabControl.TabPages[e.Index];
        var tabBounds = tabControl.GetTabRect(e.Index);
        var backBrush = e.State == DrawItemState.Selected ? SystemBrushes.GradientActiveCaption : SystemBrushes.GradientInactiveCaption;
        var textBrush = e.State == DrawItemState.Selected ? SystemBrushes.HighlightText : SystemBrushes.ControlText;

        g.FillRectangle(backBrush, e.Bounds);

        // Hier nutzen wir nun die performanten Klassenvariablen
        g.DrawString(tabPage.Text, _tabFont, textBrush, tabBounds, _tabStringFormat);

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
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void CbxFontName_DrawItem(object sender, DrawItemEventArgs e)
    {
        if (e.Index < 0) { return; }

        e.DrawBackground();

        var fontName = cbxFontName.Items[e.Index]?.ToString();
        // Schrift zentral abrufen (wird automatisch gecacht)
        var displayFont = FontManager.GetDisplayFont(fontName);

        var isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        var textColor = isSelected ? SystemColors.HighlightText : e.ForeColor;
        var backColor = isSelected ? SystemColors.Highlight : e.BackColor;

        var textRect = e.Bounds;
        var flags = TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding;

        TextRenderer.DrawText(e.Graphics, fontName, displayFont, textRect, textColor, backColor, flags);

        e.DrawFocusRectangle();
    }

    private void UpdateFontResetButtonState()
    {
        var isStandardFont = cbxFontName.SelectedItem?.ToString() == "Segoe UI";
        var isStandardSize = nudFontSize.Value == 10m;
        btnFontReset.Enabled = !(isStandardFont && isStandardSize);  // Der Button ist nur aktiviert, wenn NICHT beides bereits auf Standard steht
    }

    private void BtnFontReset_Click(object sender, EventArgs e)
    {
        var defaultFont = "Segoe UI";
        if (cbxFontName.Items.Contains(defaultFont)) { cbxFontName.SelectedItem = defaultFont; }
        else if (cbxFontName.Items.Count > 0) { cbxFontName.SelectedIndex = 0; }  // Extremer Fallback, falls "Segoe UI" fehlt
        nudFontSize.Value = 10m;  // NumericUpDown akzeptiert nur decimal, daher das 'm' für die Literal-Notation
    }

    private void CbxFontName_SelectedIndexChanged(object sender, EventArgs e) => UpdateFontResetButtonState();
    private void NudFontSize_ValueChanged(object sender, EventArgs e) => UpdateFontResetButtonState();

    private void CkbFritzMonitorEnabled_CheckedChanged(object sender, EventArgs e) => UpdateUiState();

    private void CkbAutostart_CheckedChanged(object sender, EventArgs e)
    {
        if (ckbAutostart.Focused)
        {
            if (ckbAutostart.Checked)
            {
                Utils.SetAutoStart(Application.ProductName, Application.ExecutablePath, ckbMin2Tray.Checked ? "-min2Tray" : "");
                ckbMin2Tray.Enabled = true;
            }
            else
            {
                Utils.UnSetAutoStart(Application.ProductName);
                ckbMin2Tray.Enabled = false;
            }
        }
    }

    private void CkbMin2Tray_CheckedChanged(object sender, EventArgs e)
    {
        if (ckbMin2Tray.Focused)
        {
            if (ckbMin2Tray.Checked && ckbAutostart.Checked) { Utils.SetAutoStart(Application.ProductName, Application.ExecutablePath, "-min2Tray"); }
            else if (ckbAutostart.Checked) { Utils.SetAutoStart(Application.ProductName, Application.ExecutablePath, ""); }  // Autostart bleibt aktiv, aber ohne Parameter
        }
    }

    private void LinkLblAdressen_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) => tabControl.SelectedTab = tpAdressen;
}