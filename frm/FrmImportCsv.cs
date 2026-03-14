using Adressen.cls;
using Microsoft.EntityFrameworkCore;
using System.Globalization;

namespace Adressen.frm;

public partial class FrmImportCsv : Form
{
    private List<string> _csvHeaders = new();
    private List<string[]> _csvData = new();
    private readonly List<string> _targetFields;
    private readonly string _currentDbPath;

    public string TargetDatabasePath { get; private set; } = string.Empty;

    public FrmImportCsv(List<string> availableFields, string currentDbPath)
    {
        InitializeComponent();
        _currentDbPath = currentDbPath;
        _targetFields = new List<string> { "- Ignorieren -" };
        _targetFields.AddRange(availableFields.OrderBy(f => f));
        foreach (DataGridViewColumn col in dgvMapping.Columns) { col.SortMode = DataGridViewColumnSortMode.NotSortable; }
        comboCol.DataSource = _targetFields;
        gbDuplicate.Enabled = rbCurrentDb.Enabled = !string.IsNullOrEmpty(_currentDbPath);
    }

    private async void BtnBrowse_Click(object sender, EventArgs e)
    {
        using var ofd = new OpenFileDialog();
        ofd.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*";

        if (ofd.ShowDialog() == DialogResult.OK)
        {
            // UI für den Ladevorgang sperren
            txtCsvPath.Text = "Lese CSV-Datei...";
            btnBrowse.Enabled = false;
            btnStartImport.Enabled = false;
            dgvMapping.Rows.Clear();

            try
            {
                // Asynchrones Einlesen im Hintergrund-Thread
                _csvData = await Task.Run(() => Utils.ReadCsv(ofd.FileName).ToList());
                txtCsvPath.Text = ofd.FileName;

                if (_csvData.Count < 2)
                {
                    Utils.MsgTaskDlg(Handle, "Ungültige CSV-Datei", "Die ausgewählte CSV-Datei enthält nicht genügend Daten (mindestens Header + 1 Datenzeile erforderlich).");
                    return;
                }

                // UI-Update im Haupt-Thread aufrufen
                LoadCsvPreviewUI();
                btnStartImport.Enabled = true;
            }
            catch (Exception ex)
            {
                txtCsvPath.Text = string.Empty;
                Utils.ErrTaskDlg(Handle, ex);
            }
            finally { btnBrowse.Enabled = true; }
        }
    }

    private void LoadCsvPreviewUI()
    {
        _csvHeaders = [.. _csvData[0]];
        var firstRow = _csvData[1];

        for (var i = 0; i < _csvHeaders.Count; i++)
        {
            var header = _csvHeaders[i];
            var sampleValue = i < firstRow.Length ? firstRow[i] : "";
            var rowIndex = dgvMapping.Rows.Add(header, sampleValue);
            var row = dgvMapping.Rows[rowIndex];
            var bestMatch = FindBestMatch(header);
            if (bestMatch != null) { row.Cells[2].Value = bestMatch; }
            else { row.Cells[2].Value = "- Ignorieren -"; }
            toolStripStatusLabel.Text = $"{_csvData.Count - 1} Zeilen, {_csvHeaders.Count} Spalten";  // minus 1, weil die erste Zeile die Header sind
        }
    }

    private async void BtnStartImport_Click(object sender, EventArgs e)
    {
        if (_csvData.Count < 2)
        {
            Utils.MsgTaskDlg(Handle, "Keine CSV-Daten", "Bitte wähle zuerst eine gültige CSV-Datei aus.");
            return;
        }
        var fieldMapping = new Dictionary<int, string>();
        for (var i = 0; i < dgvMapping.Rows.Count; i++)
        {
            var targetField = dgvMapping.Rows[i].Cells[2].Value?.ToString();
            if (!string.IsNullOrEmpty(targetField) && targetField != "- Ignorieren -") { fieldMapping.Add(i, targetField); }
        }

        if (fieldMapping.Count == 0)
        {
            Utils.MsgTaskDlg(Handle, "Keine Spalten zugeordnet", "Es wurde keine einzige Spalte zugeordnet. Bitte ordne mindestens eine Spalte zu, damit der Import starten kann.");
            return;
        }

        TargetDatabasePath = _currentDbPath;
        if (rbNewDb.Checked)
        {
            using var sfd = new SaveFileDialog();
            sfd.Filter = "Adressen-Datenbank (*.adb)|*.adb";
            sfd.Title = "Neue Datenbank speichern unter...";
            sfd.DefaultExt = "adb";
            if (sfd.ShowDialog() != DialogResult.OK) { return; }
            TargetDatabasePath = sfd.FileName;
        }

        btnStartImport.Enabled = false;
        btnCancel.Enabled = false;
        lnkExample.Enabled = false;
        progressBar.Visible = true;
        progressBar.Maximum = _csvData.Count - 1;
        progressBar.Value = 0;
        toolStripStatusLabel.Text = "Import läuft...";
        var progress = new Progress<int>(v => { progressBar.Value = v; });
        var skipDuplicates = rbDuplicateSkip.Checked;
        try
        {
            var (importedCount, skippedCount) = await Task.Run(() => PerformImportAsync(TargetDatabasePath, fieldMapping, rbNewDb.Checked, skipDuplicates, progress));  // Tupel
            toolStripStatusLabel.Text = $"Erfolgreich: {importedCount} importiert, {skippedCount} ignoriert.";
            Utils.MsgTaskDlg(Handle, "Import abgeschlossen", $"Der Import wurde erfolgreich abgeschlossen!\n\nImportiert: {importedCount}\nAls Duplikat übersprungen: {skippedCount}", TaskDialogIcon.ShieldSuccessGreenBar);
            DialogResult = DialogResult.OK;
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(Handle, ex);
            btnStartImport.Enabled = true;
            btnCancel.Enabled = true;
            lnkExample.Enabled = true;
        }
        finally { progressBar.Visible = false; }
    }

    private async Task<(int imported, int skipped)> PerformImportAsync(string dbPath, Dictionary<int, string> fieldMapping, bool isNewDatabase, bool skipDuplicates, IProgress<int> progress)
    {
        if (isNewDatabase && File.Exists(dbPath)) { File.Delete(dbPath); }  // context.Database.EnsureCreatedAsync() macht nichts, wenn die Datei schon existiert (belässt veraltete Tabellen)
        using var context = new AdressenDbContext(dbPath);
        if (isNewDatabase)
        {
            await context.Database.EnsureCreatedAsync();
            await context.Database.ExecuteSqlRawAsync($"PRAGMA user_version = {AppSettings.DatabaseSchemaVersion};");
        }
        context.ChangeTracker.AutoDetectChangesEnabled = false;

        using var transaction = await context.Database.BeginTransactionAsync();  // Transaktion für bessere Performance und um bei Fehlern die Datenbank nicht zu beschädigen
        var processedCount = 0;
        var skippedCount = 0;
        try
        {
            // --- 1. Vorhandene Daten für die Duplikatprüfung laden ---
            var existingMails = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (!isNewDatabase && skipDuplicates)
            {
                // Wir laden nur die relevanten Felder, um Arbeitsspeicher zu sparen
                var dbAdressen = await context.Adressen.Select(a => new { a.Mail1, a.Vorname, a.Nachname }).ToListAsync();
                foreach (var a in dbAdressen)
                {
                    if (!string.IsNullOrWhiteSpace(a.Mail1)) { existingMails.Add(a.Mail1.Trim()); }
                    var fullName = $"{a.Vorname?.Trim()}|{a.Nachname?.Trim()}";
                    if (fullName != "|") { existingNames.Add(fullName); }
                }
            }

            // --- 2. CSV-Daten verarbeiten ---
            var batch = new List<Adresse>();
            var culture = new CultureInfo("de-DE");

            foreach (var csvRow in _csvData.Skip(1))
            {
                var adresse = new Adresse();
                foreach (var kvp in fieldMapping)
                {
                    var colIndex = kvp.Key;
                    var targetProperty = kvp.Value;
                    if (colIndex >= csvRow.Length) { continue; }
                    var val = csvRow[colIndex]?.Trim();
                    if (string.IsNullOrEmpty(val)) { continue; }
                    if (targetProperty == "Gruppen")
                    {
                        var gruppenNamen = val.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                        foreach (var gName in gruppenNamen)
                        {
                            var gruppe = context.Gruppen.Local.FirstOrDefault(g => g.Name.Equals(gName, StringComparison.OrdinalIgnoreCase))
                                         ?? context.Gruppen.FirstOrDefault(g => g.Name.Equals(gName, StringComparison.CurrentCultureIgnoreCase));

                            if (gruppe == null)
                            {
                                gruppe = new Gruppe { Name = gName };
                                context.Gruppen.Add(gruppe);
                            }
                            adresse.Gruppen.Add(gruppe);
                        }
                    }
                    else if (targetProperty == "Geburtstag")
                    {
                        if (DateOnly.TryParse(val, culture, out var dt)) { adresse.Geburtstag = dt; }
                    }
                    else { adresse.SetPropertyValue(targetProperty, val); }
                }

                // --- 3. Duplikat prüfen ---
                var isDuplicate = false;

                if (skipDuplicates)
                {
                    if (!string.IsNullOrWhiteSpace(adresse.Mail1) && existingMails.Contains(adresse.Mail1.Trim())) { isDuplicate = true; }  // Prüfen nach E-Mail
                    var fullName = $"{adresse.Vorname?.Trim()}|{adresse.Nachname?.Trim()}";  // Prüfen nach Vor- und Nachname
                    if (fullName != "|" && existingNames.Contains(fullName)) { isDuplicate = true; }
                    if (!isDuplicate)  // Wenn es kein Duplikat ist, direkt den HashSets hinzufügen, damit keine Duplikate innerhalb der CSV-Datei selbst entstehen.
                    {
                        if (!string.IsNullOrWhiteSpace(adresse.Mail1)) { existingMails.Add(adresse.Mail1.Trim()); }
                        if (fullName != "|") { existingNames.Add(fullName); }
                    }
                }
                if (isDuplicate)
                {
                    skippedCount++;
                    continue; // Nächste CSV-Zeile, diesen Eintrag überspringen
                }
                batch.Add(adresse);
                processedCount++;
                if (batch.Count >= 1000)  // Performance-Tweak: Batching. Verhindert, dass der RAM bei riesigen Dateien überläuft.
                {
                    context.Adressen.AddRange(batch);
                    await context.SaveChangesAsync();
                    batch.Clear();
                }
                if ((processedCount + skippedCount) % 50 == 0) { progress.Report(processedCount + skippedCount); }  // Fortschritt im UI melden
            }
            if (batch.Count > 0)  // Restliche Datensätze speichern, die am Ende noch im Batch liegen
            {
                context.Adressen.AddRange(batch);
                await context.SaveChangesAsync();
            }
            await transaction.CommitAsync();
        }
        catch
        {
            await transaction.RollbackAsync();  // Wenn irgendwo ein Fehler auftritt, werden alle bisherigen SaveChanges() dieses Imports verworfen
            throw; // Fehler an die UI (BtnStartImport_Click) weiterleiten, damit die rote Fehlermeldung kommt
        }
        progress.Report(_csvData.Count - 1);
        return (processedCount, skippedCount);
    }

    private string? FindBestMatch(string csvHeader)
    {
        var exactMatch = _targetFields.FirstOrDefault(f => f.Equals(csvHeader, StringComparison.OrdinalIgnoreCase));
        if (exactMatch != null) { return exactMatch; }
        var cleanHeader = csvHeader.Trim().ToLowerInvariant();
        if (cleanHeader == "vorname" || cleanHeader == "vornamen") { return "Vorname"; }
        if (cleanHeader == "nachname" || cleanHeader == "familienname" || cleanHeader == "zuname" || cleanHeader == "name") { return "Nachname"; }
        if (cleanHeader == "firma" || cleanHeader == "unternehmen" || cleanHeader == "betrieb" || cleanHeader == "organisation") { return "Unternehmen"; }
        if (cleanHeader == "stadt" || cleanHeader == "wohnort" || cleanHeader == "ort") { return "Ort"; }
        if (cleanHeader == "plz" || cleanHeader == "postleitzahl") { return "PLZ"; }
        if (cleanHeader == "telefon" || cleanHeader == "festnetz" || cleanHeader == "tel" || cleanHeader == "tel.") { return "Telefon1"; }
        if (cleanHeader == "handy" || cleanHeader == "mobil" || cleanHeader == "smartphone" || cleanHeader == "mobiltelefon") { return "Mobil"; }
        if (cleanHeader == "e-mail" || cleanHeader == "email" || cleanHeader == "mail" || cleanHeader == "e-mail-adresse") { return "Mail1"; }
        if (cleanHeader == "geburtstag" || cleanHeader == "geburtsdatum" || cleanHeader == "geboren am") { return "Geburtstag"; }
        if (cleanHeader == "bemerkung" || cleanHeader == "bemerkungen" || cleanHeader == "notiz" || cleanHeader == "notizen" || cleanHeader == "kommentar") { return "Notizen"; }
        if (cleanHeader == "gruppe" || cleanHeader == "gruppen" || cleanHeader == "kategorien" || cleanHeader == "tags") { return "Gruppen"; }
        if (cleanHeader == "webseite" || cleanHeader == "website" || cleanHeader == "homepage" || cleanHeader == "url") { return "Internet"; }
        return null;
    }

    private void LnkExample_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    {
        var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var filePath = Path.Combine(desktopPath, "Adressen_Import_Vorlage.csv");
        try
        {
            var exportColumns = _targetFields.Where(static f => f != "- Ignorieren -").ToList();
            using var writer = new StreamWriter(filePath, false, System.Text.Encoding.UTF8);
            writer.WriteLine(string.Join(";", exportColumns));
            var exampleData = exportColumns.Select(col =>
            {
                return col switch
                {
                    "Vorname" => "Max",
                    "Nachname" => "Mustermann",
                    "Unternehmen" => "Muster GmbH",
                    "Mail1" => "max@muster.de",
                    "Geburtstag" => "12.05.1985",
                    "Gruppen" => "Kunden, VIP",
                    _ => ""
                };
            });
            writer.WriteLine(string.Join(";", exampleData));
            Utils.MsgTaskDlg(Handle, "Vorlage erstellt", $"Die Datei '{Path.GetFileName(filePath)}' wurde erfolgreich auf deinem Desktop gespeichert.", TaskDialogIcon.Information);
        }
        catch (Exception ex) { Utils.ErrTaskDlg(Handle, ex); }
    }
}