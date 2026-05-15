using System.Diagnostics;
using System.Text;

namespace Adressen.cls;

internal class HtmlPrintService
{
    public static void ExportToHtmlAndPrint(DataGridView dgv, string title, bool isCompact = false)
    {
        var html = new StringBuilder();
        html.AppendLine("<!DOCTYPE html><html><head><meta charset='utf-8'>");
        html.AppendLine("<style>");
        html.AppendLine("body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; }");
        html.AppendLine("table { border-collapse: collapse; margin-top: 20px; width: auto; min-width: 50%; }");
        html.AppendLine("th { background-color: #f2f2f2; text-align: left; padding: 8px; border-bottom: 2px solid #ddd; }");
        html.AppendLine("td { padding: 8px; border-bottom: 1px solid #ddd; vertical-align: top; }");
        html.AppendLine(".nowrap { white-space: nowrap; }");
        html.AppendLine("tr:nth-child(even) { background-color: #f9f9f9; }");
        html.AppendLine(".footer { margin-top: 30px; padding-top: 10px; border-top: 1px solid #ddd; text-align: center; font-size: 12px; color: #555; }");
        html.AppendLine("@media print { .no-print { display: none; } button { display: none; } .footer { position: fixed; bottom: 0; left: 0; right: 0; border: none; background-color: white; } }");
        html.AppendLine("@media print { .no-print { display: none; } button { display: none; } }");
        html.AppendLine("</style></head><body>");
        html.AppendLine($"<h2>{title}</h2>");
        html.AppendLine($"<p>Erstellt am: {DateTime.Now:dd.MM.yyyy HH:mm}</p>");
        html.AppendLine("<button onclick='window.print()' style='padding: 10px 20px; background-color: #0071ca; color: white; border-radius: 12px; cursor: pointer;' class='no-print'>Drucken / Als PDF speichern</button>");
        var visibleColumns = dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).OrderBy(c => c.DisplayIndex).ToList();  // nur sichtbare Spalten, sortiert nach Anzeige

        if (isCompact) 
        {
            // --- KOMPAKTER MODUS ---
            string[] nameCols = ["Nachname", "Vorname", "Zwischenname", "Praefix", "Suffix", "Anrede", "Nickname"];
            string[] addressCols = ["Strasse", "PLZ", "Ort", "Land", "Postfach"];
            string[] phoneCols = ["Telefon1", "Telefon2", "Mobil", "Fax"];
            string[] mailCols = ["Mail1", "Mail2"];

            var hasName = visibleColumns.Any(c => nameCols.Contains(c.Name));
            var hasAddress = visibleColumns.Any(c => addressCols.Contains(c.Name));
            var hasPhone = visibleColumns.Any(c => phoneCols.Contains(c.Name));
            var hasMail = visibleColumns.Any(c => mailCols.Contains(c.Name));

            // Alle verbleibenden Spalten (z. B. Unternehmen, Notizen, etc.) behalten wir als eigene Spalten am Ende
            var otherCols = visibleColumns.Where(c =>
                !nameCols.Contains(c.Name) &&
                !addressCols.Contains(c.Name) &&
                !phoneCols.Contains(c.Name) &&
                !mailCols.Contains(c.Name)).ToList();

            // Header generieren
            html.AppendLine("<table><thead><tr>");
            if (hasName) { html.AppendLine("<th>Name</th>"); }
            if (hasAddress) { html.AppendLine("<th>Anschrift</th>"); }
            if (hasPhone) { html.AppendLine("<th>Telefon / Fax</th>"); }
            if (hasMail) { html.AppendLine("<th>E-Mail</th>"); }
            foreach (var col in otherCols) { html.AppendLine($"<th>{col.HeaderText}</th>"); }
            html.AppendLine("</tr></thead><tbody>");

            // Hilfsfunktion: Wert nur holen, wenn die Spalte existiert UND aktuell eingeblendet ist
            string GetVisVal(DataGridViewRow r, string colName)
            {
                if (dgv.Columns.Contains(colName))
                {
                    var col = dgv.Columns[colName];
                    if (col != null && col.Visible) { return r.Cells[col.Index].Value?.ToString()?.Trim() ?? string.Empty; }
                }
                return string.Empty;
            }

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) { continue; }
                html.AppendLine("<tr>");
                if (hasName)
                {
                    var nn = GetVisVal(row, "Nachname");
                    var vn = GetVisVal(row, "Vorname");
                    var zn = GetVisVal(row, "Zwischenname");
                    var pr = GetVisVal(row, "Praefix");
                    var sf = GetVisVal(row, "Suffix");
                    var vornamenKomplett = string.Join(" ", new[] { vn, zn }.Where(s => !string.IsNullOrEmpty(s)));
                    List<string> nameParts = [];
                    if (!string.IsNullOrEmpty(nn)) { nameParts.Add(nn); }
                    if (!string.IsNullOrEmpty(vornamenKomplett)) { nameParts.Add(vornamenKomplett); }
                    if (!string.IsNullOrEmpty(pr)) { nameParts.Add(pr); }
                    if (!string.IsNullOrEmpty(sf)) { nameParts.Add(sf); }
                    html.AppendLine($"<td>{string.Join(", ", nameParts)}</td>");
                }
                if (hasAddress)
                {
                    var str = GetVisVal(row, "Strasse");
                    var pf = GetVisVal(row, "Postfach");
                    var plz = GetVisVal(row, "PLZ");
                    var ort = GetVisVal(row, "Ort");
                    var land = GetVisVal(row, "Land");

                    List<string> addressParts = [];
                    if (!string.IsNullOrEmpty(str)) { addressParts.Add(str); }
                    else if (!string.IsNullOrEmpty(pf)) { addressParts.Add($"Postfach {pf}"); }

                    var plzOrt = $"{plz} {ort}".Trim();
                    if (!string.IsNullOrEmpty(plzOrt)) { addressParts.Add(plzOrt); }
                    if (!string.IsNullOrEmpty(land)) { addressParts.Add(land); }

                    html.AppendLine($"<td>{string.Join("<br>", addressParts)}</td>");
                }
                if (hasPhone)
                {
                    var t1 = GetVisVal(row, "Telefon1");
                    var t2 = GetVisVal(row, "Telefon2");
                    var mob = GetVisVal(row, "Mobil");
                    var fax = GetVisVal(row, "Fax");
                    List<string> phoneParts = [];
                    if (!string.IsNullOrEmpty(t1)) { phoneParts.Add($"{t1} (Tel 1)"); }
                    if (!string.IsNullOrEmpty(t2)) { phoneParts.Add($"{t2} (Tel 2)"); }
                    if (!string.IsNullOrEmpty(mob)) { phoneParts.Add($"{mob} (Mobil)"); }
                    if (!string.IsNullOrEmpty(fax)) { phoneParts.Add($"{fax} (Fax)"); }
                    html.AppendLine($"<td class=\"nowrap\">{string.Join("<br>", phoneParts)}</td>");
                }
                if (hasMail)
                {
                    var m1 = GetVisVal(row, "Mail1");
                    var m2 = GetVisVal(row, "Mail2");
                    List<string> mailParts = [];
                    if (!string.IsNullOrEmpty(m1)) { mailParts.Add(m1); }
                    if (!string.IsNullOrEmpty(m2)) { mailParts.Add(m2); }
                    html.AppendLine($"<td>{string.Join("<br>", mailParts)}</td>");
                }
                foreach (var col in otherCols)
                {
                    var val = row.Cells[col.Index].Value?.ToString() ?? string.Empty;
                    html.AppendLine($"<td>{val.Replace("\n", "<br>")}</td>");
                }

                html.AppendLine("</tr>");
            }
        }
        else  // STANDARD-MODUS (1:1 Abbildung)
        {
            html.AppendLine("<table><thead><tr>");
            foreach (var col in visibleColumns) { html.AppendLine($"<th>{col.HeaderText}</th>"); }
            html.AppendLine("</tr></thead><tbody>");
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) { continue; }
                html.AppendLine("<tr>");
                foreach (var col in visibleColumns)
                {
                    var val = row.Cells[col.Index].Value?.ToString() ?? string.Empty;
                    html.AppendLine($"<td>{val.Replace("\n", "<br>")}</td>");
                }
                html.AppendLine("</tr>");
            }
        }
        html.AppendLine("</tbody></table>");
        html.AppendLine($"<div class='footer'>Adressen &amp; Kontakte, <a href='http://www.netradio.info' style='color: #555; text-decoration: none;'>www.netradio.info</a></div>");
        html.AppendLine("</body></html>");

        // --- Dateiverwaltung & Ausgabe ---
        var tempFolder = Path.Combine(Path.GetTempPath(), "AdressenApp_Exports");
        if (!Directory.Exists(tempFolder)) { Directory.CreateDirectory(tempFolder); }
        else
        {
            var oldFiles = Directory.GetFiles(tempFolder, "*.html");
            foreach (var file in oldFiles)
            {
                try { File.Delete(file); }
                catch { }  // überspringen, falls Datei noch im Browser geöffnet ist
            }
        }
        var tempPath = Path.Combine(tempFolder, $"Export_{DateTime.Now:yyyyMMdd_HHmmss}.html");
        File.WriteAllText(tempPath, html.ToString());

        var startInfo = new ProcessStartInfo
        {
            FileName = tempPath,
            UseShellExecute = true
        };
        Process.Start(startInfo);
    }
}