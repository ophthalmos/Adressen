using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Globalization;
using System.IO.Compression;
using System.Media;
using System.Net;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Win32;

namespace Adressen.cls;

internal static partial class Utils
{
    [GeneratedRegex(@"[^\d+]")]
    private static partial Regex NonDigitOrPlusRegex();
    private const string runLocation = @"Software\Microsoft\Windows\CurrentVersion\Run";

    public static void MsgTaskDlg(nint hwnd, string heading, string message, TaskDialogIcon? icon = null)
    {
        TaskDialog.ShowDialog(hwnd, new TaskDialogPage() { Caption = Application.ProductName, SizeToContent = true, Heading = heading, Text = message, Icon = icon ?? TaskDialogIcon.None, AllowCancel = true, Buttons = { TaskDialogButton.OK } });
    }

    public static void ErrTaskDlg(nint? hwnd, Exception error)
    {
        TaskDialogPage page = new()
        {
            Caption = Application.ProductName,
            Heading = error.GetType().ToString(),
            Text = error.Message,
            Icon = TaskDialogIcon.Error,
            SizeToContent = true,
            AllowCancel = true,
            Buttons = { TaskDialogButton.OK },
            Expander = new TaskDialogExpander()
            {
                Text = error.ToString(),
                CollapsedButtonText = "Technische Details anzeigen",
                ExpandedButtonText = "Details ausblenden",
                Position = TaskDialogExpanderPosition.AfterFootnote
            }
        };
        TaskDialog.ShowDialog(hwnd ?? 0, page);
    }

    public static async Task<bool> RunWithProgressDialogAsync(IWin32Window owner, string caption, string text, Func<CancellationToken, Task> work)
    {
        using var cts = new CancellationTokenSource();
        var btnCancel = TaskDialogButton.Cancel;

        var pageProgress = new TaskDialogPage()
        {
            Caption = caption,
            Heading = "Bitte warten…",
            Text = text,
            Icon = TaskDialogIcon.None,
            SizeToContent = true,
            ProgressBar = new TaskDialogProgressBar() { State = TaskDialogProgressBarState.Marquee },
            Buttons = { btnCancel }
        };
        btnCancel.Click += (s, e) => { cts.Cancel(); };
        var success = false;

        pageProgress.Created += async (s, args) =>
        {
            try
            {
                if (owner is Control c) { c.Cursor = Cursors.WaitCursor; }  // Cursor auf dem Owner ändern, falls möglich 
                await work(cts.Token);
                success = true;
                pageProgress.BoundDialog?.Close();
            }
            catch (OperationCanceledException) { pageProgress.BoundDialog?.Close(); }
            catch (Exception ex)
            {
                pageProgress.BoundDialog?.Close();
                ErrTaskDlg(owner.Handle, ex);
            }
            finally
            {
                if (owner is Control c) { c.Cursor = Cursors.Default; }
            }
        };
        TaskDialog.ShowDialog(owner, pageProgress);
        return success;
    }

    internal static void SetAutoStart(string? appName, string assemblyLocation, string? commandLineArgs)
    {
        using var key = Registry.CurrentUser.CreateSubKey(runLocation);
        key?.SetValue(appName, $"{assemblyLocation} {commandLineArgs}".Trim());
    }

    internal static (bool IsEnabled, bool HasMin2Tray) IsAutoStartEnabled(string? appName, string assemblyLocation)
    {
        using var key = Registry.CurrentUser.OpenSubKey(runLocation);
        if (key == null) { return (false, false); }
        var value = key.GetValue(appName) as string;
        if (string.IsNullOrEmpty(value)) { return (false, false); }
        var isEnabled = value.StartsWith(assemblyLocation, StringComparison.OrdinalIgnoreCase);
        var hasMin2Tray = isEnabled && value.Contains("-min2Tray", StringComparison.OrdinalIgnoreCase);
        return (isEnabled, hasMin2Tray);
    }

    internal static void UnSetAutoStart(string? appName)
    {
        using var key = Registry.CurrentUser.OpenSubKey(runLocation, writable: true);
        key?.DeleteValue(appName ?? "Adressen", throwOnMissingValue: false);
    }

    public static void SortContacts(BindingList<Contact>? contacts)
    {
        if (contacts == null || contacts.Count == 0) { return; }
        var sortedList = contacts
            .OrderBy(x => x.Nachname, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(x => x.Vorname, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(x => x.Unternehmen, StringComparer.CurrentCultureIgnoreCase)
            .ToList();

        // Um Event-Spam zu vermeiden, schalten wir die Benachrichtigung kurz ab
        contacts.RaiseListChangedEvents = false;
        try
        {
            contacts.Clear();
            foreach (var c in sortedList) { contacts.Add(c); }
        }
        finally
        {
            contacts.RaiseListChangedEvents = true;
            contacts.ResetBindings(); // Ein einziges Event für die ganze Liste
        }
    }

    public static int GetAddressInsertIndex(BindingSource source, Adresse newItem)
    {
        var compareInfo = CultureInfo.InvariantCulture.CompareInfo;  // SQLite's NOCASE nur ASCII-Werte 
        var options = CompareOptions.IgnoreCase | CompareOptions.StringSort;  // behandelt Sonderzeichen und Leerzeichen korrekt
        for (var i = 0; i < source.Count; i++)
        {
            if (source[i] is Adresse current)
            {
                var cmp = compareInfo.Compare(newItem.Nachname ?? "", current.Nachname ?? "", options);
                if (cmp == 0) { cmp = compareInfo.Compare(newItem.Vorname ?? "", current.Vorname ?? "", options); }
                if (cmp == 0) { cmp = compareInfo.Compare(newItem.Unternehmen ?? "", current.Unternehmen ?? "", options); }
                if (cmp < 0) { return i; }  // Wenn cmp < 0, ist newItem alphabetisch VOR current.
            }
        }
        return source.Count;
    }

    public static void SortAddresses(BindingSource source)
    {
        // Wir arbeiten auf der aktuellen Liste der Source
        if (source.List is not IEnumerable<Adresse> currentItems) { return; }

        // Wir nutzen die gleiche Culture wie in deiner SQLite-Verbindung (de-DE)
        var culture = new CultureInfo("de-DE");

        var sorted = currentItems
            .OrderBy(a => a.Nachname ?? "", StringComparer.Create(culture, true))
            .ThenBy(a => a.Vorname ?? "", StringComparer.Create(culture, true))
            .ThenBy(a => a.Unternehmen ?? "", StringComparer.Create(culture, true))
            .ToList();

        // DER TRICK: Nicht die Liste leeren, sondern die DataSource der BindingSource tauschen.
        // Das ist in .NET 10 nahezu instantan.
        source.DataSource = new BindingList<Adresse>(sorted);
    }

    public static List<(DateOnly Datum, string Name, int Alter, int Tage, string Id)> CalculateUpcomingBirthdays(IEnumerable<IContactEntity> contacts, int daysLookBack, int daysLookAhead)
    {
        var heute = DateOnly.FromDateTime(DateTime.Today);
        var result = new List<(DateOnly Datum, string Name, int Alter, int Tage, string Id)>();
        foreach (var contact in contacts.Where(c => c.BirthdayDate.HasValue && c.Reminder))
        {
            var gebDatum = contact.BirthdayDate!.Value;
            var targetYear = heute.Year;

            // 1. Tag für das aktuelle Jahr ermitteln (inklusive Schaltjahr-Korrektur)
            var day = (gebDatum.Month == 2 && gebDatum.Day == 29 && !DateTime.IsLeapYear(targetYear)) ? 28 : gebDatum.Day;
            var gebTagDiesesJahr = new DateOnly(targetYear, gebDatum.Month, day);
            var tage = gebTagDiesesJahr.DayNumber - heute.DayNumber;

            // 2. Jahreswechsel-Logik
            if (tage < -daysLookBack)
            {
                // Geburtstag war schon vor langer Zeit, wir schauen aufs nächste Jahr
                targetYear = heute.Year + 1;
                var dayNext = (gebDatum.Month == 2 && gebDatum.Day == 29 && !DateTime.IsLeapYear(targetYear)) ? 28 : gebDatum.Day;
                tage = new DateOnly(targetYear, gebDatum.Month, dayNext).DayNumber - heute.DayNumber;
            }
            else if (tage > daysLookAhead)
            {
                // Geburtstag ist zu weit weg im aktuellen Jahr. Vielleicht schauen wir im Januar auf den Dezember des Vorjahres zurück?
                var prevYear = heute.Year - 1;
                var dayPrev = (gebDatum.Month == 2 && gebDatum.Day == 29 && !DateTime.IsLeapYear(prevYear)) ? 28 : gebDatum.Day;
                var prevTage = new DateOnly(prevYear, gebDatum.Month, dayPrev).DayNumber - heute.DayNumber;

                if (prevTage >= -daysLookBack)
                {
                    targetYear = prevYear;
                    tage = prevTage;
                }
            }

            // 3. Wenn es ins Zeitfenster passt, Alter berechnen und hinzufügen
            if (tage >= -daysLookBack && tage <= daysLookAhead)
            {
                // Das Alter ist schlicht das Zieljahr der Feier minus das Geburtsjahr
                var alter = targetYear - gebDatum.Year;
                result.Add((gebDatum, contact.DisplayName, alter, tage, contact.UniqueId));
            }
        }

        return [.. result.OrderBy(x => x.Tage)];
    }

    internal static Font GetSafeFont(string fontName, float fontSize, FontFamily fallbackFamily)
    {
        try
        {
            var font = new Font(fontName, fontSize);

            // Wenn Windows die Schriftart nicht findet, wird automatisch "Microsoft Sans Serif" genommen.
            // Das prüfen wir hier, um stattdessen unseren gewünschten Fallback zu nutzen.
            if (!string.Equals(font.Name, fontName, StringComparison.OrdinalIgnoreCase))
            {
                font.Dispose();
                return new Font(fallbackFamily, fontSize);
            }

            return font;
        }
        catch
        {
            // Falls ein anderer Fehler beim Instanziieren auftritt
            return new Font(fallbackFamily, fontSize);
        }
    }

    internal static void StartFile(nint handle, string filePath)
    {
        try
        {
            if (File.Exists(filePath))
            {
                ProcessStartInfo psi = new(filePath) { UseShellExecute = true, WorkingDirectory = Path.GetDirectoryName(filePath) };
                Process.Start(psi);
            }
            else { MsgTaskDlg(handle, "Datei nicht gefunden!", "'" + filePath + "' fehlt.", TaskDialogIcon.ShieldWarningYellowBar); }
        }
        catch (Exception ex) when (ex is Win32Exception || ex is InvalidOperationException) { ErrTaskDlg(handle, ex); }
    }

    internal static void StartLink(nint handle, string url)
    {
        try
        {
            if (Uri.TryCreate(url, UriKind.Absolute, out var uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps))
            {
                var psi = new ProcessStartInfo(url) { UseShellExecute = true };
                Process.Start(psi);
            }
            else { MsgTaskDlg(handle, "Ungültiger Link!", $"'{url}' ist keine gültige URL.", TaskDialogIcon.ShieldWarningYellowBar); }
        }
        catch (Exception ex) { ErrTaskDlg(handle, ex); }
    }

    internal static async Task<bool> GoogleConnectionCheckAsync(nint hwnd, string path)
    {
        try
        {
            var ping = new Ping();
            // Löst keinen UI-Freeze aus, auch wenn es 1000ms dauert
            var reply = await ping.SendPingAsync(new IPAddress([8, 8, 8, 8]), 1000);

            if (reply.Status != IPStatus.Success)
            {
                MsgTaskDlg(hwnd, "Keine Internetverbindung!", "Überprüfe das Netzwerk.", TaskDialogIcon.ShieldWarningYellowBar);
                return false;
            }
        }
        catch
        {
            // Fängt Ausnahmen ab, z.B. wenn gar keine Netzwerkkarte aktiv ist
            MsgTaskDlg(hwnd, "Keine Internetverbindung!", "Überprüfe das Netzwerk.", TaskDialogIcon.ShieldWarningYellowBar);
            return false;
        }

        if (!File.Exists(path))
        {
            MsgTaskDlg(hwnd, "Der Key-File wurde nicht gefunden!", $"'{path}' fehlt.", TaskDialogIcon.ShieldWarningYellowBar);
            return false;
        }

        return true;
    }

    internal static IEnumerable<Control> GetAllControls(Control container)
    {
        foreach (Control c in container.Controls)
        {
            yield return c; // Gib das aktuelle Control zurück
            foreach (var child in GetAllControls(c)) { yield return child; }
        }
    }

    internal static string GenerateDetailedDiff(Contact current, Contact old, string[] fields)
    {
        var sb = new StringBuilder();
        var type = typeof(Contact);

        foreach (var fieldName in fields)
        {
            // PropertyInfo holen
            var prop = type.GetProperty(fieldName);
            if (prop == null) { continue; } // Sollte nicht passieren, wenn Array korrekt ist

            var valOld = prop.GetValue(old);
            var valNew = prop.GetValue(current);

            // Unterscheidung nach Typ für korrekte Formatierung/Vergleich
            if (prop.PropertyType == typeof(string))
            {
                // Strings normalisieren (null == empty)
                var sOld = (valOld as string) ?? string.Empty;
                var sNew = (valNew as string) ?? string.Empty;

                if (!string.Equals(sOld, sNew, StringComparison.Ordinal))
                {
                    if (fieldName == nameof(Contact.Notizen))
                    {
                        var status = string.Empty;
                        if (sOld == string.Empty && sNew != string.Empty) { status = "Text hinzugefügt"; }
                        else if (sOld != string.Empty && sNew == string.Empty) { status = "Text gelöscht"; }
                        else { status = "Text geändert"; }
                        sb.AppendLine($"Notizen: {status}");
                    }
                    else  // Reguläres Verhalten für alle anderen Strings
                    {
                        var displayOld = sOld == string.Empty ? "[Leer]" : sOld;
                        var displayNew = sNew == string.Empty ? "∅" : sNew;
                        sb.AppendLine($"{fieldName}: {displayOld} ➔ {displayNew}");
                    }
                }
            }
            else // z.B. Datum (Geburtstag) oder Zahlen
            {
                if (!Equals(valOld, valNew)) { sb.AppendLine($"{fieldName}: {FormatObj(valOld)} ➔ {FormatObj(valNew)}"); }
            }
        }

        // Spezialbehandlung für Gruppen (da diese nicht in dataFields stehen)
        var oldGroups = old.GroupNames ?? [];
        var newGroups = current.GroupNames ?? [];

        if (!oldGroups.OrderBy(x => x).SequenceEqual(newGroups.OrderBy(x => x)))
        {
            var displayOldGroups = oldGroups.Count == 0 ? "[Keine]" : string.Join(", ", oldGroups);
            var displayNewGroups = newGroups.Count == 0 ? "[Keine]" : string.Join(", ", newGroups);
            sb.AppendLine($"Gruppen: {displayOldGroups} ➔ {displayNewGroups}");
        }

        return sb.ToString().TrimEnd(); // TrimEnd entfernt den letzten überflüssigen Zeilenumbruch

        // Lokale Funktion zur Formatierung von Nicht-Strings
        static string FormatObj(object? o)
        {
            if (o == null) { return "[Leer]"; }
            if (o is DateTime d) { return d.ToShortDateString(); }
            if (o is DateOnly dO) { return dO.ToString(); }
            return o.ToString() ?? string.Empty;
        }
    }

    internal static bool IsFileReady(string filename)
    {
        try
        {
            using var stream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.None);
            return true;
        }
        catch (IOException) { return false; }  // Die Datei wird gerade noch von einem anderen Prozess geschrieben oder kopiert
    }

    internal static GraphicsPath GetRoundedRectanglePath(Rectangle bounds, int radius)
    {
        var path = new GraphicsPath();
        if (radius <= 0)
        {
            path.AddRectangle(bounds);
            return path;
        }
        var diameter = radius * 2;
        var size = new Size(diameter, diameter);
        var arc = new Rectangle(bounds.Location, size);
        path.AddArc(arc, 180, 90);  // Oben Links
        arc.X = bounds.Right - diameter;  // Oben Rechts
        path.AddArc(arc, 270, 90);
        arc.Y = bounds.Bottom - diameter;  // Unten Rechts
        path.AddArc(arc, 0, 90);
        arc.X = bounds.Left;  // Unten Links
        path.AddArc(arc, 90, 90);
        path.CloseFigure();
        return path;
    }

    internal static void StartSearchCacheWarmup(IEnumerable<IContactEntity> items)
    {
        // Wir erstellen erst einen Snapshot (Array), solange wir noch im UI-Thread sind.
        // Das verhindert Abstürze, wenn sich die Original-Liste während des Warmups ändert.
        var snapshot = items.ToArray();
        _ = Task.Run(() => { foreach (var item in snapshot) { _ = item.SearchText; } });
    }

    public static IEnumerable<string[]> ReadCsv(string filePath)
    {
        using var reader = new StreamReader(filePath, Encoding.UTF8);
        var currentFields = new List<string>();
        var currentField = new StringBuilder();
        var inQuotes = false;

        while (reader.Peek() >= 0)
        {
            var line = reader.ReadLine();
            if (line == null)
            {
                break;
            }

            for (var i = 0; i < line.Length; i++)
            {
                var c = line[i];

                if (c == '"')
                {
                    // Wenn wir in Quotes sind und das nächste Zeichen auch ein Quote ist -> Escaped Quote ""
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++; // Das zweite Quote überspringen
                    }
                    else
                    {
                        inQuotes = !inQuotes; // Quote-Modus umschalten
                    }
                }
                else if (c == ';' && !inQuotes)
                {
                    currentFields.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }

            if (!inQuotes)
            {
                // Zeile zu Ende und nicht in Quotes -> Datensatz komplett
                currentFields.Add(currentField.ToString());
                yield return [.. currentFields];
                currentFields.Clear();
                currentField.Clear();
            }
            else
            {
                // Zeilenumbruch innerhalb von Quotes -> \n hinzufügen und weiterlesen
                currentField.Append(Environment.NewLine);
            }
        }
    }

    internal static void WendeExifOrientierungAn(Image bild)
    {
        const int ExifOrientationId = 0x112;  // PropertyTagOrientation (ID: 0x0112 = 274)
        if (bild.PropertyIdList.Contains(ExifOrientationId))
        {
            var item = bild.GetPropertyItem(ExifOrientationId);
            if (item is null || item.Value is null || item.Value.Length == 0) { return; } // Frühzeitiger Abbruch, falls null oder leer
            var rotation = RotateFlipType.RotateNoneFlipNone; // Standardwert   
            switch (item.Value[0])
            {
                case 1: rotation = RotateFlipType.RotateNoneFlipNone; break;
                case 2: rotation = RotateFlipType.RotateNoneFlipX; break;
                case 3: rotation = RotateFlipType.Rotate180FlipNone; break;
                case 4: rotation = RotateFlipType.Rotate180FlipX; break;
                case 5: rotation = RotateFlipType.Rotate90FlipX; break;
                case 6: rotation = RotateFlipType.Rotate90FlipNone; break; // Hochkant-Foto
                case 7: rotation = RotateFlipType.Rotate270FlipX; break;
                case 8: rotation = RotateFlipType.Rotate270FlipNone; break; // Hochkant-Foto
            }
            if (item.Value[0] != 1) { bild.RotateFlip(rotation); }  // Wir drehen nur, wenn es nicht der normale Zustand (1) ist
            bild.RemovePropertyItem(ExifOrientationId); // Orientierungs-Tag wird entfernt, sicherer falls noch als JPEG gespeichert wird
        }
    }

    internal static Image SkaliereBildDaten(Image originalBild, int neueBreite)
    {
        var originalBreite = originalBild.Width;
        if (originalBreite <= neueBreite) { return (Image)originalBild.Clone(); }
        var originalHoehe = originalBild.Height;
        var neueHoehe = (int)((double)originalHoehe / originalBreite * neueBreite);
        var neuesBild = new Bitmap(neueBreite, neueHoehe);
        using (var graphics = Graphics.FromImage(neuesBild))
        {
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.DrawImage(originalBild, new Rectangle(0, 0, neueBreite, neueHoehe));
        }
        return neuesBild; // Rückgabe der neuen Bitmap
    }

    internal static Image BeschneideZuQuadrat(Image originalBild, bool? priority = false)  // null = Oben, true = Unten, false = Mitte 
    {
        var breite = originalBild.Width;
        var hoehe = originalBild.Height;
        if (hoehe <= breite) { return (Image)originalBild.Clone(); }
        var yOffset = priority == null ? 0 : priority == true ? hoehe - breite : (hoehe - breite) / 2;
        var rechteck = new Rectangle(0, yOffset, breite, breite); // Ausschnittsquadrat, Höhe = Breite, yOffset je nach Priorität
        var quadratischesBild = new Bitmap(breite, breite); // Korrekt: Kein 'using'
        using (var graphics = Graphics.FromImage(quadratischesBild))
        {
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.DrawImage(originalBild, new Rectangle(0, 0, breite, breite), rechteck, GraphicsUnit.Pixel);
        }
        return quadratischesBild; // Rückgabe der neuen Bitmap
    }

    internal static Image ReduziereWieGoogle(Image originalBild, int newHeight)
    {
        var originalHeight = originalBild.Height;
        if (originalHeight <= newHeight) { return (Image)originalBild.Clone(); }
        var originalWidth = originalBild.Width;
        var newWidth = (int)((double)originalWidth / originalHeight * newHeight);
        var neuesBild = new Bitmap(newWidth, newHeight); // KEIN 'using'
        using (var graphics = Graphics.FromImage(neuesBild))
        {
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.DrawImage(originalBild, new Rectangle(0, 0, newWidth, newHeight));
        }
        return neuesBild; // Rückgabe der neuen Bitmap
    }


    internal static string FormatBytes(long bytes)  // Effizienter (Loop statt Logarithmen)
    {
        string[] suffix = ["Bytes", "KB", "MB", "GB", "TB"];
        if (bytes == 0) { return "0 " + suffix[0]; }
        var i = 0;
        double dBytes = bytes;
        while (dBytes >= 1024 && i < suffix.Length - 1)
        {
            dBytes /= 1024;
            i++;
        }
        return $"{dBytes.ToString("F2", CultureInfo.GetCultureInfo("de-DE"))} {suffix[i]}";  // Verwendet die de-DE Kultur für das Komma
    }

    internal static void StartDir(nint handle, string dirPath)
    {
        try
        {
            if (Directory.Exists(dirPath))
            {
                ProcessStartInfo psi = new(dirPath) { UseShellExecute = true, WorkingDirectory = Path.GetDirectoryName(dirPath) };
                Process.Start(psi);
            }
        }
        catch (Exception ex) when (ex is Win32Exception || ex is InvalidOperationException) { ErrTaskDlg(handle, ex); }
    }

    public static bool IsPrinterAvailable(string printerName)
    {
        foreach (string installedPrinter in PrinterSettings.InstalledPrinters)
        {
            if (string.Equals(installedPrinter, printerName, StringComparison.OrdinalIgnoreCase)) { return true; }
        }
        return false;
    }

    public static async Task<(Version? Version, string? ReleaseDate)> GetLatestVersionInfoAsync(CancellationToken ct = default)
    {
        var xmlUrl = "https://www.netradio.info/download/adressen.xml";
        try
        {
            using var requestMessage = new HttpRequestMessage(HttpMethod.Get, xmlUrl);
            requestMessage.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));

            using var response = await HttpService.Client.SendAsync(requestMessage, ct);
            response.EnsureSuccessStatusCode();

            var xmlContent = await response.Content.ReadAsStringAsync(ct);
            var doc = XDocument.Parse(xmlContent);

            var rawVersion = doc.Element("adressen")?.Element("version")?.Value;
            var releaseDate = doc.Element("adressen")?.Element("date")?.Value;

            if (!string.IsNullOrEmpty(rawVersion))
            {
                var cleanVersionString = rawVersion.Split(['+', '-'])[0];
                if (!cleanVersionString.Contains('.')) { cleanVersionString += ".0"; }

                if (Version.TryParse(cleanVersionString, out var parsedVersion)) { return (parsedVersion, releaseDate); }
            }
        }
        catch (OperationCanceledException) { } // Abbruch durch CancellationToken, nichts tun
        catch (Exception ex) { Debug.WriteLine($"Fehler beim Abrufen der Versionsinfo: {ex.Message}"); }
        return (null, null);
    }

    public static bool IsUpdateCheckDue(int updateIndex, DateTime lastUpdateCheck)
    {
        if (updateIndex == 3) { return false; }  // "Niemals"
        var elapsed = DateTime.Now - lastUpdateCheck;
        return updateIndex switch
        {
            0 => elapsed.TotalDays >= 1,  // Jeden Tag
            1 => elapsed.TotalDays >= 7,  // Jede Woche
            2 => elapsed.TotalDays >= 30, // Jeden Monat
            _ => false
        };
    }

    internal static void HelpMsgTaskDlg(nint hwnd, string appName, Icon? icon, int? dbVersion = null)
    {

        var curVersion = Assembly.GetExecutingAssembly().GetName().Version;
        var threeVersion = curVersion?.ToString(3) ?? "unbekannt"; //curVersion is not null ? $"{curVersion.Major}.{curVersion.Minor}.{curVersion.Build}" : "unbekannt";
        var buildDate = GetBuildDate();
        TaskDialogButton paypalButton = new TaskDialogCommandLinkButton("Anerkennung spenden via PayPal");
        //TaskDialogButton updateButton = new TaskDialogCommandLinkButton("Nach Programm-Update suchen…") { AllowCloseDialog = false };
        var indent = new string(' ', 14);
        var foot = $"{indent}© {buildDate:yyyy} Wilhelm Happe\n{indent}Version {threeVersion} ({buildDate:d})";
        if (dbVersion.HasValue) { foot += $"\n{indent}Datenbank-Schema: v{dbVersion.Value}"; }
        var msg = "Adressverwaltung für die komfortable Zusammen-" + Environment.NewLine +
            "arbeit mit Microsoft-Word und LibreOffice-Writer" + Environment.NewLine +
            "und der Möglichkeit, Briefumschläge zu bedrucken." + Environment.NewLine +
            "Neben den lokal gespeicherten Adressen können" + Environment.NewLine + "Google-Kontakte geladen und verwendet werden.";
        var initialPage = new TaskDialogPage()
        {
            Caption = "Über " + appName,
            Heading = appName,
            Text = msg,
            Icon = icon == null ? null : new TaskDialogIcon(icon),
            AllowCancel = true,
            SizeToContent = true,
            Buttons = { paypalButton, TaskDialogButton.OK },
            DefaultButton = TaskDialogButton.OK,
            Footnote = foot
        };
        var result = TaskDialog.ShowDialog(hwnd, initialPage);
        if (result == paypalButton) { StartLink(hwnd, "https://www.paypal.com/donate/?hosted_button_id=3HRQZCUW37BQ6"); }
        //else if (result == downloadButton) { StartLink(hwnd, urlString); }
    }

    internal static bool? AskWordProcessingProgram(nint hwnd)
    {
        TaskDialogButton wordButton = new TaskDialogCommandLinkButton("Microsoft Word");
        TaskDialogButton libreButton = new TaskDialogCommandLinkButton("LibreOffice Writer");
        //var settingsButton = new TaskDialogButton("Einstellungen öffnen");  // geht nicht gleichzeitig mit CommandLinkButtons
        var page = new TaskDialogPage
        {
            Caption = Application.ProductName,
            Heading = "Wähle die Textverarbeitung",
            Text = "Du kannst deine bevorzugte Auswahl in den Einstellungen\nfestlegen, so dass dieser Dialog nicht mehr angezeigt wird.",  // \nUm die Einstellungen zu öffnen, kannst du 'OK' klicken.",
            Buttons = { wordButton, libreButton, TaskDialogButton.Cancel },
            AllowCancel = true,
            SizeToContent = true
        };
        var result = TaskDialog.ShowDialog(hwnd, page);
        if (result == wordButton) { return true; }
        if (result == libreButton) { return false; }
        //if (result == settingsButton) { return null; }  // null wird schon für 'Abbrechen' verwendet, daher verzichten wir auf diesen Button
        return null;
    }

    internal static (bool IsYes, bool IsCancelled) YesNo_TaskDialog(IWin32Window? owner, string caption, string heading, string text, string yes = "", string no = "")
    {
        var yesButton = string.IsNullOrEmpty(yes) ? TaskDialogButton.Yes : new TaskDialogButton(yes);
        var noButton = string.IsNullOrEmpty(no) ? TaskDialogButton.No : new TaskDialogButton(no);
        using var customIcon = Properties.Resources.question32;         // Beide Instanzen sauber kapseln,
        using var questionDialogIcon = new TaskDialogIcon(customIcon);  // damit keine GDI-Leaks entstehen
        var page = new TaskDialogPage
        {
            Caption = caption,
            Heading = heading,
            Text = text,
            Icon = questionDialogIcon,
            Buttons = { yesButton, noButton },
            AllowCancel = true,
            SizeToContent = true
        };
        var result = owner is not null ? TaskDialog.ShowDialog(owner, page) : TaskDialog.ShowDialog(page);
        var isYes = result == yesButton;
        var isCancelled = result == TaskDialogButton.Cancel;
        return (isYes, isCancelled);
    }

    internal static string SanitizeFileName(string fileName)  // Entfernt alle ungültigen Zeichen und reduziert Punkte auf einen
    {
        var invalidChars = Regex.Escape(new string(Path.GetInvalidFileNameChars()));
        var invalidRegStr = string.Format(@"([{0}]*\.{{2,}})|([{0}]+)", invalidChars);
        var clean = Regex.Replace(fileName, invalidRegStr, match => match.Value.Contains("..") ? "." : "");
        return clean;
    }

    internal static string NormalizePhoneNumber(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) { return string.Empty; }
        var cleaned = NonDigitOrPlusRegex().Replace(input, string.Empty);  // Alle Zeichen außer Ziffern und das Plus entfernen
        if (cleaned.StartsWith('+')) { cleaned = "00" + cleaned[1..]; }  // Führendes '+' vereinheitlichen zu '00'
        // DACH-Länderpräfixe in nationale 0-Vorwahl umwandeln. Dabei fehlerhafte Formate wie +49(0)... direkt korrigieren
        if (cleaned.StartsWith("00490")) { cleaned = "0" + cleaned[5..]; }
        else if (cleaned.StartsWith("0049")) { cleaned = "0" + cleaned[4..]; }
        else if (cleaned.StartsWith("00430")) { cleaned = "0" + cleaned[5..]; }
        else if (cleaned.StartsWith("0043")) { cleaned = "0" + cleaned[4..]; }
        else if (cleaned.StartsWith("00410")) { cleaned = "0" + cleaned[5..]; }
        else if (cleaned.StartsWith("0041")) { cleaned = "0" + cleaned[4..]; }
        return new string([.. cleaned.Where(char.IsDigit)]);  // Sicherstellen, dass absolut keine Fremdzeichen (z.B. verbliebene '+') enthalten sind
    }

    internal static void PlayIncomingCallSound(string appPath)
    {
        var file = Path.Combine(appPath, "ringing.wav");
        if (!string.IsNullOrEmpty(file) && File.Exists(file))
        {
            Task.Run(() => { using var player = new SoundPlayer(file); player.PlaySync(); });  // Blockierend auf Thread-Pool – UI bleibt frei, Dispose erst nach Ende
        }
        else { SystemSounds.Asterisk.Play(); }  // Windows-Systemton, immer verfügbar
    }

    internal static (bool askBefore, bool deleteNow) AskBeforeDeleteContact(nint handle, IContactEntity contact, bool askBeforeDelete, bool showVerification = true)
    {
        var deleteNow = false;
        try
        {
            var details = contact.DisplayName;  // Wir holen die Daten direkt vom Objekt, nicht aus den Grid-Zellen. Das ist viel schneller und weniger fehleranfällig
            var zusatzInfo = "";
            if (contact is Contact c) { zusatzInfo = $"\n{c.Unternehmen}\n{c.Strasse}\n{c.PLZ} {c.Ort}"; }
            else if (contact is Adresse a) { zusatzInfo = $"\n{a.Unternehmen}\n{a.Strasse}\n{a.PLZ} {a.Ort}"; }
            using var customIcon = Properties.Resources.question32;         // Beide Instanzen sauber kapseln,
            using var questionDialogIcon = new TaskDialogIcon(customIcon);  // damit keine GDI-Leaks entstehen
            var page = new TaskDialogPage()
            {
                Heading = "Möchtest du den Datensatz löschen?",
                Text = (details + zusatzInfo).Trim(),
                Caption = Application.ProductName,
                Icon = questionDialogIcon,
                AllowCancel = true,
                SizeToContent = true,
                Verification = showVerification ? new TaskDialogVerificationCheckBox() { Text = "Immer fragen" } : null,
                Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
            };
            if (page.Verification is TaskDialogVerificationCheckBox check) { check.Checked = askBeforeDelete; }
            var resultButton = TaskDialog.ShowDialog(handle, page);

            // Logik für die Checkbox
            if (page.Verification is TaskDialogVerificationCheckBox finalCheck)
            {
                if (askBeforeDelete && !finalCheck.Checked)
                {
                    MsgTaskDlg(page.BoundDialog?.Handle ?? 0, "Hinweis", "Du kannst die Sicherheitsabfrage in\nden Einstellungen wieder einschalten.", new(Properties.Resources.info32));
                    askBeforeDelete = false;
                }
                else if (finalCheck.Checked) { askBeforeDelete = true; }
            }
            if (resultButton == TaskDialogButton.Yes) { deleteNow = true; }
        }
        catch (Exception ex) { ErrTaskDlg(handle, ex); }
        return (askBeforeDelete, deleteNow);
    }

    internal static (bool askBefore, bool deleteNow) AskBeforeDeleteAddress(nint hwnd, Adresse adresse, bool askBeforeDelete, bool showVerification = true)
    {
        var deleteNow = false;
        try
        {
            var vorname = adresse.Vorname ?? string.Empty;
            var nachname = adresse.Nachname ?? string.Empty;
            var unternehmen = adresse.Unternehmen ?? string.Empty;
            var strasse = adresse.Strasse ?? string.Empty;
            var plz = adresse.PLZ ?? string.Empty;
            var ort = adresse.Ort ?? string.Empty;
            using var customIcon = Properties.Resources.question32;         // Beide Instanzen sauber kapseln,
            using var questionDialogIcon = new TaskDialogIcon(customIcon);  // damit keine GDI-Leaks entstehen

            var page = new TaskDialogPage()
            {
                Heading = "Möchtest du den Datensatz löschen?",
                Text = $"{vorname} {nachname}\n{unternehmen}\n{strasse}\n{plz} {ort}".Trim(),
                Caption = Application.ProductName,
                Icon = questionDialogIcon,
                AllowCancel = true,
                SizeToContent = true,
                Verification = showVerification ? new TaskDialogVerificationCheckBox() { Text = "Immer fragen" } : null, // Korrigiert: null statt ""
                Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
            };

            // Korrigiert: Sicherer Null-Check für die Zuweisung
            if (page.Verification is TaskDialogVerificationCheckBox check) { check.Checked = askBeforeDelete; }
            var resultButton = TaskDialog.ShowDialog(hwnd, page);

            // Korrigiert: Sicheres Auslesen des Ergebnisses
            if (page.Verification is TaskDialogVerificationCheckBox finalCheck)
            {
                if (askBeforeDelete && !finalCheck.Checked)
                {
                    MsgTaskDlg(hwnd, "Hinweis", "Du kannst die Sicherheitsabfrage in\nden Einstellungen wieder einschalten.", new(Properties.Resources.info32));
                    askBeforeDelete = false;
                }
                else if (finalCheck.Checked) { askBeforeDelete = true; }
            }
            if (resultButton == TaskDialogButton.Yes) { deleteNow = true; }
        }
        catch (Exception ex) { ErrTaskDlg(hwnd, ex); }
        return (askBeforeDelete, deleteNow);
    }

    internal static bool TryParseInput(string? text, out DateTime date) => DateTime.TryParseExact(text?.Trim(), ["d.M.yy", "dd.MM.yyyy", "d.M.yyyy", "dd.MM.yy"], CultureInfo.GetCultureInfo("de-DE"), DateTimeStyles.None, out date);

    internal struct DateDiff
    {
        public int years, months, days;
    }

    internal static DateDiff CalcDateDiff(DateTime d1, DateTime d2)
    {// toDate muss immer vor fromDate liegen (toDate < fromDate), ansonsten liefert die Funktion falsche Werte!
        int years, months, days;
        if (d2 < d1) { (d1, d2) = (d2, d1); }
        years = d2.Year - d1.Year;
        var dt = d1.AddYears(years);
        if (dt > d2)
        {
            years--;
            dt = d1.AddYears(years);
        }
        months = d2.Month - d1.Month;
        if (d2.Day < d1.Day) { months--; }
        months = (months + 12) % 12;
        dt = dt.AddMonths(months);
        if (months == 1) { dt = dt.AddMonths(-1); months = 0; } // 30.8.20 neu eingefügt
        days = (d2 - dt).Days;
        DateDiff ddf;
        ddf.years = years; ddf.months = months; ddf.days = days;
        return ddf;
    }

    internal static bool IsInnoSetupValid(string appPath)
    {
        if (appPath.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))) { return true; }
        var appDir = Path.GetDirectoryName(appPath);
        if (appDir is null) { return false; }
        if (File.Exists(Path.Combine(appDir, "unins000.exe"))) { return true; }
        //var localSettings = Path.ChangeExtension(appPath, ".json");
        //if (File.Exists(localSettings)) { return false; } // Existiert bereits eine lokale Einstellungsdatei? (typisch für Portable)
        return false;
    }

    internal static string CorrectUNC(string unc)
    {
        if (string.IsNullOrWhiteSpace(unc)) { return string.Empty; }
        // Wenn es kein lokaler Pfad (C:\) und kein relativer Pfad ist, aber mit einem Backslash startet, erzwingen wir genau zwei.
        if (unc.StartsWith('\\') && !unc.StartsWith(@"\\")) { return @"\\" + unc.TrimStart('\\'); }
        return unc;
    }

    internal static bool SetClipboardText(string text)
    {
        try
        {// It retries 5 times with 250 milliseconds between each retry
            Clipboard.SetDataObject(text, false, 5, 250);
            return true;
        }
        catch (Exception ex) when (ex is ExternalException) { return false; }
    }

    private static DateTime GetBuildDate()
    { //s. <SourceRevisionId>build$([System.DateTime]::UtcNow.ToString("yyyyMMddHHmmss"))</SourceRevisionId> in ClipMenu.csproj
        const string BuildVersionMetadataPrefix = "+build";
        var attribute = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>();
        if (attribute?.InformationalVersion != null)
        {
            var value = attribute.InformationalVersion;
            var index = value.IndexOf(BuildVersionMetadataPrefix);
            if (index > 0)
            {
                value = value[(index + BuildVersionMetadataPrefix.Length)..];
                if (DateTime.TryParseExact(value, "yyyyMMddHHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var result)) { return result; }
            }
        }
        return default;
    }

    public static async Task UpdateZipBackupAsync(string sourceDbPath, string targetZipFilePath)
    {
        if (string.IsNullOrWhiteSpace(sourceDbPath)) { return; }
        if (string.IsNullOrWhiteSpace(targetZipFilePath)) { return; }
        var dbFileName = Path.GetFileName(sourceDbPath);
        await Task.Run(async () =>  // weil File.Copy und File.Move blockierende I/O-Aufrufe sind
        {
            var targetDir = Path.GetDirectoryName(targetZipFilePath);
            if (!string.IsNullOrEmpty(targetDir))
            {
                if (!Directory.Exists(targetDir)) { Directory.CreateDirectory(targetDir); }
            }
            var tempZipPath = targetZipFilePath + ".tmp";
            var maxRetries = 3;
            var delayMs = 500;
            for (var i = 0; i < maxRetries; i++)
            {
                try
                {

                    if (File.Exists(targetZipFilePath)) { File.Copy(targetZipFilePath, tempZipPath, true); }  // Kopieren des Originals in eine Temp-Datei (falls es schon existiert)
                    var mode = File.Exists(tempZipPath) ? ZipArchiveMode.Update : ZipArchiveMode.Create;
                    using (var fileStream = new FileStream(tempZipPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))  // Temp-Datei öffnen (FileShare.None hält OneDrive fern)
                    {
                        using var archive = new ZipArchive(fileStream, mode);
                        if (mode == ZipArchiveMode.Update)
                        {
                            var existingEntry = archive.GetEntry(dbFileName);
                            existingEntry?.Delete();
                        }
                        var newEntry = archive.CreateEntry(dbFileName, CompressionLevel.Optimal);
                        newEntry.LastWriteTime = File.GetLastWriteTime(sourceDbPath);  // Metadaten setzen: Die echte Modifikationszeit der Datenbank übernehmen
                        using var entryStream = newEntry.Open();
                        using var sourceStream = new FileStream(sourceDbPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        await sourceStream.CopyToAsync(entryStream);  // Asynchrones Kopieren
                    }
                    File.Move(tempZipPath, targetZipFilePath, overwrite: true);  // Atomarer Tausch: Die fertige Temp-Datei ersetzt das Original.
                    break; // Erfolg! Schleife abbrechen.
                }
                catch (IOException) when (i < maxRetries - 1) { await Task.Delay(delayMs); }  // Wenn OneDrive die Datei blockiert, kurz warten und nochmal versuchen
                catch
                {
                    if (File.Exists(tempZipPath))
                    {
                        try { File.Delete(tempZipPath); } catch { }  // Fehler ignorieren, temporäre Datei wird beim nächsten Lauf überschrieben
                    }
                }
            }
        });
    }

    internal static async Task DailyBackupAsync(string filePath, string backupDir)
    {
        try
        {
            // 1. Pfadvorbereitung
            backupDir = Path.Combine(backupDir, new CultureInfo("de-DE").DateTimeFormat.GetDayName(DateTime.Today.DayOfWeek));
            if (!Directory.Exists(backupDir)) { Directory.CreateDirectory(backupDir); }
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var extension = Path.GetExtension(filePath);
            var todaysBackupFile = Path.Combine(backupDir, $"{fileName}_{DateTime.Now:yyyy_MM_dd}{extension}");

            if (File.Exists(todaysBackupFile)) { return; }

            // 2. Sicherer, asynchroner Kopiervorgang (Löst auch das Lock-Problem); FileShare.ReadWrite ist entscheidend für SQLite!
            await using (var sourceStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, useAsync: true))
            {
                await using var destStream = new FileStream(todaysBackupFile, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
                await sourceStream.CopyToAsync(destStream);
            }

            // 3. Rotation (synchron ok, da nur Dateinamen-Operationen)
            var existingBackups = Directory.GetFiles(backupDir, fileName + "*.adb");
            if (existingBackups.Length >= 2)
            {
                var oldestFile = existingBackups.OrderBy(f => new FileInfo(f).CreationTime).First();
                File.Delete(oldestFile);
            }
        }
        catch (Exception ex) { Debug.WriteLine($"Backup fehlgeschlagen: {ex.Message}"); }
    }

    public static string TruncateMiddle(string name, string suffix, Font font, int maxWidth)
    {
        var fullText = $"{name}{suffix}";
        var availableWidth = maxWidth - 4;  // Puffer leicht anpassen, da ListBox weniger fixes internes Padding hat.
        if (TextRenderer.MeasureText(fullText, font).Width <= availableWidth) { return fullText; }
        var leftLen = name.Length / 2;
        var rightLen = name.Length - leftLen;

        while (leftLen + rightLen > 0)
        {
            if (leftLen > rightLen) { leftLen--; }
            else { rightLen--; }
            var testName = name[..leftLen] + "…" + name[^rightLen..];
            var testFull = $"{testName}{suffix}";
            if (TextRenderer.MeasureText(testFull, font).Width <= availableWidth) { return testFull; }
        }
        return $"…{suffix}";
    }

    public static void RestoreWindowBounds(Form form, WindowPlacement? placement, bool isMaximized = false)
    {
        if (isMaximized)
        {
            form.WindowState = FormWindowState.Maximized;
            return;
        }
        if (placement == null) { return; }
        form.StartPosition = FormStartPosition.Manual;
        form.WindowState = FormWindowState.Normal;
        var targetRect = new Rectangle(placement.X, placement.Y, placement.Width, placement.Height);
        var screen = Screen.FromRectangle(targetRect);  // Screen.FromRectangle ist robuster als FromPoint, da es prüft, wo der größte Teil des Fensters liegt.
        var workArea = screen.WorkingArea;
        var width = Math.Max(targetRect.Width, form.MinimumSize.Width);  // nicht größer als Bildschirm, aber nicht kleiner als MinimumSize
        var height = Math.Max(targetRect.Height, form.MinimumSize.Height);
        width = Math.Min(width, workArea.Width);
        height = Math.Min(height, workArea.Height);
        targetRect.Width = width;
        targetRect.Height = height;
        if (targetRect.Right > workArea.Right) { targetRect.X = workArea.Right - targetRect.Width; }
        if (targetRect.Left < workArea.Left) { targetRect.X = workArea.Left; }
        if (targetRect.Bottom > workArea.Bottom) { targetRect.Y = workArea.Bottom - targetRect.Height; }
        if (targetRect.Top < workArea.Top) { targetRect.Y = workArea.Top; }
        form.DesktopBounds = targetRect;
    }

    public static void AdjustComboBoxDropDownWidth(ComboBox cb)
    {
        var maxWidth = cb.Width;
        using var g = cb.CreateGraphics();
        foreach (var item in cb.Items)
        {
            var itemWidth = (int)g.MeasureString(item.ToString(), cb.Font).Width;
            if (itemWidth > maxWidth) { maxWidth = itemWidth; }
        }
        cb.DropDownWidth = maxWidth + SystemInformation.VerticalScrollBarWidth; // Platz für den vertikalen Scrollbalken addieren, um horizontales Scrollen zu vermeiden
    }

    internal static bool RowIsVisible(DataGridView dgv, DataGridViewRow row)
    {
        if (dgv.FirstDisplayedCell == null) { return false; }
        var firstVisibleRowIndex = dgv.FirstDisplayedCell.RowIndex;
        var lastVisibleRowIndex = firstVisibleRowIndex + dgv.DisplayedRowCount(false) - 1;
        return row.Index >= firstVisibleRowIndex && row.Index <= lastVisibleRowIndex;
    }
    internal static void MoveCursorToControl(Control control)
    {
        if (control == null || control.IsDisposed || !control.IsHandleCreated) { return; }
        var clientCenter = new Point(control.Width / 2, control.Height / 2);
        var screenCenter = control.PointToScreen(clientCenter);
        Cursor.Position = screenCenter;
    }

    internal static Bitmap CreateIconFromText(string text, Font font, Color textColor, Size imageSize)
    {
        var bitmap = new Bitmap(imageSize.Width, imageSize.Height);
        using var g = Graphics.FromImage(bitmap);
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        g.SmoothingMode = SmoothingMode.HighQuality;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;

        g.Clear(Color.Transparent);  // Hintergrund transparent halten
        using var brush = new SolidBrush(textColor);
        using var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        var rect = new Rectangle(0, 0, imageSize.Width, imageSize.Height);
        g.DrawString(text, font, brush, rect, format);
        return bitmap;
    }

}