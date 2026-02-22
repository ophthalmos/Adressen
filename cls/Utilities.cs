using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Globalization;
using System.IO.Compression;
using System.Net;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;

namespace Adressen.cls;

internal static class Utils
{
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
        TaskDialog.ShowDialog(hwnd ?? IntPtr.Zero, page);
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

    public static void SortContacts(BindingList<Contact>? contacts)
    {
        if (contacts == null || contacts.Count == 0) { return; }
        var sortedList = contacts.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname).ThenBy(x => x.Unternehmen).ToList();  // ignoriert Groß-/Kleinschreibung
        contacts.Clear();  // BindingList wird geleert, weil sie keine Sortiermethode hat
        foreach (var c in sortedList) { contacts.Add(c); }
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

    public static List<(DateOnly Datum, string Name, int Alter, int Tage, string Id)> CalculateUpcomingBirthdays(IEnumerable<IContactEntity> contacts, int daysLookBack, int daysLookAhead)
    {
        var heute = DateOnly.FromDateTime(DateTime.Today);

        return [.. contacts.Where(x => x.BirthdayDate.HasValue).Select(x => {
                var g = x.BirthdayDate!.Value;

                // Schaltjahr-Korrektur
                var day = (g.Month == 2 && g.Day == 29 && !DateTime.IsLeapYear(heute.Year)) ? 28 : g.Day;

                var gebTagDiesesJahr = new DateOnly(heute.Year, g.Month, day);
                var tage = gebTagDiesesJahr.DayNumber - heute.DayNumber;

                // Jahreswechsel-Logik
                if (tage < -daysLookBack)
                {
                    var dayNext = (g.Month == 2 && g.Day == 29 && !DateTime.IsLeapYear(heute.Year + 1)) ? 28 : g.Day;
                    tage = new DateOnly(heute.Year + 1, g.Month, dayNext).DayNumber - heute.DayNumber;
                }
                else if (tage > daysLookAhead)
                {
                    var dayPrev = (g.Month == 2 && g.Day == 29 && !DateTime.IsLeapYear(heute.Year - 1)) ? 28 : g.Day;
                    var tageLetztesJahr = new DateOnly(heute.Year - 1, g.Month, dayPrev).DayNumber - heute.DayNumber;
                    if (tageLetztesJahr >= -daysLookBack) { tage = tageLetztesJahr; } }

                return new { Entity = x, Tage = tage, OriginalGeb = g };
            })
            .Where(x => x.Tage >= -daysLookBack && x.Tage <= daysLookAhead).OrderBy(x => x.Tage).Select(x =>            {
                var alter = heute.Year - x.OriginalGeb.Year;
                if (x.Tage > 0) { alter--; } return (Datum: x.OriginalGeb, Name: x.Entity.DisplayName, Alter: alter, x.Tage, Id: x.Entity.UniqueId);
            })];
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
            if (Uri.IsWellFormedUriString(url, UriKind.Absolute))
            {
                ProcessStartInfo psi = new(url) { UseShellExecute = true };
                Process.Start(psi);
            }
            else { MsgTaskDlg(handle, "Ungültiger Link!", "'" + url + "' ist keine gültige URL.", TaskDialogIcon.ShieldWarningYellowBar); }
        }
        catch (Exception ex) when (ex is Win32Exception || ex is InvalidOperationException) { ErrTaskDlg(handle, ex); }
    }

    internal static bool GoogleConnectionCheck(nint hwnd, string path)
    {
        if (new Ping().Send(new IPAddress([8, 8, 8, 8]), 1000).Status != IPStatus.Success)
        {
            MsgTaskDlg(hwnd, "Keine Internetverbindung!", "Überprüfen Sie das Netzwerk.", TaskDialogIcon.ShieldWarningYellowBar);
            return false;
        }
        else if (!File.Exists(path))
        {
            MsgTaskDlg(hwnd, "Der Key-File wurde nicht gefunden!", "'" + path + "' fehlt.", TaskDialogIcon.ShieldWarningYellowBar);
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

                if (sOld != sNew)
                {
                    var displayOld = string.IsNullOrEmpty(sOld) ? "" : sOld;
                    var displayNew = string.IsNullOrEmpty(sNew) ? "∅" : sNew;
                    sb.AppendLine($"{fieldName}: {displayOld} ➔ {displayNew}");
                }
            }
            else // z.B. Datum (Geburtstag) oder Zahlen
            {
                if (!Equals(valOld, valNew))
                {
                    // Formatierung für Nicht-Strings (übernimmt Ihre Logik für [Leer])
                    static string FormatObj(object? o)
                    {
                        if (o == null) { return "[Leer]"; }
                        if (o is DateTime d) { return d.ToShortDateString(); }
                        if (o is DateOnly dO) { return dO.ToString(); }  // Falls .NET 10 DateOnly nutzt
                        return o.ToString() ?? "";
                    }
                    sb.AppendLine($"{fieldName}: {FormatObj(valOld)} ➔ {FormatObj(valNew)}");
                }
            }
        }
        return sb.ToString();
    }

    internal static void StartSearchCacheWarmup(IEnumerable<IContactEntity> items) => Task.Run(() => { foreach (var item in items) { var warmup = item.SearchText; } });

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
        //using TaskDialogIcon questionDialogIcon = new(Properties.Resources.question32);
        var page = new TaskDialogPage
        {
            Caption = Application.ProductName,
            Heading = "Wählen Sie die Textverarbeitung",
            Icon = TaskDialogIcon.ShieldBlueBar,
            Buttons = { wordButton, libreButton, TaskDialogButton.Cancel },
            AllowCancel = true,
            SizeToContent = true
        };
        var result = TaskDialog.ShowDialog(hwnd, page);
        if (result == wordButton) { return true; }
        if (result == libreButton) { return false; }
        return null;
    }

    internal static (bool IsYes, bool IsNo, bool IsCancelled) YesNo_TaskDialog(IWin32Window? owner, string caption, string heading, string text, string yes = "", string no = "", bool defBtn = true)
    {
        var yesButton = string.IsNullOrEmpty(yes) ? TaskDialogButton.Yes : new TaskDialogButton(yes);
        var noButton = string.IsNullOrEmpty(no) ? TaskDialogButton.No : new TaskDialogButton(no);
        var page = new TaskDialogPage
        {
            Caption = caption,
            Heading = heading,
            Text = text,
            Icon = new TaskDialogIcon(Properties.Resources.question32),
            Buttons = { yesButton, noButton },
            DefaultButton = defBtn ? yesButton : noButton,
            AllowCancel = true,
            SizeToContent = true
        };
        var result = owner is not null ? TaskDialog.ShowDialog(owner, page) : TaskDialog.ShowDialog(page);
        var isYes = result == yesButton;
        var isNo = result == noButton;
        var isCancelled = result == TaskDialogButton.Cancel || (!isYes && !isNo);
        return (isYes, isNo, isCancelled);
    }


    internal static bool ValuesEqual(object? a, object? b)
    {
        if (a is DBNull) { a = string.Empty; }
        if (b is DBNull) { b = string.Empty; }
        if (a is string sa && b is string sb) { return string.Equals(sa, sb, StringComparison.Ordinal); }
        return string.Equals(a?.ToString(), b?.ToString(), StringComparison.Ordinal); // Fallback: ToString-Vergleich
    }

    public static string BuildMask(params string[] fields) => string.Join(",", fields.Where(f => !string.IsNullOrWhiteSpace(f)).Select(f => f.Trim()));

    internal static (bool askBefore, bool deleteNow) AskBeforeDeleteContact(nint handle, IContactEntity contact, bool askBeforeDelete, bool showVerification = true)
    {
        var deleteNow = false;
        try
        {
            // Wir holen die Daten direkt vom Objekt, nicht aus den Grid-Zellen
            // Das ist viel schneller und weniger fehleranfällig
            var details = contact.DisplayName;

            // Falls du mehr Details wie Unternehmen/Ort anzeigen willst, 
            // musst du diese ggf. in das Interface aufnehmen oder hier casten:
            var zusatzInfo = "";
            if (contact is Contact c)
            {
                zusatzInfo = $"\n{c.Unternehmen}\n{c.Strasse}\n{c.PLZ} {c.Ort}";
            }
            else if (contact is Adresse a)
            {
                zusatzInfo = $"\n{a.Unternehmen}\n{a.Strasse}\n{a.PLZ} {a.Ort}";
            }

            using TaskDialogIcon questionDialogIcon = new(Properties.Resources.question32);
            var page = new TaskDialogPage()
            {
                Heading = "Möchten Sie den Datensatz löschen?",
                Text = (details + zusatzInfo).Trim(),
                Caption = Application.ProductName,
                Icon = questionDialogIcon,
                AllowCancel = true,
                SizeToContent = true,
                Verification = showVerification ? new TaskDialogVerificationCheckBox() { Text = "Diese Frage immer anzeigen" } : null,
                Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
            };

            if (page.Verification is TaskDialogVerificationCheckBox check)
            {
                check.Checked = askBeforeDelete;
            }

            var resultButton = TaskDialog.ShowDialog(handle, page);


            // Logik für die Checkbox
            if (page.Verification is TaskDialogVerificationCheckBox finalCheck)
            {
                if (askBeforeDelete && !finalCheck.Checked)
                {
                    MsgTaskDlg(page.BoundDialog?.Handle ?? IntPtr.Zero, "Hinweis", "Sie können die Sicherheitsabfrage in\nden Einstellungen wieder einschalten.", new(Properties.Resources.info32));
                    askBeforeDelete = false;
                }
                else if (finalCheck.Checked)
                {
                    askBeforeDelete = true;
                }
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
            using TaskDialogIcon questionDialogIcon = new(Properties.Resources.question32);
            var page = new TaskDialogPage()
            {
                Heading = "Möchten Sie den Datensatz löschen?",
                Text = $"{vorname} {nachname}\n{unternehmen}\n{strasse}\n{plz} {ort}".Trim(),
                Caption = Application.ProductName,
                Icon = questionDialogIcon,
                AllowCancel = true,
                SizeToContent = true,
                Verification = showVerification ? new TaskDialogVerificationCheckBox() { Text = "Diese Frage immer anzeigen" } : "",
                Buttons = { TaskDialogButton.Yes, TaskDialogButton.No },
            };
            page.Verification.Checked = askBeforeDelete;
            var resultButton = TaskDialog.ShowDialog(hwnd, page);

            if (askBeforeDelete && !page.Verification.Checked)
            {
                MsgTaskDlg(hwnd, "Hinweis", "Sie können die Sicherheitsabfrage in\nden Einstellungen wieder einschalten.", new(Properties.Resources.info32));
                askBeforeDelete = false;
            }
            else { askBeforeDelete = true; }
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

    internal static string CorrectUNC(string unc) => unc.StartsWith('\\') ? @"\\" + unc.TrimStart('\\') : unc;

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

    public static IEnumerable<string> ReadAsLines(string filename)
    {
        using var reader = new StreamReader(filename);
        while (!reader.EndOfStream) { yield return reader.ReadLine()!; }
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

            if (!Directory.Exists(backupDir))
            {
                Directory.CreateDirectory(backupDir);
            }

            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var extension = Path.GetExtension(filePath);
            var todaysBackupFile = Path.Combine(backupDir, $"{fileName}_{DateTime.Now:yyyy_MM_dd}{extension}");

            if (File.Exists(todaysBackupFile)) { return; }

            // 2. Sicherer, asynchroner Kopiervorgang (Löst auch das Lock-Problem)
            // FileShare.ReadWrite ist entscheidend für SQLite!
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
        catch (Exception ex)
        {
            Debug.WriteLine($"Backup fehlgeschlagen: {ex.Message}");
        }
    }
}

public interface IContactEntity
{
    string UniqueId
    {
        get;
    }
    string DisplayName
    {
        get;
    }
    string SearchText
    {
        get;
    }
    DateOnly? BirthdayDate
    {
        get;
    }
    IList<string> GroupList
    {
        get;
    }
    Task<Image?> GetPhotoAsync();
    void ResetSearchCache();
}
