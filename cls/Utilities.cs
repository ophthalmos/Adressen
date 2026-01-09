using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Globalization;
using System.Net;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
//using Google.Apis.Util.Store;
using Microsoft.Win32;
using Timer = System.Windows.Forms.Timer;
using Word = Microsoft.Office.Interop.Word;

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

    internal static string GenerateDetailedDiff(Contact current, Contact old)
    {
        var sb = new StringBuilder();
        void Check(string label, string? oldVal, string? newVal)
        {
            var o = oldVal ?? string.Empty; // Normalisieren um null/empty gleich zu behandeln
            var n = newVal ?? string.Empty;
            if (o != n)
            {
                var displayOld = string.IsNullOrEmpty(o) ? "" : o; // Formatierung: Leere Werte sichtbar machen
                var displayNew = string.IsNullOrEmpty(n) ? "∅" : n;
                sb.AppendLine($"{label}: {displayOld} ➔ {displayNew}");
            }
        }
        Check("Anrede", old.Anrede, current.Anrede);
        Check("Präfix", old.Praefix, current.Praefix);
        Check("Nachname", old.Nachname, current.Nachname);
        Check("Vorname", old.Vorname, current.Vorname);
        Check("Zwischenname", old.Zwischenname, current.Zwischenname);
        Check("Nickname", old.Nickname, current.Nickname);
        Check("Suffix", old.Suffix, current.Suffix);
        Check("Firma", old.Firma, current.Firma);
        Check("Straße", old.Strasse, current.Strasse);
        Check("PLZ", old.PLZ, current.PLZ);
        Check("Ort", old.Ort, current.Ort);
        Check("Land", old.Land, current.Land);
        Check("Betreff", old.Betreff, current.Betreff);
        Check("Grussformel", old.Grussformel, current.Grussformel);
        Check("Schlussformel", old.Schlussformel, current.Schlussformel);
        Check("E-Mail 1", old.Mail1, current.Mail1);
        Check("Mail 2", old.Mail2, current.Mail2);
        Check("Telefon 1", old.Telefon1, current.Telefon1);
        Check("Telefon 2", old.Telefon2, current.Telefon2);
        Check("Mobil", old.Mobil, current.Mobil);
        Check("Fax", old.Fax, current.Fax);
        Check("Internet", old.Internet, current.Internet);
        Check("Notizen", old.Notizen, current.Notizen);
        if (old.Geburtstag != current.Geburtstag)
        {
            var oldDate = old.Geburtstag.HasValue ? old.Geburtstag.Value.ToShortDateString() : "[Leer]";
            var newDate = current.Geburtstag.HasValue ? current.Geburtstag.Value.ToShortDateString() : "[Leer]";
            sb.AppendLine($"Geburtstag: {oldDate} ➔ {newDate}");
        }
        return sb.ToString();
    }

    internal static void StartSearchCacheWarmup(IEnumerable<IContactEntity> items) => Task.Run(() => { foreach (var item in items) { var warmup = item.SearchText; } });

    internal static async Task<PeopleServiceService> GetPeopleServiceAsync(string secretPath, string tokenDir)
    {
        string[] scopes = [PeopleServiceService.Scope.Contacts]; // für OAuth2-Freigabe, mehrere Eingaben mit Komma gerennt (PeopleServiceService.Scope.ContactsOtherReadonly)
        UserCredential credential;
        using (FileStream stream = new(secretPath, FileMode.Open, FileAccess.Read))
        {
            credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets, scopes, "user", CancellationToken.None, new Google.Apis.Util.Store.FileDataStore(tokenDir, true));
        }
        return new PeopleServiceService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = Application.ProductName,
        });
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

    //internal static string CleanTemporaryWordPrefix(string fullPath)
    //{
    //    if (string.IsNullOrEmpty(fullPath)) { return fullPath; }
    //    var directory = Path.GetDirectoryName(fullPath);
    //    var fileName = Path.GetFileName(fullPath);
    //    if (fileName.StartsWith("~$")) { fileName = fileName[2..]; }
    //    return Path.Combine(directory ?? string.Empty, fileName);
    //}

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


    internal static void WordInfoTaskDlg(nint hwnd, string[] allKeys, TaskDialogIcon icon)
    {
        var btnCreateDoc = new TaskDialogButton("Beispieldokument erstellen");
        var page = new TaskDialogPage()
        {
            Caption = Application.ProductName,
            Heading = "Folgende Textmarken können in einem Word-Dokument verwendet werden:",
            Text = string.Join(", ", allKeys),
            Icon = icon,
            Footnote = "Tipp: Erstellen Sie eigene Vorlagen mit passenden Textmarken.",
            AllowCancel = true,
            Buttons = { btnCreateDoc, TaskDialogButton.Close }
        };
        //btnCreateDoc.Click += (s, e) => { CreateTextMakerDocument(allKeys, hwnd); };
        btnCreateDoc.Click += (s, e) => { CreateWordDocument(allKeys, hwnd); };
        TaskDialog.ShowDialog(hwnd, page);
    }

    //internal static void HotkeysTaskDlg(IntPtr hwnd)
    //{
    //    var sb = new StringBuilder();
    //    sb.AppendLine("Strg+Enter:");
    //    sb.AppendLine("Shift+Tab:");
    //    sb.AppendLine("F5/F6:             Wechsel zwischen Adressen und Kontakte");
    //    TaskDialog.ShowDialog(hwnd, new TaskDialogPage()
    //    {
    //        SizeToContent = true,
    //        Caption = Application.ProductName,
    //        Heading = "Hilfreiche Tastenkombinationen",
    //        Text = sb.ToString(),
    //        Icon = TaskDialogIcon.ShieldSuccessGreenBar,
    //        AllowCancel = true,
    //        Buttons = { TaskDialogButton.Close }
    //    });
    //}

    internal static void HelpMsgTaskDlg(nint hwnd, string appName, Icon? icon)
    {
        var curVersion = Assembly.GetExecutingAssembly().GetName().Version;
        var threeVersion = curVersion?.ToString(3) ?? "unbekannt"; //curVersion is not null ? $"{curVersion.Major}.{curVersion.Minor}.{curVersion.Build}" : "unbekannt";
        var buildDate = GetBuildDate();
        TaskDialogButton paypalButton = new TaskDialogCommandLinkButton("Anerkennung spenden via PayPal");
        TaskDialogButton updateButton = new TaskDialogCommandLinkButton("Nach Programm-Update suchen…") { AllowCloseDialog = false };
        var foot = "              © " + buildDate.ToString("yyyy") + " Wilhelm Happe, Version " + threeVersion + " (" + buildDate.ToString("d") + ")";
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
            Buttons = { paypalButton, updateButton, TaskDialogButton.OK },
            DefaultButton = TaskDialogButton.OK,
            Footnote = foot
        };

        TaskDialogButton downloadButton = new TaskDialogCommandLinkButton("AdressenSetup.exe herunterladen", "AdressenSetup.exe wird im Download-Ordner\ngespeichert. Führen Sie das Setupprogramm\naus, um die neueste Version zu installieren.");
        var updatePage = new TaskDialogPage()
        {
            Caption = appName,
            Heading = $"{appName} ist auf dem neuesten Stand.",
            Text = $"Version {threeVersion} (Offizielles Build, 64-Bit)", //\n\nAutomatische Suche nach Updates:",
            Icon = TaskDialogIcon.Information,
            AllowCancel = true,
            SizeToContent = true,
            Buttons = { TaskDialogButton.Close }
        };

        //var radioButton1 = updatePage.RadioButtons.Add("täglich");
        //var radioButton2 = updatePage.RadioButtons.Add("wöchentlich");
        //var radioButton3 = updatePage.RadioButtons.Add("monatlich");
        //var radioButton4 = updatePage.RadioButtons.Add("niemals");
        //radioButton4.Checked = true;

        //radioButton1.CheckedChanged += (s, e) => Console.WriteLine("RadioButton1 CheckedChanged: " + radioButton1.Checked);
        //radioButton2.CheckedChanged += (s, e) => Console.WriteLine("RadioButton2 CheckedChanged: " + radioButton2.Checked);
        //radioButton3.CheckedChanged += (s, e) => Console.WriteLine("RadioButton3 CheckedChanged: " + radioButton3.Checked);
        //radioButton4.CheckedChanged += (s, e) => Console.WriteLine("RadioButton4 CheckedChanged: " + radioButton4.Checked);


        var urlString = string.Empty;
        updateButton.Click += async (sender, e) =>
        {
            updateButton.Enabled = false; // um doppelte Klicks zu verhindern
            var xmlURL = "https://www.netradio.info/download/adressen.xml";
            Version? updateVersion = null;
            var dateString = string.Empty;
            try
            {
                using var requestMessage = new HttpRequestMessage(HttpMethod.Get, xmlURL);
                requestMessage.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/xml"));
                using var response = await HttpService.Client.SendAsync(requestMessage);
                response.EnsureSuccessStatusCode(); // Wirft eine Exception bei Fehlern wie 404 oder 500
                var xmlContent = await response.Content.ReadAsStringAsync();
                var doc = XDocument.Parse(xmlContent);
                var versionString = doc.Element("adressen")?.Element("version")?.Value;
                if (versionString != null) { updateVersion = new Version(versionString); }
                dateString = doc.Element("adressen")?.Element("date")?.Value;
                urlString = doc.Element("adressen")?.Element("url64")?.Value;
            }
            catch (HttpRequestException ex) // when (ex is WebException or NullReferenceException or ArgumentNullException or XmlException or ArgumentException or IOException)
            {
                updatePage.Icon = TaskDialogIcon.Error;
                updatePage.Heading = "Es ist ein Fehler aufgetreten.";
                var exStatusCode = ex.StatusCode;
                if (exStatusCode == HttpStatusCode.NotFound) { updatePage.Text = "Die Update-Informationen wurden nicht gefunden."; }
                else { updatePage.Text = exStatusCode?.ToString().Length > 0 ? $"Status-Code: {exStatusCode}\nFehlermeldung: {ex.Message}" : $"Fehlermeldung: {ex.Message}"; } // + "\n\nAutomatische Suche nach Updates:"; }
            }
            catch (Exception ex) // when (ex is WebException or NullReferenceException or ArgumentNullException or XmlException or ArgumentException or IOException)
            {
                updatePage.Icon = TaskDialogIcon.Error;
                updatePage.Heading = ex.GetType().ToString();
                updatePage.Text = ex.Message;  // + "\n\nAutomatische Suche nach Updates:"; 
            }
            if (updateVersion != null && updateVersion.CompareTo(curVersion) > 0)
            {
                updatePage.Heading = "Es steht ein Update zur Verfügung!";
                updatePage.Text = $"Version {updateVersion?.ToString()} vom {dateString}";  //\n\nAutomatische Suche nach Updates:";
                updatePage.Buttons.Add(downloadButton);
            }
            initialPage.Navigate(updatePage);  // When the user clicks updateButton, navigate to the second page.
        };

        var result = TaskDialog.ShowDialog(hwnd, initialPage);
        if (result == paypalButton) { StartLink(hwnd, "https://www.paypal.com/donate/?hosted_button_id=3HRQZCUW37BQ6"); }
        else if (result == downloadButton) { StartLink(hwnd, urlString); }
    }


    //public static bool IsLibreOfficeInstalled()
    //{
    //    foreach (var root in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })  // Sowohl HKCU als auch HKLM prüfen
    //    {
    //        using var key = RegistryKey.OpenBaseKey(root, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath");
    //        if (key != null && key.ValueCount > 0)
    //        {
    //            var path = key.GetValue(key.GetValueNames()[0]) as string;
    //            if (!string.IsNullOrEmpty(path)) { return true; }
    //        }
    //    }
    //    return false;
    //}

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

    //internal static IEnumerable<string> DeserializeGroups(nint hwnd, string json, JsonSerializerOptions options)
    //{
    //    try { return JsonSerializer.Deserialize<List<string>>(json, options) ?? Enumerable.Empty<string>(); }
    //    catch (JsonException ex)
    //    {
    //        ErrTaskDlg(hwnd, ex);
    //        return [];
    //    }
    //}

    public static string BuildMask(params string[] fields) => string.Join(",", fields.Where(f => !string.IsNullOrWhiteSpace(f)).Select(f => f.Trim()));

    internal static (bool askBefore, bool deleteNow) AskBeforeDeleteContact(nint handle, IContactEntity contact, bool askBeforeDelete, bool showVerification = true)
    {
        var deleteNow = false;
        try
        {
            // Wir holen die Daten direkt vom Objekt, nicht aus den Grid-Zellen
            // Das ist viel schneller und weniger fehleranfällig
            var details = contact.DisplayName;

            // Falls du mehr Details wie Firma/Ort anzeigen willst, 
            // musst du diese ggf. in das Interface aufnehmen oder hier casten:
            var zusatzInfo = "";
            if (contact is Contact c)
            {
                zusatzInfo = $"\n{c.Firma}\n{c.Strasse}\n{c.PLZ} {c.Ort}".Trim();
            }
            else if (contact is Adresse a)
            {
                zusatzInfo = $"\n{a.Firma}\n{a.Strasse}\n{a.PLZ} {a.Ort}".Trim();
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
            var firma = adresse.Firma ?? string.Empty;
            var strasse = adresse.Strasse ?? string.Empty;
            var plz = adresse.PLZ ?? string.Empty;
            var ort = adresse.Ort ?? string.Empty;
            using TaskDialogIcon questionDialogIcon = new(Properties.Resources.question32);
            var page = new TaskDialogPage()
            {
                Heading = "Möchten Sie den Datensatz löschen?",
                Text = $"{vorname} {nachname}\n{firma}\n{strasse}\n{plz} {ort}".Trim(),
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

    //internal static T GetOrAddKey<T>(IList<T> list, Func<T, bool> predicate, Func<T> factory)
    //{
    //    var item = list.FirstOrDefault(predicate);
    //    if (item == null)
    //    {
    //        item = factory();
    //        list.Add(item);
    //    }
    //    return item;
    //}

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

    private static void CreateWordDocument(string[] allKeys, nint handle)
    {
        if (Type.GetTypeFromProgID("Word.Application") == null)
        {
            MsgTaskDlg(handle, "Microsoft Word is not installed", "Installieren Sie Microsoft Word.");
            return;
        }
        Word.Document? wordDoc = null;
        dynamic? wordApp = null;
        try
        {
            wordApp = new Word.Application { Visible = true };
            wordDoc = wordApp.Documents.Add();
            dynamic properties = wordDoc.BuiltInDocumentProperties; // Indizierung erst zur Laufzeit zu prüfen
            properties[Word.WdBuiltInProperty.wdPropertyAuthor].Value = "Wilhelm Happe";
            properties[Word.WdBuiltInProperty.wdPropertyTitle].Value = "Adressen-Vorlage";
            properties[Word.WdBuiltInProperty.wdPropertySubject].Value = "Nur als Beispiel gedacht";
            properties[Word.WdBuiltInProperty.wdPropertyKeywords].Value = "Adressen, Briefvorlage";
            properties[Word.WdBuiltInProperty.wdPropertyComments].Value = "Die Datei wurde gespeichert...";
            var para0 = wordDoc.Paragraphs.Add();
            para0.Range.Font.Size = 14;
            para0.Range.Text = "Präfix_Vorname_Zwischenname_Nachname";
            wordDoc.Bookmarks.Add("Präfix_Vorname_Zwischenname_Nachname", para0.Range);
            para0.Format.SpaceAfter = 0f;
            para0.Range.InsertParagraphAfter();

            var para1 = wordDoc.Paragraphs.Add();
            para1.Range.Font.Size = 14;
            para1.Range.Text = "StraßeNr";
            para1.Range.Bookmarks.Add("StraßeNr", para1.Range);
            para1.Format.SpaceAfter = 6f;
            para1.Range.InsertParagraphAfter();

            var para2 = wordDoc.Paragraphs.Add();
            para2.Range.Font.Size = 14;
            para2.Range.Text = "PLZ_Ort";
            para2.Range.Bookmarks.Add("PLZ_Ort", para2.Range);
            para2.Format.SpaceAfter = 12f;
            para2.Range.InsertParagraphAfter();

            //var para3 = wordDoc.Paragraphs.Add();
            //para3.Format.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;

            var para4 = wordDoc.Paragraphs.Add();
            para4.Range.Text = "Probieren Sie nun das Einfügen einer Adresse aus, indem Sie im Adressen-Programm eine Adresse selektieren und dann auf den Button »In Brief einfügen« klicken. Wiederholen Sie den Vorgang mit anderen Adressen!";
            para4.Range.InsertParagraphAfter();

            var para5 = wordDoc.Paragraphs.Add();
            para5.Range.Text = "Wenn Sie eine Textmarke hinzufügen möchten, markieren Sie zuerst die Stelle der Textmarke in Ihrem Dokument. Wählen Sie die Registerkarte »Einfügen« und dann »Textmarke« aus. Schneller geht es, wenn Sie die Tastenkombination Strg+Shift+F5 drücken.";
            para5.Range.InsertParagraphAfter();

            var para6 = wordDoc.Paragraphs.Add();
            para6.Range.Text = "Um Textmarken-Klammen anzuzeigen, führen Sie die folgenden Schritte aus:\x0BKlicken Sie auf Datei > Optionen > Erweitert.\x0BWählen Sie unter \"Dokumentinhalt anzeigen\" die Option \"Textmarken anzeigen\".";
            para6.Range.InsertParagraphAfter();

            var para7 = wordDoc.Paragraphs.Add();
            para7.Range.Font.Bold = 1;
            para7.Range.Text = "Liste der möglichen Textmarkierungen:";
            para7.Format.SpaceAfter = 0f;
            para7.Range.InsertParagraphAfter();

            var para8 = wordDoc.Paragraphs.Add();
            para8.Range.Font.Name = "Courier New";
            para8.Range.NoProofing = 1;
            para8.Range.Text = string.Join(Environment.NewLine, allKeys);
            para8.Format.SpaceAfter = 6f;
            para8.Range.InsertParagraphAfter();

            var para9 = wordDoc.Paragraphs.Add();
            para9.Range.Text = "Die kombinierten Textmarken (mit Unterstrich) dienen dazu, doppelte Leerzeichen zu vermeiden.";
            //para9.Range.InsertParagraphAfter();

            var downloadPath = NativeMethods.SHGetKnownFolderPath(new Guid("374DE290-123F-4565-9164-39C4925E467B"), 0);
            wordDoc.SaveAs2(downloadPath + @"\Adressen-Vorlage.dotx", Word.WdSaveFormat.wdFormatXMLTemplate);
            wordApp.Activate();
            wordApp.Dialogs[Word.WdWordDialog.wdDialogFileSummaryInfo].Show();
        }
        catch (Exception ex) { ErrTaskDlg(handle, ex); } //  + Environment.NewLine + ex.StackTrace

        finally
        {
            if (wordDoc != null)
            {
                try { wordDoc.Close(false); } catch { }  // Ignorieren falls Dokument schon geschlossen
                Marshal.ReleaseComObject(wordDoc);
            }
            if (wordApp != null)
            {
                try { wordApp.Quit(); } catch { } // Ignorieren falls Word App schon geschlossen 
                Marshal.ReleaseComObject(wordApp);
            }
            GC.Collect(); // den Garbage Collector zwingen, die COM-Wrapper sofort aufzuräumen
            GC.WaitForPendingFinalizers();
            //GC.Collect(); // COM-Objekte benötigen oft zwei Durchläufe des Garbage Collectors, um die endgültige Freigabe (Finalization) der COM-Proxys zu gewährleisten.
            //GC.WaitForPendingFinalizers();
        }

        //finally
        //{
        //    if (wordDoc != null)
        //    {
        //        Marshal.ReleaseComObject(wordDoc);
        //        wordDoc = null;
        //    }
        //    if (wordApp != null)
        //    {
        //        Marshal.ReleaseComObject(wordApp);
        //        wordApp = null;
        //    }
        //    GC.Collect();
        //}
    }

    /*    private static void CreateTextMakerDocument(string[] allKeys, IntPtr handle)
        {
            var textMakerType = Type.GetTypeFromProgID("TextMaker.Application");
            if (textMakerType == null)
            {
                ErrorMsgTaskDlg(handle, "TextMaker is not installed", "Installieren Sie SoftMaker Office.");
                return;
            }
            dynamic? textDoc = null;
            dynamic? textApp = null;

            try
            {
                textApp = Activator.CreateInstance(textMakerType);
                if (textApp == null)
                {
                    ErrorMsgTaskDlg(handle, "TextMaker could not be started", "Stellen Sie sicher, dass TextMaker korrekt installiert ist.");
                    return;
                }
                textApp.WindowState = TmWindowState.tmWindowStateMaximize; // textApp[SmoWindowState.smoWindowStateMaximize]; // = true; // Maximieren des Fensters   

                textApp.Visible = true;
                textDoc = textApp.Documents.Add();
                textDoc.BuiltInDocumentProperties[SmoBuiltInProperty.smoPropertyAuthor].Value = "Wilhelm Happe"; // textApp.ActiveDocument.BuiltInDocumentProperties
                textDoc.BuiltInDocumentProperties[SmoBuiltInProperty.smoPropertyTitle].Value = "Adressen-Vorlage";
                textDoc.BuiltInDocumentProperties[SmoBuiltInProperty.smoPropertySubject].Value = "Nur als Beispiel gedacht";
                textDoc.BuiltInDocumentProperties[SmoBuiltInProperty.smoPropertyKeywords].Value = "Adressen, Briefvorlage";
                textDoc.BuiltInDocumentProperties[SmoBuiltInProperty.smoPropertyComments].Value = "Die Datei wurde in Ihrem Download-Ordner gespeichert.\nSie kann gelöschte werden, wenn Sie sie nicht benötigen.";

                textApp.ActiveWindow.View.ShowBookmarks = true; 
                textApp.Application.Options.EnableSound = true; // Sound aktivieren 
                textDoc.PageSetup.TopMargin = textApp.Application.MillimetersToPoints(25); // oberen Rand auf n Millimeter setzen
                textDoc.Paragraphs(1).PreferredLineSpacing = 150; // Zeilenabstand auf "Automatisch 150%" setzen
                textDoc.Selection.Font.Name = "Courier New";
                textDoc.Selection.Font.Size = 14;
                textDoc.Selection.TypeText("Programmieren mit BasicMaker"); //  An der aktuellen Schreibmarke Text einfügen
                textDoc.Selection.TypeParagraph();

                //textDoc.Selection.TypeText("[Präfix_Vorname_Zwischenname_Nachname]");
                //textDoc.Selection.TypeParagraph();

                //foreach (var text in allKeys)
                //{
                //    textDoc.Selection.Font.Name = "Calibri";
                //    textDoc.Selection.Font.Size = 12;
                //    textDoc.Selection.TypeText(text); // string.Join(Environment.NewLine, allKeys);
                //    textDoc.Selection.TypeParagraph(); //  Wagenrücklauf an der aktuellen Schreibmarke einfügen
                //}


                var downloadPath = NativeMethods.SHGetKnownFolderPath(new Guid("374DE290-123F-4565-9164-39C4925E467B"), 0);
                textDoc.SaveAs(downloadPath + @"\Adressen-Vorlage.tmdx", TmSaveFormat.tmFormatDocument);
                textApp.Activate();
                //textApp.Application.Dialogs[smoDialogFileSummaryInfo].Show();  // funktioniert nicht  
            }
            catch (Exception ex) { ErrorMsgTaskDlg(handle, ex.GetType().ToString(), ex.Message); }
            finally
            {
                if (textDoc != null) { Marshal.ReleaseComObject(textDoc); }
                if (textApp != null) { Marshal.ReleaseComObject(textApp); }
                GC.Collect();
            }
        } */

    internal static bool IsInnoSetupValid(string assemblyLocation)
    {
        var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Adressen_is1");
        if (key == null) { return false; }
        var value = (string?)key.GetValue("UninstallString");
        if (value == null) { return false; }
        else if (Debugger.IsAttached) { return true; } // run by Visual Studio
        else { return assemblyLocation.Equals(RemoveFromEnd(value.Trim('"'), "\\unins000.exe"), StringComparison.Ordinal); } // "C:\Program Files\ClipMenu\unins000.exe"
    }

    private static string RemoveFromEnd(string str, string toRemove) => str.EndsWith(toRemove) ? str[..^toRemove.Length] : str;

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

    internal static bool RowIsVisible(DataGridView dgv, DataGridViewRow row)
    {
        if (dgv.FirstDisplayedCell == null) { return false; }
        var firstVisibleRowIndex = dgv.FirstDisplayedCell.RowIndex;
        var lastVisibleRowIndex = firstVisibleRowIndex + dgv.DisplayedRowCount(false) - 1;
        return row.Index >= firstVisibleRowIndex && row.Index <= lastVisibleRowIndex;
    }

    internal static int GetFirstVisibleRowIndex(DataGridView dgv)
    {
        var firstVisibleIndex = -1;
        foreach (DataGridViewRow row in dgv.Rows)
        {
            if (row.Visible && row.Displayed)
            {
                firstVisibleIndex = row.Index;
                break;
            }
        }
        return firstVisibleIndex;
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

    internal static void ApplyColumnSettings(DataGridView dgv, int[] widths, bool[] hideStatus)
    {
        if (dgv.Columns.Count == 0) { return; }

        for (var i = 0; i < dgv.Columns.Count; i++)
        {
            // 1. Sichtbarkeit (Array hideColumnStd)
            if (hideStatus != null && i < hideStatus.Length)
            {
                dgv.Columns[i].Visible = !hideStatus[i];
            }

            // 2. Breite (Array columnWidths)
            if (widths != null && i < widths.Length)
            {
                dgv.Columns[i].Width = Math.Max(20, widths[i]); // Min 20px
            }
            else
            {
                dgv.Columns[i].Width = dgv.Columns[i].Name == "Nachname" ? 200 : 100;
            }

            // 3. Relationen IMMER ausblenden (Sicherheitsnetz)
            if (dgv.Columns[i].Name is "Gruppen" or "Dokumente" or "Foto" or "SearchText" or "GroupList") // or "UniqueId" or "BirthdayDate" 
            {
                dgv.Columns[i].Visible = false;
            }
        }
    }

    //internal static void SetColumnWidths(int[] columnWidths, DataGridView dgv)
    //{
    //    var widths = columnWidths ?? [];
    //    if (widths.Length == 0) // noch keine Einstellungen vorhanden
    //    {
    //        for (var i = 0; i < dgv.Columns.Count; i++)
    //        {
    //            if (dgv.Columns[i].Name == "Nachname") { dgv.Columns[i].Width = 200; }
    //            else { dgv.Columns[i].Width = 100; }
    //        }
    //    }
    //    else
    //    {
    //        for (var i = 0; i < widths.Length && i < dgv.Columns.Count; i++) { dgv.Columns[i].Width = widths[i]; }
    //    }
    //}

    internal static string GetGooglePhoneByType(Person person, string type)
    {
        foreach (var phone in person.PhoneNumbers ?? []) // falls PhoneNumbers null ist, wird die Schleife dank ?? [] einfach übersprungen
        {
            if (phone.Type?.Contains(type, StringComparison.OrdinalIgnoreCase) == true) { return phone.Value ?? string.Empty; }
        }
        return string.Empty;
    }


    //public static bool[]? FromBase64String(string base64String)
    //{
    //    try
    //    {
    //        var bytes = Convert.FromBase64String(base64String); // 1. Decodierung des Base64-Strings
    //        var boolArray = new bool[(bytes.Length * 8)]; // 8 Bits pro Byte // 2. Ermitteln der Länge des Bool-Arrays
    //        for (var i = 0; i < bytes.Length; i++)        // 3. Umwandlung in ein Bool-Array
    //        {
    //            for (var j = 0; j < 8; j++)
    //            {
    //                var bit = bytes[i] >> j & 1;    // Extrahiert das j-te Bit
    //                boolArray[i * 8 + j] = bit == 1;  // Wandelt Bit in Bool um
    //            }
    //        }
    //        return boolArray;
    //    }
    //    catch (FormatException) { return null; }
    //}

    //public static string BoolArray2Base64String(bool[] boolArray)
    //{
    //    var bytes = new byte[boolArray.Length / 8 + 1];
    //    for (var i = 0; i < boolArray.Length; i++)
    //    {
    //        if (boolArray[i]) { bytes[i / 8] |= (byte)(1 << i % 8); }
    //    }
    //    return Convert.ToBase64String(bytes);
    //}

    public static string NormalizeString(string input) => string.IsNullOrEmpty(input) ? "" : input.ToLower().Replace("ä", "ae").Replace("ö", "oe").Replace("ü", "ue").Replace("ß", "ss");

    public static IEnumerable<string> ReadAsLines(string filename)
    {
        using var reader = new StreamReader(filename);
        while (!reader.EndOfStream) { yield return reader.ReadLine()!; }
    }

    internal static void DailyBackup(string filePath, string backupDir, bool success, decimal duration, bool silent = false)
    {
        try
        {
            backupDir = Path.Combine(backupDir, new CultureInfo("de-DE").DateTimeFormat.GetDayName(DateTime.Today.DayOfWeek));
            Directory.CreateDirectory(backupDir); // Sicherstellen, dass das Tages-Verzeichnis existiert
            var todaysBackupFile = Path.Combine(backupDir, Path.GetFileNameWithoutExtension(filePath) + "_" + DateTime.Now.ToString("yyyy_MM_dd") + Path.GetExtension(filePath));
            if (File.Exists(todaysBackupFile)) { return; }  // Überprüfen, ob bereits ein Backup für heute existiert
            File.Copy(filePath, todaysBackupFile, true);
            var existingBackups = Directory.GetFiles(backupDir, Path.GetFileNameWithoutExtension(filePath) + "*.adb");
            if (existingBackups.Length >= 2) { File.Delete(existingBackups.OrderBy(f => new FileInfo(f).CreationTime).First()); }
            if (success && !silent)
            {
                var okButton = TaskDialogButton.OK;
                var page = new TaskDialogPage()
                {
                    SizeToContent = true,
                    AllowCancel = true,
                    Caption = Application.ProductName,
                    Heading = "Die lokale Datenbank wurde gesichert.",
                    Text = todaysBackupFile,
                    Icon = TaskDialogIcon.ShieldSuccessGreenBar,
                    Buttons = { TaskDialogButton.OK },
                };
                using var timer = new Timer() { Enabled = true, Interval = (int)duration };
                timer.Tick += (s, e) =>
                {
                    page.BoundDialog?.Close();
                    timer.Enabled = false;
                };
                TaskDialog.ShowDialog(page);
            }
        }
        catch (Exception ex) { ErrTaskDlg(IntPtr.Zero, ex); }
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
