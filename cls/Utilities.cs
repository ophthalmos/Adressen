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
using System.Xml.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
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
        Check("Unternehmen", old.Unternehmen, current.Unternehmen);
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


    internal static void WordInfoTaskDlg(nint hwnd, string[] allKeys, TaskDialogIcon icon, Word.Application? wordApp, Word.Document? wordDoc)
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
        btnCreateDoc.Click += (s, e) => { CreateWordDocument(allKeys, hwnd, wordApp, wordDoc); };
        TaskDialog.ShowDialog(hwnd, page);
    }

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

            // Falls du mehr Details wie Unternehmen/Ort anzeigen willst, 
            // musst du diese ggf. in das Interface aufnehmen oder hier casten:
            var zusatzInfo = "";
            if (contact is Contact c)
            {
                zusatzInfo = $"\n{c.Unternehmen}\n{c.Strasse}\n{c.PLZ} {c.Ort}".Trim();
            }
            else if (contact is Adresse a)
            {
                zusatzInfo = $"\n{a.Unternehmen}\n{a.Strasse}\n{a.PLZ} {a.Ort}".Trim();
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

    internal static bool IsWordInstalled => Type.GetTypeFromProgID("Word.Application") is not null;

    internal static bool IsLibreOfficeInstalled => Type.GetTypeFromProgID("com.sun.star.ServiceManager") is not null;

    private static void CreateWordDocument(string[] allKeys, nint handle, Word.Application? wordApp, Word.Document? wordDoc)
    {
        if (!IsWordInstalled)
        {
            MsgTaskDlg(handle, "Microsoft Word is not installed", "Installieren Sie Microsoft Word.");
            return;
        }
        try
        {
            NativeMethods.SHGetKnownFolderPath(new Guid("374DE290-123F-4565-9164-39C4925E467B"), 0, IntPtr.Zero, out var downloadPath); // Downloads folder    
            downloadPath = Path.Combine(downloadPath, "Adressen-Vorlage.dotx");

            Control? owner = null;
            try { owner = Control.FromHandle(handle); }
            catch { owner = null; }
            try
            {
                try { wordApp = (Word.Application?)Marshal2.GetActiveObject("Word.Application"); }
                catch (COMException) { wordApp = null; }
                if (wordApp != null)
                {
                    var docCount = wordApp.Documents.Count; // Office Collections sind 1-basiert!
                    for (var i = 1; i <= docCount; i++)
                    {
                        try
                        {
                            var openDoc = wordApp.Documents[i]; // Zugriff per Index statt Enumerator
                            if (!string.IsNullOrEmpty(openDoc.FullName) &&
                                string.Equals(Path.GetFullPath(openDoc.FullName), Path.GetFullPath(downloadPath), StringComparison.OrdinalIgnoreCase))
                            {
                                wordApp.Activate();
                                openDoc.Activate();
                                return;
                            }
                        }
                        catch (Exception) { } // Falls die Prüfung fehlschlägt, einfach weiterfahren (keine Blockade)
                    }
                }
            }
            catch (Exception) { } // Falls die Prüfung fehlschlägt, einfach weiterfahren (keine Blockade)
            if (File.Exists(downloadPath))
            {
                var (IsYes, IsNo, _) = YesNo_TaskDialog(owner, "Datei existiert bereits", "Möchten Sie sie die vorhandene Vorlage löschen und neu erstellen?", downloadPath, "Ja, löschen und neu erstellen", "Nein, nur öffnen", true);
                if (IsNo)
                {
                    try // Öffnet die Datei selbst (Vorlage), nicht: neues Dokument aus der Vorlage
                    {
                        wordApp ??= new Word.Application { Visible = true };
                        var openedDoc = wordApp.Documents.Open(FileName: downloadPath, ReadOnly: false, AddToRecentFiles: true);
                        wordApp.Activate();
                        openedDoc.Activate();
                    }
                    catch { StartFile(owner?.Handle ?? IntPtr.Zero, downloadPath); }
                    return;
                }
                else if (!IsYes) { return; }
                try { File.Delete(downloadPath); }
                catch (Exception ex) { ErrTaskDlg(handle, ex); return; }
            }

            wordApp = new Word.Application { Visible = true };
            wordDoc = wordApp.Documents.Add();

            wordDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.5f);
            wordDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1.0f);

            wordDoc.Styles[Word.WdBuiltinStyle.wdStyleNormal].Font.Name = "Calibri";
            wordDoc.Styles[Word.WdBuiltinStyle.wdStyleNormal].Font.Size = 11;

            wordDoc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyAuthor].Value = "Wilhelm Happe";
            wordDoc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyTitle].Value = "Adressen-Vorlage";
            wordDoc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertySubject].Value = "Nur als Beispiel gedacht";
            wordDoc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyKeywords].Value = "Adressen, Briefvorlage";
            wordDoc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyComments].Value = "";

            var para0 = wordDoc.Paragraphs.Add();
            para0.Range.Font.Size = 12; // explizit 12 behalten
            para0.Range.Text = "Präfix_Vorname_Zwischenname_Nachname";
            wordDoc.Bookmarks.Add("Präfix_Vorname_Zwischenname_Nachname", para0.Range);
            para0.Format.SpaceAfter = 0f;
            para0.Range.InsertParagraphAfter();

            var para1 = wordDoc.Paragraphs.Add();
            para1.Range.Font.Size = 12; // explizit 12 behalten
            para1.Range.Text = "Strasse";
            para1.Range.Bookmarks.Add("Strasse", para1.Range);
            para1.Format.SpaceAfter = 6f;
            para1.Range.InsertParagraphAfter();

            var para2 = wordDoc.Paragraphs.Add();
            para2.Range.Font.Size = 12; // explizit 12 behalten
            para2.Range.Text = "PLZ_Ort";
            para2.Range.Bookmarks.Add("PLZ_Ort", para2.Range);
            para2.Format.SpaceAfter = 12f;
            para2.Range.InsertParagraphAfter();

            var para4 = wordDoc.Paragraphs.Add();
            para4.Range.Font.Size = 11;
            para4.Range.Text = "Probieren Sie nun das Einfügen einer Adresse aus, indem Sie im Adressen-Programm eine Adresse selektieren und dann auf den Button »In Brief einfügen« klicken. Wiederholen Sie den Vorgang mit anderen Adressen!";
            para4.Range.InsertParagraphAfter();

            var para5 = wordDoc.Paragraphs.Add();
            para5.Range.Font.Size = 11;
            para5.Range.Text = "Wenn Sie eine Textmarke hinzufügen möchten, markieren Sie zuerst die Stelle der Textmarke in Ihrem Dokument. Wählen Sie die Registerkarte »Einfügen« und dann »Textmarke« aus. Schneller geht es, wenn Sie die Tastenkombination Strg+Shift+F5 drücken.";
            para5.Range.InsertParagraphAfter();

            var para6 = wordDoc.Paragraphs.Add();
            para6.Range.Font.Size = 11;
            para6.Range.Text = "Um Textmarken-Klammen anzuzeigen, klicken Sie auf Datei > Optionen > Erweitert. Wählen Sie unter \"Dokumentinhalt anzeigen\" die Option \"Textmarken anzeigen\".";
            para6.Range.InsertParagraphAfter();

            var para7 = wordDoc.Paragraphs.Add();
            para7.Range.Font.Size = 11;
            para7.Range.Text = "Kombinierte Textmarken (mit Unterstrich) sind nützlich, weil bei ihnen unnötige Leerzeichen zwischen den Elementen automatisch entfernt werden.";
            para7.Range.InsertParagraphAfter();

            var para8 = wordDoc.Paragraphs.Add();
            para8.Range.Font.Bold = 1;
            para8.Range.Text = "Liste der möglichen Textmarkierungen:";
            para8.Format.SpaceAfter = 0f;
            para8.Range.InsertParagraphAfter();

            var para9 = wordDoc.Paragraphs.Add();
            para9.Range.Font.Name = "Courier New";
            para9.Range.NoProofing = 1;
            para9.Range.Text = string.Join(Environment.NewLine, allKeys);  // Zeilenumbruch im Text: \x0B

            wordDoc.SaveAs2(downloadPath, Word.WdSaveFormat.wdFormatXMLTemplate);
            wordApp.Activate();
            //wordApp.Dialogs[Word.WdWordDialog.wdDialogFileSummaryInfo].Show(); // Öffnet den Eigenschaften-Dialog. Cave: Blockiert UI
        }
        catch (Exception ex) { ErrTaskDlg(handle, ex); }
        finally { ReleaseWordObjects(ref wordDoc, ref wordApp); }
    }

    internal static void ReleaseWordObjects(ref Word.Document? wordDoc, ref Word.Application? wordApp)
    {
        if (wordDoc is not null)
        {
            try { Marshal.FinalReleaseComObject(wordDoc); }
            catch { }
            finally { wordDoc = null; }
        }
        if (wordApp is not null)
        {
            try { Marshal.FinalReleaseComObject(wordApp); } // 'Visible = true' => Word bleibt offen
            catch { }
            finally { wordApp = null; }
        }
        try
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        catch { }
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

    internal static string GetGooglePhoneByType(Person person, string type)
    {
        foreach (var phone in person.PhoneNumbers ?? []) // falls PhoneNumbers null ist, wird die Schleife dank ?? [] einfach übersprungen
        {
            if (phone.Type?.Contains(type, StringComparison.OrdinalIgnoreCase) == true) { return phone.Value ?? string.Empty; }
        }
        return string.Empty;
    }

    public static IEnumerable<string> ReadAsLines(string filename)
    {
        using var reader = new StreamReader(filename);
        while (!reader.EndOfStream) { yield return reader.ReadLine()!; }
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

            if (File.Exists(todaysBackupFile))
            {
                return;
            }

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
