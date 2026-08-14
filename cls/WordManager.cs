using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace Adressen.cls;

internal class WordManager
{
    internal static bool IsWordInstalled => Type.GetTypeFromProgID("Word.Application") is not null;

    internal static bool IsLibreOfficeInstalled => Type.GetTypeFromProgID("com.sun.star.ServiceManager") is not null;

    // Rückgabe: true, wenn der Anwender im Info-Dialog in die Programmeinstellungen wechseln möchte (s. ShowWordBookmarksInfoDialog)
    public static bool TransferDataToActiveDocument(Dictionary<string, string> bookmarkData, nint ownerHandle)
    {
        var wordApp = default(Word.Application);
        var wordDoc = default(Word.Document);
        var openSettings = false;

        try
        {
            try { wordApp = (Word.Application?)Marshal2.GetActiveObject("Word.Application"); }
            catch (Exception) { wordApp = null; }

            if (wordApp == null)
            {
                wordApp = new Word.Application { Visible = true };
                // Adressenprogramm hinter alle Fenster schieben, damit Words Dialog sichtbar vorne erscheint. SWP_NOACTIVATE: keine Fokusänderung, nur Z-Reihenfolge ändern.
                NativeMethods.SetWindowPos(ownerHandle, NativeMethods.HWND_BOTTOM, 0, 0, 0, 0, NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOACTIVATE);
                wordApp.Activate();
                wordApp.Dialogs[Word.WdWordDialog.wdDialogFileNew].Show();  // blockiert, bis Word-Dialog geschlossen wird (synchroner, blockierender COM-Aufruf)
            }
            else
            {
                wordApp.Visible = true;
                NativeMethods.SetWindowPos(ownerHandle, NativeMethods.HWND_BOTTOM, 0, 0, 0, 0, NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOACTIVATE);
                wordApp.Activate();
            }
            if (wordApp == null || wordApp.Documents.Count == 0) { return false; }

            wordDoc = wordApp.ActiveDocument;

            if (wordDoc != null)
            {
                wordApp.ScreenUpdating = false; // Performance boost
                foreach (var entry in bookmarkData)
                {
                    if (wordDoc.Bookmarks.Exists(entry.Key))
                    {
                        var bm = wordDoc.Bookmarks[entry.Key];
                        var range = bm.Range;
                        range.Text = entry.Value;
                        wordDoc.Bookmarks.Add(entry.Key, range);  // Textmarke wiederherstellen (da sie durch das Ersetzen gelöscht wird!)
                    }
                }
                wordApp.ScreenUpdating = true;
            }
            else { openSettings = ShowWordBookmarksInfoDialog(ownerHandle, [.. bookmarkData.Keys]); }
        }
        catch (Exception ex) { Utils.ErrTaskDlg(ownerHandle, ex); }
        finally
        {
            wordApp?.ScreenUpdating = true;
            ReleaseWordObjects(ref wordDoc, ref wordApp);
        }
        return openSettings;
    }

    public static void ConnectMailMergeDataSource(string templatePath, string csvPath, nint ownerHandle)  // Seriendruck: CSV-Datenquelle an Word-Dokument anhängen
    {
        if (!IsWordInstalled)
        {
            Utils.MsgTaskDlg(ownerHandle, "Microsoft Word fehlt", "Bitte installiere Microsoft Word.");
            return;
        }
        var createNew = string.IsNullOrEmpty(templatePath); // leer = neues, leeres Dokument
        if (!createNew && !File.Exists(templatePath)) { Utils.MsgTaskDlg(ownerHandle, "Hauptdokument nicht gefunden", templatePath, TaskDialogIcon.ShieldWarningYellowBar); return; }
        if (!File.Exists(csvPath)) { Utils.MsgTaskDlg(ownerHandle, "Datenquelle nicht gefunden", csvPath, TaskDialogIcon.ShieldWarningYellowBar); return; }

        var wordApp = default(Word.Application);
        var wordDoc = default(Word.Document);

        try
        {
            try { wordApp = (Word.Application?)Marshal2.GetActiveObject("Word.Application"); }
            catch { wordApp = null; }

            wordApp ??= new Word.Application();
            wordApp.Visible = true;
            try { EnsureMailMergeFieldMappings(wordApp.Version); } catch { }  // Fehlschlag ist unkritisch
            if (createNew) { wordDoc = wordApp.Documents.Add(); }  // Kein Pfad → neues, leeres Dokument. Sonst: .dotx/.dotm/.dot → neues Dokument aus Vorlagendatei;
            else
            {
                var ext = Path.GetExtension(templatePath).ToLowerInvariant();
                wordDoc = ext is ".dotx" or ".dotm" or ".dot"
                    ? wordApp.Documents.Add(Template: templatePath)
                    : wordApp.Documents.Open(FileName: templatePath, ConfirmConversions: false, ReadOnly: false, AddToRecentFiles: false);
            }

            var previousAlerts = wordApp.DisplayAlerts;
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone; // keine störenden Dialoge während des Anhängens
            try
            {
                wordDoc.MailMerge.MainDocumentType = Word.WdMailMergeMainDocType.wdFormLetters;
                // schema.ini neben der CSV regelt Trennzeichen (;), Kopfzeile, UTF-8 und Spaltentypen.
                wordDoc.MailMerge.OpenDataSource(
                    Name: csvPath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    LinkToSource: true,
                    AddToRecentFiles: false,
                    Revert: false,
                    Connection: string.Empty,
                    SQLStatement: string.Empty,
                    SQLStatement1: string.Empty,
                    SubType: Word.WdMergeSubType.wdMergeSubTypeOther);
                wordDoc.MailMerge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument;
            }
            finally { wordApp.DisplayAlerts = previousAlerts; }
            NativeMethods.SetWindowPos(ownerHandle, NativeMethods.HWND_BOTTOM, 0, 0, 0, 0, NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOACTIVATE);
            wordApp.Activate();
        }
        catch (Exception ex) { Utils.ErrTaskDlg(ownerHandle, ex); }
        finally { ReleaseWordObjects(ref wordDoc, ref wordApp); }
    }

    private static void EnsureMailMergeFieldMappings(string officeVersion)  // Word: »Übereinstimmende Felder festlegen«
    {
        var uiLcid = GetOfficeUiLcid(officeVersion);  // (LCID & 0x3FF) == 0x07 deckt alle de-Varianten ab: DE, AT, CH, LU, LI.
        if (uiLcid != 0 && (uiLcid & 0x3FF) != 0x07) { return; }  //  0 = unbekannt → annehmen (DACH-only-App)
        (string Column, string WordField)[] mappings =
        [
            ("Zwischenname", "Weitere Vornamen"),
            ("Unternehmen",  "Firma"),
            ("Adresse",      "Adresse 1"),
            ("Telefon1",     "Telefon Büro"),
            ("Telefon2",     "Telefon (privat)"),
            ("Mail1",        "E-Mail-Adresse"),
            ("Internet",     "Webseite"),
        ];
        var basePath = $@"Software\Microsoft\Office\{officeVersion}\Common\DataServices\MMMatchedFields";
        foreach (var (column, wordField) in mappings)
        {
            using var key = Registry.CurrentUser.CreateSubKey($@"{basePath}\{column}");
            key?.SetValue(wordField, Array.Empty<byte>(), RegistryValueKind.None);
        }
    }

    private static int GetOfficeUiLcid(string officeVersion)
    {
        using var key = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{officeVersion}\Common\LanguageResources");  // 1031 = Deutsch (de-DE)
        return (key?.GetValue("UILanguage") ?? key?.GetValue("InstallLanguage")) is int lcid ? lcid : 0;  // 0 = unbekannt
    }

    internal static string GetDownloadsPath()
    {
        var hr = NativeMethods.SHGetKnownFolderPath(new Guid("374DE290-123F-4565-9164-39C4925E467B"), 0, nint.Zero, out var pPath);
        Marshal.ThrowExceptionForHR(hr);
        var path = Marshal.PtrToStringUni(pPath) ?? throw new DirectoryNotFoundException("Downloads-Ordner nicht gefunden.");
        Marshal.FreeCoTaskMem(pPath);
        return path;
    }

    public static void CreateTemplateDocument(string[] allKeys, nint ownerHandle)
    {
        if (!IsWordInstalled)
        {
            Utils.MsgTaskDlg(ownerHandle, "Microsoft Word fehlt", "Bitte installiere Microsoft Word.");
            return;
        }

        var wordApp = default(Word.Application);
        var wordDoc = default(Word.Document);

        try
        {
            var downloadPath = Path.Combine(GetDownloadsPath(), "TextmarkenBeispiel.dotx");
            var owner = Control.FromHandle(ownerHandle);

            try { wordApp = (Word.Application?)Marshal2.GetActiveObject("Word.Application"); }
            catch { wordApp = null; }

            if (wordApp != null)
            {
                for (var i = 1; i <= wordApp.Documents.Count; i++) // Prüfen, ob Datei schon offen ist
                {
                    try
                    {
                        var doc = wordApp.Documents[i];
                        if (string.Equals(doc.FullName, downloadPath, StringComparison.OrdinalIgnoreCase))
                        {
                            doc.Activate();
                            wordApp.Visible = true;
                            // Adressenprogramm nach hinten schieben, sonst bleibt es vor Word (Foreground-Lock von Windows)
                            NativeMethods.SetWindowPos(ownerHandle, NativeMethods.HWND_BOTTOM, 0, 0, 0, 0, NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOACTIVATE);
                            wordApp.Activate();
                            return; // Schon offen, fertig
                        }
                    }
                    catch { }  // Ignorieren, falls auf ein einzelnes Dokument nicht zugegriffen werden kann
                }
            }

            if (File.Exists(downloadPath))
            {
                var (isYes, isCancelled) = Utils.YesNo_TaskDialog(
                    owner,
                    "Datei existiert bereits",
                    "Möchtest du die vorhandene Vorlage löschen und neu erstellen?",
                    downloadPath,
                    "Ja, löschen und neu erstellen",
                    "Nein, nur öffnen");

                if (isCancelled) { return; }  // Abbrechen
                else if (!isYes)  // isNo: Datei nur öffnen, nicht löschen
                {
                    Utils.StartFile(ownerHandle, downloadPath);  // Nur öffnen
                    return;
                }
                try { File.Delete(downloadPath); }
                catch (Exception ex)
                {
                    Utils.ErrTaskDlg(ownerHandle, ex);
                    return;
                }
            }

            wordApp ??= new Word.Application { Visible = true };
            wordDoc = wordApp.Documents.Add();
            wordDoc.ShowSpellingErrors = false;
            wordDoc.ShowGrammaticalErrors = false;

            wordDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.5f);
            wordDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1.0f);

            var style = wordDoc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
            style.Font.Name = "Calibri";
            style.Font.Size = 11;

            var props = wordDoc.BuiltInDocumentProperties;
            props[Word.WdBuiltInProperty.wdPropertyTitle].Value = "TextmarkenBeispiel";
            props[Word.WdBuiltInProperty.wdPropertyAuthor].Value = "AdressenApp";

            AddParagraph(wordDoc, "ANSCHRIFT", 12f, 0f, true);

            AddParagraph(wordDoc, "Klicke nun im Adressen & Kontakte-Programm auf »In Brief einfügen«," +
                "\vum die oben stehende Textmarke »ANSCHRIFT« mit Inhalten zu füllen." +
                "\vDen „Textmarke“-Befehl findest du auf der Registerkarte „Einfügen“." +
                "\vBitte zu Übungszwecken anklicken und im Dialog „Gehe zu“ wählen.",
                11f, 0f, false, false, "In Brief einfügen", isRed: true);
            AddParagraph(wordDoc, "Um einen der unten stehenen Begriffe als neue Textmarke hinzuzufügen," +
                "\vmusst du zuerst den Begriff selektieren (per Doppelklick) und kopieren" +
                "\v(Strg+C). Füge ihn dann im Textmarke-Dialog als neuen Textmarkenamen" +
                "\vein (Strg+V) und betätige anschließend die Schaltfläche „hinzufügen“.",
                11f, 0f, false, false, "In Brief einfügen", isRed: true);
            AddParagraph(wordDoc, "Liste der möglichen Textmarken:", 11f, 0f, false, true);

            var listPara = wordDoc.Paragraphs.Add();
            listPara.Range.Font.Name = "Courier New";
            listPara.Range.Text = string.Join("\v", allKeys);   //.OrderBy(k => k).ToArray());

            wordApp.Selection.EndKey(Word.WdUnits.wdStory);  // Cursor ans Ende stellen
            wordDoc.SaveAs2(downloadPath, Word.WdSaveFormat.wdFormatXMLTemplate);
            // Adressenprogramm nach hinten schieben, sonst bleibt es vor Word (Foreground-Lock von Windows)
            NativeMethods.SetWindowPos(ownerHandle, NativeMethods.HWND_BOTTOM, 0, 0, 0, 0, NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOACTIVATE);
            wordApp.Activate();
        }
        catch (Exception ex) { Utils.ErrTaskDlg(ownerHandle, ex); }
        finally { ReleaseWordObjects(ref wordDoc, ref wordApp); }
    }

    private static void AddParagraph(Word.Document doc, string text, float fontSize, float spaceAfter = 0f, bool asBookmark = false, bool bold = false, string? boldSubstring = null, bool isRed = false)
    {
        var p = doc.Paragraphs.Add();
        p.Range.Font.Color = isRed ? Word.WdColor.wdColorDarkRed : Word.WdColor.wdColorBlack;
        p.Range.Font.Bold = bold ? 1 : 0;  // Ganze Zeile fett, wenn bold=true
        p.Range.Font.Size = fontSize;
        p.Range.Text = text;
        if (!string.IsNullOrEmpty(boldSubstring))
        {
            var findRange = p.Range;
            if (findRange.Find.Execute(FindText: boldSubstring)) { findRange.Font.Bold = 1; }
        }
        if (asBookmark) { doc.Bookmarks.Add(text, p.Range); }
        if (spaceAfter > 0f) { p.Format.SpaceAfter = spaceAfter; }
        p.Range.InsertParagraphAfter();
    }

    internal const string LibreSampleFileName = "LesezeichenBeispiel.odt";  // liegt im Programmverzeichnis

    // Gegenstück zu CreateTemplateDocument: Ein ODF-Textdokument lässt sich nicht sinnvoll programmatisch
    // erzeugen, deshalb wird die mitgelieferte Beispieldatei in den Downloads-Ordner kopiert und von dort
    // geöffnet – so kann der Anwender sie gefahrlos verändern, das Original bleibt unangetastet.
    internal static void OpenLibreOfficeSample(nint ownerHandle)
    {
        try
        {
            var sourcePath = Path.Combine(AppContext.BaseDirectory, LibreSampleFileName);
            if (!File.Exists(sourcePath))
            {
                Utils.MsgTaskDlg(ownerHandle, "Beispieldatei nicht gefunden", sourcePath, TaskDialogIcon.ShieldWarningYellowBar);
                return;
            }

            var targetPath = Path.Combine(GetDownloadsPath(), LibreSampleFileName);
            var copyNeeded = true;

            if (File.Exists(targetPath))
            {
                var (isYes, isCancelled) = Utils.YesNo_TaskDialog(
                    Control.FromHandle(ownerHandle),
                    "Datei existiert bereits",
                    "Möchtest du die vorhandene Beispieldatei ersetzen?",
                    targetPath,
                    "Ja, ersetzen",
                    "Nein, vorhandene öffnen");

                if (isCancelled) { return; }
                copyNeeded = isYes;  // "Nein" → bestehende (evtl. bearbeitete) Datei einfach öffnen
            }

            if (copyNeeded)
            {
                try { File.Copy(sourcePath, targetPath, true); }
                catch (Exception ex)
                {
                    Utils.MsgTaskDlg(ownerHandle, "Kopieren nicht möglich",
                        $"{ex.Message}\n\nVermutlich ist die Datei noch in LibreOffice geöffnet.",
                        TaskDialogIcon.ShieldWarningYellowBar);
                    return;
                }
            }

            OpenWithLibreOffice(ownerHandle, targetPath);
        }
        catch (Exception ex) { Utils.ErrTaskDlg(ownerHandle, ex); }
    }

    // Öffnet das Dokument möglichst gezielt mit LibreOffice Writer. Nur auf die Dateizuordnung zu vertrauen
    // wäre unsicher, weil .odt auch mit Word verknüpft sein kann.
    private static void OpenWithLibreOffice(nint ownerHandle, string filePath)
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath");
            var installDir = key?.GetValue(null) as string;
            if (!string.IsNullOrEmpty(installDir))
            {
                var exePath = Path.Combine(installDir, "soffice.exe");
                if (File.Exists(exePath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = exePath,
                        Arguments = $"--writer \"{filePath}\"",
                        UseShellExecute = true
                    });
                    return;
                }
            }
        }
        catch { /* Rückfall auf die Dateizuordnung */ }
        Utils.StartFile(ownerHandle, filePath);
    }

    // Erklärungen für die kombinierten Einträge – nur für die ANZEIGE im Info-Dialog, nicht für die Zwischenablage.
    // Muss zur Zusammensetzung in FrmAdressen.FillWordProcessingDictionary passen!
    private static readonly Dictionary<string, string> CombinedKeyHints = new(StringComparer.Ordinal)
    {
        ["NAMEN"] = "Titel + Vor-, Zwischen-, Nachname + Suffix; Ø Nachname: Unternehmen",
        ["EMPFAENGER"] = "Anrede + Titel + Vor- + Nachname; Ø Nachname: Unternehmen",
        ["ANSCHRIFT"] = "EMPFAENGER + Unternehmen + Adresse + PLZ (Pf.) Ort; mehrzeilig",
        ["Praefix"] = "Titel"
    };

    // useWord: true = nur Word, false = nur LibreOffice, null = noch nicht festgelegt → beides anbieten
    // Rückgabe: true, wenn der Anwender über den Link in die Programmeinstellungen wechseln möchte
    internal static bool ShowWordBookmarksInfoDialog(nint ownerHandle, string[] allKeys, bool? useWord = true)
    {
        if (allKeys == null || allKeys.Length == 0) { return false; }
        var btnClose = TaskDialogButton.Close;
        var buttons = new TaskDialogButtonCollection();
        if (useWord != false)
        {
            var btnCreateDoc = new TaskDialogCommandLinkButton("&Microsoft-Word-Vorlage erstellen", "Beispielvorlage mit Liste möglicher Textmarken");
            btnCreateDoc.Click += (s, e) => { CreateTemplateDocument(allKeys, ownerHandle); };
            buttons.Add(btnCreateDoc);
        }
        if (useWord != true)
        {
            var btnOpenOdt = new TaskDialogCommandLinkButton("&LibreOffice-Writer-Vorlage öffnen", "Beispielvorlage mit Liste möglicher Lesezeichen");
            btnOpenOdt.Click += (s, e) => { OpenLibreOfficeSample(ownerHandle); };
            buttons.Add(btnOpenOdt);
        }

        buttons.Add(btnClose);
        var expander = new TaskDialogExpander
        {
            // Anzeige mit Erklärung der kombinierten Einträge; kopiert werden weiter nur die reinen Namen (s. u.)
            Text = $"Folgende Textmarken/Lesezeichen stehen zur Verfügung:\n\n" +
                   $"{string.Join(Environment.NewLine, allKeys.Select(k => CombinedKeyHints.TryGetValue(k, out var hint) ? $"{k} ({hint})" : k))}\n\n" +
                   "(Die Liste wurde soeben in die Zwischenablage kopiert!)",
            CollapsedButtonText = "Verfügbare Textmarken anzeigen",
            ExpandedButtonText = "Verfügbare Textmarken ausblenden",
            Position = TaskDialogExpanderPosition.AfterText,  //AfterFootnote
        };
        expander.ExpandedChanged += (s, e) =>
        {
            if (expander.Expanded) { Utils.SetClipboardText(string.Join(Environment.NewLine, allKeys)); }
        };
        var choice = useWord == true ? "Word eingestellt. Deshalb wird hier das Writer-Beispiel nicht" 
            : useWord == false ? "LibreOffice eingestellt. Deshalb wird nur das Writer-Beispiel" 
            : "„Jedesmal auswählen“ eingestellt. Deshalb werden hier beide";

        var page = new TaskDialogPage()
        {
            Caption = Application.ProductName,
            Heading = "So fügen Sie Textmarken resp. Lesezeichen in eine Vorlage ein:",
            Text = "Das Vorgehen ist in Word (Textmarken) und LibreOffice Writer (Lesezeichen) ähnlich:\n" +
                   "Cursor positionieren > Menü: Einfügen > Textmarke/Lesezeichen: Name einfügen …\n" +
                   "Der Name muss muss exakt den Vorgaben entsprechen (Groß-/Kleinschreibung).\n\n" +
                   "In den  <a href=\"OpenSettings\">Programmeinstellungen</a> kannst du Word oder Writer als Standard festlegen.\n" +
                   $"Zu Zeit ist {choice} angezeigt.",
            Icon = new TaskDialogIcon(Properties.Resources.Word32),
            AllowCancel = true,
            EnableLinks = true,
            SizeToContent = true,
            Buttons = buttons,
            Expander = expander
        };
        var openSettings = false;
        page.LinkClicked += (sender, e) =>  // MUSS vor ShowDialog stehen – danach ist der Dialog bereits wieder geschlossen
        {
            if (e.LinkHref != "OpenSettings") { return; }
            openSettings = true;
            page.BoundDialog?.Close();  // BoundDialog ist gesetzt, solange der Dialog angezeigt wird
        };
        TaskDialog.ShowDialog(ownerHandle, page);
        return openSettings;
    }

    private static void ReleaseWordObjects(ref Word.Document? wordDoc, ref Word.Application? wordApp)
    {
        if (wordDoc != null)
        {
            try { Marshal.FinalReleaseComObject(wordDoc); }
            catch { }
            finally { wordDoc = null; }
        }

        if (wordApp != null)
        {
            try { Marshal.FinalReleaseComObject(wordApp); }
            catch { }
            finally { wordApp = null; }
        }

        try
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        catch { }
    }
}