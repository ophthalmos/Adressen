using System.Runtime.InteropServices;
using Adressen.Properties; // Für Resources
using Word = Microsoft.Office.Interop.Word;

namespace Adressen.cls;

internal class WordManager
{
    internal static bool IsWordInstalled => Type.GetTypeFromProgID("Word.Application") is not null;

    internal static bool IsLibreOfficeInstalled => Type.GetTypeFromProgID("com.sun.star.ServiceManager") is not null;

    public static void TransferDataToActiveDocument(Dictionary<string, string> bookmarkData, IntPtr ownerHandle)
    {
        Word.Application? wordApp = null;
        Word.Document? wordDoc = null;

        try
        {
            try { wordApp = (Word.Application?)Marshal2.GetActiveObject("Word.Application"); }
            catch (Exception) { wordApp = null; }
            if (wordApp == null)
            {
                wordApp = new Word.Application { Visible = true };
                wordApp.Dialogs[Word.WdWordDialog.wdDialogFileNew].Show();
            }
            if (wordApp != null)
            {
                wordApp.Visible = true;
                try
                {
                    wordApp.Activate();
                    var hwnd = new IntPtr(wordApp.ActiveWindow.Hwnd);
                    if (hwnd != IntPtr.Zero) { NativeMethods.SetForegroundWindow(hwnd); }
                }
                catch { } // Ignorieren, falls Dialog offen
            }
            else { return; }
            if (wordApp.Documents.Count == 0) { return; }
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
            else { ShowWordBookmarksInfoDialog(ownerHandle, [.. bookmarkData.Keys]); }
        }
        catch (Exception ex) { Utils.ErrTaskDlg(ownerHandle, ex); }
        finally
        {
            wordApp?.ScreenUpdating = true;
            ReleaseWordObjects(ref wordDoc, ref wordApp);
        }
    }

    public static void CreateTemplateDocument(string[] allKeys, IntPtr ownerHandle)
    {
        if (!WordManager.IsWordInstalled)
        {
            Utils.MsgTaskDlg(ownerHandle, "Microsoft Word fehlt", "Bitte installieren Sie Microsoft Word.");
            return;
        }

        Word.Application? wordApp = null;
        Word.Document? wordDoc = null;

        try
        {
            NativeMethods.SHGetKnownFolderPath(new Guid("374DE290-123F-4565-9164-39C4925E467B"), 0, IntPtr.Zero, out var downloadPath);
            downloadPath = Path.Combine(downloadPath, "Adressen-Vorlage.dotx");
            var owner = Control.FromHandle(ownerHandle);
            try { wordApp = (Word.Application?)Marshal2.GetActiveObject("Word.Application"); }
            catch { wordApp = null; }
            if (wordApp != null)
            {
                for (var i = 1; i <= wordApp.Documents.Count; i++) // Prüfen ob Datei schon offen ist
                {
                    try
                    {
                        var doc = wordApp.Documents[i];
                        if (string.Equals(doc.FullName, downloadPath, StringComparison.OrdinalIgnoreCase))
                        {
                            doc.Activate();
                            wordApp.Visible = true;
                            return; // Schon offen, fertig
                        }
                    }
                    catch { }
                }
            }
            if (File.Exists(downloadPath))
            {
                var (IsYes, IsNo, _) = Utils.YesNo_TaskDialog(owner, "Datei existiert bereits",
                    "Möchten Sie die vorhandene Vorlage löschen und neu erstellen?", downloadPath,
                    "Ja, löschen und neu erstellen", "Nein, nur öffnen", true);
                if (IsNo)
                {
                    Utils.StartFile(ownerHandle, downloadPath);  // Nur öffnen
                    return;
                }
                else if (!IsYes) { return; } // Abbrechen

                try { File.Delete(downloadPath); }
                catch (Exception ex) { Utils.ErrTaskDlg(ownerHandle, ex); return; }
            }
            wordApp ??= new Word.Application { Visible = true };
            wordDoc = wordApp.Documents.Add();

            wordDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.5f);
            wordDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1.0f);

            var style = wordDoc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
            style.Font.Name = "Calibri";
            style.Font.Size = 11;

            var props = wordDoc.BuiltInDocumentProperties;
            props[Word.WdBuiltInProperty.wdPropertyTitle].Value = "Adressen-Vorlage";
            props[Word.WdBuiltInProperty.wdPropertyAuthor].Value = "AdressenApp";

            AddParagraph(wordDoc, "Präfix_Vorname_Zwischenname_Nachname", 12, 0, true);
            AddParagraph(wordDoc, "Strasse", 12, 6, true);
            AddParagraph(wordDoc, "PLZ_Ort", 12, 12, true);

            AddParagraph(wordDoc, "Probieren Sie nun das Einfügen einer Adresse aus, indem Sie im Adressen-Programm auf »In Brief einfügen« klicken.", 11);
            AddParagraph(wordDoc, "Liste der möglichen Textmarkierungen:", 11, 0, false, true); // Bold

            var listPara = wordDoc.Paragraphs.Add();
            listPara.Range.Font.Name = "Courier New";
            listPara.Range.Text = string.Join("\v", allKeys);  // => Zeilenumbruch

            wordDoc.SaveAs2(downloadPath, Word.WdSaveFormat.wdFormatXMLTemplate);
            wordApp.Activate();
        }
        catch (Exception ex) { Utils.ErrTaskDlg(ownerHandle, ex); }
        finally { ReleaseWordObjects(ref wordDoc, ref wordApp); }
    }

    private static void AddParagraph(Word.Document doc, string text, float fontSize, float spaceAfter = 0, bool asBookmark = false, bool bold = false)
    {
        var p = doc.Paragraphs.Add();
        p.Range.Font.Size = fontSize;
        p.Range.Font.Bold = bold ? 1 : 0;
        p.Range.Text = text;
        if (asBookmark) { doc.Bookmarks.Add(text, p.Range); }
        if (spaceAfter > 0) { p.Format.SpaceAfter = spaceAfter; }
        p.Range.InsertParagraphAfter();
    }

    internal static void ShowWordBookmarksInfoDialog(IntPtr ownerHandle, string[] allKeys)
    {
        var btnCreateDoc = new TaskDialogButton("Beispieldokument erstellen");
        var btnClose = TaskDialogButton.Close;
        btnCreateDoc.Click += (s, e) => { CreateTemplateDocument(allKeys, ownerHandle); };
        var page = new TaskDialogPage()
        {
            Caption = Application.ProductName,
            Heading = "Kein aktives Dokument gefunden",
            Text = "Es wurde kein offenes Word-Dokument gefunden, in das Daten eingefügt werden könnten.",
            Icon = new TaskDialogIcon(Resources.word32),
            Footnote = "Tipp: Öffnen Sie ein Dokument oder erstellen Sie eine Vorlage.",
            AllowCancel = true,
            Buttons = { btnCreateDoc, btnClose },
            Expander = new TaskDialogExpander
            {
                Text = $"Folgende Textmarken stehen zur Verfügung:\n{string.Join(", ", allKeys)}",
                CollapsedButtonText = "Verfügbare Textmarken anzeigen",
                ExpandedButtonText = "Verfügbare Textmarken ausblenden",
                Position = TaskDialogExpanderPosition.AfterText
            }
        };
        TaskDialog.ShowDialog(ownerHandle, page);
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