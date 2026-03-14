using System.Runtime.InteropServices;
using Adressen.Properties; // Für Resources
using Word = Microsoft.Office.Interop.Word;

namespace Adressen.cls;

internal class WordManager
{
    internal static bool IsWordInstalled => Type.GetTypeFromProgID("Word.Application") is not null;

    internal static bool IsLibreOfficeInstalled => Type.GetTypeFromProgID("com.sun.star.ServiceManager") is not null;

    public static void TransferDataToActiveDocument(Dictionary<string, string> bookmarkData, nint ownerHandle)
    {
        var wordApp = default(Word.Application);
        var wordDoc = default(Word.Document);

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
                    var hwnd = (nint)wordApp.ActiveWindow.Hwnd;

                    if (hwnd != nint.Zero) { NativeMethods.SetForegroundWindow(hwnd); }
                }
                catch { }  // Ignorieren, falls Dialog offen
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
            var hr = NativeMethods.SHGetKnownFolderPath(new Guid("374DE290-123F-4565-9164-39C4925E467B"), 0, nint.Zero, out var pPath);
            Marshal.ThrowExceptionForHR(hr);

            var knownPath = Marshal.PtrToStringUni(pPath);
            Marshal.FreeCoTaskMem(pPath);

            if (knownPath == null) { return; }

            var downloadPath = Path.Combine(knownPath, "Adressen-Vorlage.dotx");
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
                            return; // Schon offen, fertig
                        }
                    }
                    catch { }  // Ignorieren, falls auf ein einzelnes Dokument nicht zugegriffen werden kann
                }
            }

            if (File.Exists(downloadPath))
            {
                var (isYes, isNo, _) = Utils.YesNo_TaskDialog(
                    owner,
                    "Datei existiert bereits",
                    "Möchtest du die vorhandene Vorlage löschen und neu erstellen?",
                    downloadPath,
                    "Ja, löschen und neu erstellen",
                    "Nein, nur öffnen",
                    true);

                if (isNo)
                {
                    Utils.StartFile(ownerHandle, downloadPath);  // Nur öffnen
                    return;
                }
                else if (!isYes) { return; }  // Abbrechen

                try { File.Delete(downloadPath); }
                catch (Exception ex)
                {
                    Utils.ErrTaskDlg(ownerHandle, ex);
                    return;
                }
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

            AddParagraph(wordDoc, "Praefix_Vorname_Zwischenname_Nachname", 12f, 0f, true);
            AddParagraph(wordDoc, "Strasse", 12f, 6f, true);
            AddParagraph(wordDoc, "PLZ_Ort", 12f, 12f, true);

            AddParagraph(wordDoc, "Probiere nun das Einfügen einer Adresse aus, indem du im Adressen-Programm auf »In Brief einfügen« klickst.", 11f);
            AddParagraph(wordDoc, "Liste der möglichen Textmarkierungen:", 11f, 0f, false, true);

            var listPara = wordDoc.Paragraphs.Add();
            listPara.Range.Font.Name = "Courier New";
            listPara.Range.Text = string.Join("\v", allKeys);

            wordDoc.SaveAs2(downloadPath, Word.WdSaveFormat.wdFormatXMLTemplate);
            wordApp.Activate();
        }
        catch (Exception ex)
        {
            Utils.ErrTaskDlg(ownerHandle, ex);
        }
        finally
        {
            ReleaseWordObjects(ref wordDoc, ref wordApp);
        }
    }

    private static void AddParagraph(Word.Document doc, string text, float fontSize, float spaceAfter = 0f, bool asBookmark = false, bool bold = false)
    {
        var p = doc.Paragraphs.Add();
        p.Range.Font.Size = fontSize;
        p.Range.Font.Bold = bold ? 1 : 0;
        p.Range.Text = text;

        if (asBookmark)
        {
            doc.Bookmarks.Add(text, p.Range);
        }

        if (spaceAfter > 0f)
        {
            p.Format.SpaceAfter = spaceAfter;
        }

        p.Range.InsertParagraphAfter();
    }

    internal static void ShowWordBookmarksInfoDialog(nint ownerHandle, string[] allKeys)
    {
        var btnCreateDoc = new TaskDialogButton("Beispieldokument erstellen");
        var btnClose = TaskDialogButton.Close;
        btnCreateDoc.Click += (s, e) => { CreateTemplateDocument(allKeys, ownerHandle); };
        var expander = new TaskDialogExpander
        {
            Text = $"Folgende Textmarken stehen zur Verfügung:\n{string.Join(", ", allKeys)}\n\nDie Namen wurden gerade in die Zwischenablage kopiert!",
            CollapsedButtonText = "Verfügbare Textmarken anzeigen",
            ExpandedButtonText = "Verfügbare Textmarken ausblenden",
            Position = TaskDialogExpanderPosition.AfterText
        };

        expander.ExpandedChanged += (s, e) =>
        {
            if (expander.Expanded && allKeys != null && allKeys.Length > 0) { Clipboard.SetText(string.Join(Environment.NewLine, allKeys)); }
        };
        var page = new TaskDialogPage()
        {
            Caption = Application.ProductName,
            Heading = "Kein aktives Word-Dokument gefunden",
            Text = "Es wurde kein offenes Dokument gefunden, in das Daten eingefügt werden könnten.",
            Icon = new TaskDialogIcon(Resources.word32),
            Footnote = "Tipp: Öffne ein Dokument oder erstelle eine Vorlage.",
            AllowCancel = true,
            Buttons = { btnCreateDoc, btnClose },
            Expander = expander
        };
        TaskDialog.ShowDialog(ownerHandle, page);
    }

    private static void ReleaseWordObjects(ref Word.Document? wordDoc, ref Word.Application? wordApp)
    {
        if (wordDoc != null)
        {
            try
            {
                Marshal.FinalReleaseComObject(wordDoc);
            }
            catch
            {
            }
            finally
            {
                wordDoc = null;
            }
        }

        if (wordApp != null)
        {
            try
            {
                Marshal.FinalReleaseComObject(wordApp);
            }
            catch
            {
            }
            finally
            {
                wordApp = null;
            }
        }

        try
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        catch
        {
        }
    }
}