//¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯//
// Konsolenhilfsprogramm für LibreOffice. Die Kompilierung muss für .NETFramework,Version=v4.8.1 erstellt werden!  //
//_________________________________________________________________________________________________________________//
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using uno.util;
using unoidl.com.sun.star.awt;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.text;

internal class Program
{
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    private static void Main(string[] args)
    {
        if (args.Length == 0) { Console.WriteLine("Es wurde keine Argumente übergeben."); }
        else
        {
            var bookmarksFound = false;
            XWindowPeer xWindowPeer = null;
            try
            {
                //var receivedData = new Dictionary<string, string> { { "Anrede", "Fräulein" }, { "PLZ_Ort", "12345 Entenhausen" } };
                var receivedData = JsonConvert.DeserializeObject<Dictionary<string, string>>(args[0]);
                var xContext = Bootstrap.bootstrap();
                var xServiceManager = (XMultiServiceFactory)xContext.getServiceManager();
                var xComponentLoader = (XComponentLoader)xServiceManager.createInstance("com.sun.star.frame.Desktop");
                var xDesktop = (XDesktop)xComponentLoader; // Desktop-Instanz holen
                var components = xDesktop.getComponents(); // Alle geöffneten Komponenten abrufen
                var xEnumerationAccess = components;
                var xEnum = xEnumerationAccess.createEnumeration();
                while (xEnum.hasMoreElements())
                {
                    var elementAny = xEnum.nextElement(); // UNO-Objekt extrahieren
                    var element = elementAny.Value;
                    if (element is XTextDocument xTextDoc && element is XModel xModel) // Prüfen, ob es ein Textdokument ist    
                    {
                        var xBookmarksSupplier = (XBookmarksSupplier)xTextDoc;
                        var xController = xModel.getCurrentController();
                        var xFrame = xController.getFrame();

                        xWindowPeer = (XWindowPeer)xFrame.getContainerWindow(); // soll vor continue stehen, damit es im Fehlerfall noch gesetzt ist 

                        var xBookmarks = xBookmarksSupplier.getBookmarks();
                        if (!xBookmarks.hasElements()) { continue; } // Wenn keine Lesezeichen vorhanden sind, nächstes Dokument   
                        bookmarksFound = true; // Bereitet Workaround vor, um irgendein Fenster in den Vordergrund zu bringen, wenn kein Dokument mit Lesezeichen gefunden wurde

                        TrySetForeground(xWindowPeer);
                        xFrame.activate(); // Fenster aktivieren und sichtbar machen
                        xFrame.getContainerWindow().setVisible(true);
                        foreach (var entry in receivedData)
                        {
                            var bookmark = entry.Key;
                            if (xBookmarks.hasByName(bookmark))
                            {
                                var bookmarkObj = xBookmarks.getByName(bookmark);
                                var xBookmark = (XTextContent)bookmarkObj.Value;
                                var xTextRange = xBookmark.getAnchor();
                                xTextRange.setString(entry.Value);
                            }
                        }
                        break;
                    }
                }
            }
            catch (unoidl.com.sun.star.uno.Exception ex) { Console.WriteLine(ex.GetType().ToString() + ": " + ex.Message); }
            catch (Exception ex) { Console.WriteLine(ex.GetType().ToString() + ": " + ex.Message); }
            if (!bookmarksFound && xWindowPeer != null) { TrySetForeground(xWindowPeer); } // für den Fall, dass kein Dokument mit Bookmarks gefunden wurde, aber ein Writer-Fenster mit geöffnetem Dokument existiert   
        }
    }

    private static void TrySetForeground(XWindowPeer xWindow)
    {
        if (xWindow is XSystemDependentWindowPeer xSysDepPeer)
        {
            var anyHandle = xSysDepPeer.getWindowHandle(new byte[0], 1); // 1 = SYSTEM_WIN32
            var handleValue = anyHandle.Value;
            var hwnd = IntPtr.Zero;
            if (handleValue is long l) { hwnd = new IntPtr(l); } // 64‑Bit Handle
            else if (handleValue is int i) { hwnd = new IntPtr(i); } // 32‑Bit Handle
            if (hwnd != IntPtr.Zero) { SetForegroundWindow(hwnd); }
        }
    }

}