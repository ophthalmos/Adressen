using System.Diagnostics;
using Adressen.cls;
using Adressen.frm;

namespace Adressen;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        using Mutex singleMutex = new(true, "{0d16d58e-f98e-4055-9af4-e222e85d7449}", out var isNewInstance);
        if (!isNewInstance)
        {
            MessageBox.Show("Adressen wird bereits ausgeführt!", "Adressen", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        try
        {
            ApplicationConfiguration.Initialize();

            // Globales Sicherheitsnetz: Ausnahmen aus Event-Handlern (auch async void nach dem ersten await) landen im UI-Thread
            // in Application.ThreadException; ohne Handler würde der Prozess ohne Meldung beendet. Muss vor dem ersten Fenster gesetzt werden.
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

            Application.SetColorMode(SystemColorMode.System); // .NET 10 unterstützt Dark Mode nativ!
            FontManager.StartPreloading();  // Vorladen so früh wie möglich anstoßen (läuft asynchron im Hintergrund)
            var showSplash = !args.Contains("-nosplash", StringComparer.OrdinalIgnoreCase);
            FrmSplashScreen? splashScreen = null;
            if (showSplash)
            {
                splashScreen = new FrmSplashScreen();
                splashScreen.Show();
                splashScreen.Refresh(); // Statt DoEvents()
            }
            Application.Run(new FrmAdressen(splashScreen, args));
        }
        catch (Exception ex) { MessageBox.Show(ex.Message + Environment.NewLine + Environment.NewLine + ex.StackTrace, "Startfehler"); }
        finally { FontManager.Cleanup(); }  // Globales Aufräumen der GDI-Ressourcen beim regulären oder fehlerhaften Beenden
    }

    private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e) => ShowUnhandled(e.Exception);  // UI-Thread: Programm läuft nach dem Dialog weiter

    private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)  // fremder Thread: Prozess wird danach beendet, wenigstens die Ursache zeigen
    {
        if (e.ExceptionObject is Exception ex) { ShowUnhandled(ex); }
    }

    private static void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)  // verwaiste Tasks: nur protokollieren, nicht abstürzen
    {
        Debug.WriteLine($"[UnobservedTaskException] {e.Exception}");
        e.SetObserved();
    }

    private static void ShowUnhandled(Exception ex)
    {
        try { Utils.ErrTaskDlg(Form.ActiveForm?.Handle, ex); }
        catch { MessageBox.Show(ex.ToString(), "Unerwarteter Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error); }  // Fallback, falls der TaskDialog selbst scheitert
    }
}
