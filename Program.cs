using Adressen.frm;

namespace Adressen;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        using Mutex singleMutex = new(true, "{0d16d58e-f98e-4055-9af4-e222e85d7449}", out var isNewInstance);
        if (isNewInstance)
        {
            ApplicationConfiguration.Initialize();
            //#pragma warning disable WFO5001 
            //Application.SetColorMode(SystemColorMode.System);  // nicht löschen, wird unter .NET 9 funktionieren
            //#pragma warning restore WFO5001
            var showSplash = !args.Contains("-nosplash", StringComparer.OrdinalIgnoreCase);
            FrmSplashScreen? splashScreen = null;
            if (showSplash)
            {
                splashScreen = new FrmSplashScreen();
                splashScreen.Show();
                Application.DoEvents(); // Verarbeitet alle ausstehenden Nachrichten, um sicherzustellen, dass der Splash Screen sofort gezeichnet wird
            }
            Application.Run(new FrmAdressen(splashScreen, args));
        }
        else { MessageBox.Show("Adressen wird bereits ausgeführt!", "Adressen"); }
    }
}