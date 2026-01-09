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
            Application.SetColorMode(SystemColorMode.System); // .NET 10 unterstützt Dark Mode nativ! 
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
        catch (Exception ex)
        {
            var errorMsg = $"Kritischer Fehler beim Start:\n\n{ex.Message}";
            if (ex.InnerException != null) { errorMsg += $"\n\nDetails: {ex.InnerException.Message}"; }
            MessageBox.Show(errorMsg, "Startfehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}