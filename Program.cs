namespace Adressen;

internal static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    private static void Main(string[] args)
    {
        using Mutex singleMutex = new(true, "{0d16d58e-f98e-4055-9af4-e222e85d7449}", out var isNewInstance);
        if (isNewInstance)
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
        Application.Run(new FrmAdressen(args));
        }
        else { MessageBox.Show("Adressen wird bereits ausgeführt!", "Adressen"); } // make the currently running instance jump on top of all the other windows
    }
}