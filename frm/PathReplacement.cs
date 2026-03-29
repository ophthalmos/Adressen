using System.ComponentModel;  // Für die Attribute benötigt

namespace Adressen.frm;

public partial class PathReplacement : Form
{
    [Browsable(false)] // Versteckt die Eigenschaft im Eigenschaften-Fenster des Designers
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)] // Verhindert die Serialisierung
    public string SearchText
    {
        get => tbSearch.Text;
        set => tbSearch.Text = value;
    }

    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public string ReplaceText => tbReplace.Text;

    public PathReplacement()
    {
        InitializeComponent();
    }

    private void BtnBrowse_Click(object sender, EventArgs e)
    {
        using var fbd = new FolderBrowserDialog { Description = "Neuen Basisordner wählen" };
        if (fbd.ShowDialog() == DialogResult.OK) { tbReplace.Text = fbd.SelectedPath; }
    }

    private void PathReplacement_FormClosing(object sender, FormClosingEventArgs e)
    {
        // Wir prüfen nur, wenn der Nutzer auf "Ersetzen" (OK) geklickt hat
        if (DialogResult == DialogResult.OK)
        {
            var newText = tbReplace.Text.Trim();
            var oldText = tbSearch.Text.Trim();

            if (string.IsNullOrEmpty(oldText))
            {
                TaskDialog.ShowDialog(this, new TaskDialogPage()
                {
                    Caption = "Fehlende Eingabe",
                    Heading = "Suchtext fehlt",
                    Text = "Bitte geben Sie an, welcher Textteil ersetzt werden soll.",
                    Icon = TaskDialogIcon.Error
                });
                e.Cancel = true;
                tbSearch.Focus();
                return;
            }

            // Weiche Validierung: Sieht es aus wie ein voller Basis-Pfad (Laufwerk oder UNC)?
            if (Path.IsPathRooted(newText) && !Directory.Exists(newText))
            {
                var btnYes = new TaskDialogButton("&Ja, trotzdem verwenden");
                var btnNo = new TaskDialogButton("&Nein, ich korrigiere das");

                var page = new TaskDialogPage()
                {
                    Caption = "Ordner nicht gefunden",
                    Heading = "Ziel-Pfad existiert nicht",
                    Text = $"Das Verzeichnis '{newText}' konnte auf diesem System nicht gefunden werden.\n\n" +
                           "Möglicherweise ist es ein Tippfehler oder ein Laufwerk ist nicht verbunden. " +
                           "Möchten Sie diesen Pfad trotzdem verwenden?",
                    Icon = TaskDialogIcon.ShieldWarningYellowBar,
                    Buttons = { btnNo, btnYes },
                    DefaultButton = btnNo
                };

                var result = TaskDialog.ShowDialog(this, page);

                // Wenn der Nutzer kalte Füße bekommt, brechen wir das Schließen ab
                if (result == btnNo)
                {
                    e.Cancel = true;
                    tbReplace.Focus();
                }
            }
        }
    }
}