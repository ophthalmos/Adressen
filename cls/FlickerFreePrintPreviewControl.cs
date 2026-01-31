using System.Drawing.Printing;
using System.Reflection;

namespace Adressen.cls;

public class FlickerFreePrintPreviewControl : PrintPreviewControl
{
    public event EventHandler? ZoomChanged;

    public FlickerFreePrintPreviewControl()
    {
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
        UpdateStyles();
    }

    protected override void OnPaintBackground(PaintEventArgs pevent)
    { /* Flackern verhindern */
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        if (Focused)
        {
            var newZoom = Zoom * (e.Delta > 0 ? 1.1 : 0.9);
            Zoom = Math.Clamp(newZoom, 0.1, 5.0);
            ZoomChanged?.Invoke(this, EventArgs.Empty);
            if (e is HandledMouseEventArgs he) { he.Handled = true; }
        }
        else { base.OnMouseWheel(e); }
    }

    /// <summary>
    /// Generiert die Vorschau OHNE den Status-Dialog ("Seite wird gedruckt...").
    /// Umgeht die interne ComputePreview-Methode von .NET.
    /// </summary>
    public void GeneratePreviewSilently()
    {
        if (Document == null) { return; }

        // 1. Controller sichern
        var originalController = Document.PrintController;

        // 2. Stummen PreviewController nutzen
        var previewController = new PreviewPrintController { UseAntiAlias = UseAntiAlias };

        // 3. Drucken (ohne UI)
        Document.PrintController = previewController;
        Document.Print();

        // 4. Controller zurücksetzen
        Document.PrintController = originalController;

        // 5. PreviewPageInfo per Reflection injizieren
        var pageInfo = previewController.GetPreviewPageInfo();
        var fieldInfo = typeof(PrintPreviewControl).GetField("_pageInfo", BindingFlags.Instance | BindingFlags.NonPublic)
                     ?? typeof(PrintPreviewControl).GetField("pageInfo", BindingFlags.Instance | BindingFlags.NonPublic);
        fieldInfo?.SetValue(this, pageInfo);

        // 6. LAYOUT-UPDATE ERZWINGEN (Korrektur)
        // Statt Zoom = Zoom rufen wir direkt die interne Layout-Methode auf.
        // Das berechnet die Scrollbalken und Seitenpositionen neu.
        var positionMethod = typeof(PrintPreviewControl).GetMethod("PositionPage", BindingFlags.Instance | BindingFlags.NonPublic);
        positionMethod?.Invoke(this, null);
        Invalidate();  // Neu zeichnen
    }
}