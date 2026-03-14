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

    public void GeneratePreviewSilently()
    {
        if (Document == null) { return; }

        var originalController = Document.PrintController;
        var previewController = new PreviewPrintController { UseAntiAlias = UseAntiAlias };

        Document.PrintController = previewController;
        Document.Print();
        Document.PrintController = originalController;

        var pageInfo = previewController.GetPreviewPageInfo();

        var fieldInfo = typeof(PrintPreviewControl).GetField("_pageInfo", BindingFlags.Instance | BindingFlags.NonPublic)
                     ?? typeof(PrintPreviewControl).GetField("pageInfo", BindingFlags.Instance | BindingFlags.NonPublic);

        if (fieldInfo != null)
        {
            if (fieldInfo.GetValue(this) is PreviewPageInfo[] oldPages)
            {
                foreach (var page in oldPages) { page.Image?.Dispose(); }  // Alte Bilder freigeben, um Speicherlecks zu vermeiden
            }
            fieldInfo.SetValue(this, pageInfo);
        }

        var positionMethod = typeof(PrintPreviewControl).GetMethod("PositionPage", BindingFlags.Instance | BindingFlags.NonPublic);
        positionMethod?.Invoke(this, null);

        Invalidate();
    }
}