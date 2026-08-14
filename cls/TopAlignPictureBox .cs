using System.ComponentModel;
using System.Drawing.Drawing2D;

namespace Adressen.cls;

internal class TopAlignZoomPictureBox : PictureBox
{
    public TopAlignZoomPictureBox()
    {
        SizeMode = PictureBoxSizeMode.Normal; // Wir machen das Zoom-Scaling selbst
        DoubleBuffered = true;
    }

    [Browsable(true)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
    [TypeConverter(typeof(ImageConverter))]
    public new Image? Image
    {
        get => base.Image;
        set
        {
            if (ReferenceEquals(base.Image, value)) { return; }
            var oldImage = base.Image; // Altes Bild zwischenspeichern
            base.Image = value; // Neues Bild setzen
            oldImage?.Dispose(); // Altes Bild (und seine GDI-Ressource) freigeben!
            UpdateHeightForImage(); // Höhe anpassen; Resize-Event wird hierfür nicht mehr benötigt
            //Invalidate(); // neu zeichnen
        }
    }


    private void UpdateHeightForImage()
    {
        var scaledHeight = 0;  // kein Bild -> Höhe 0
        if (base.Image is not null)
        {
            var pbWidth = Width;
            if (pbWidth == 0) { return; } // Wenn die Box noch keine Breite hat (z.B. beim Initialisieren), nichts tun
            var img = base.Image;
            scaledHeight = img.Width < pbWidth ? img.Height : (int)((double)img.Height * pbWidth / img.Width); // verbesserte Präzision durch Double-Berechnung
        }
        //SetHeightAndRefreshParent(scaledHeight);
        if (Height == scaledHeight) { return; }
        // Die Höhenänderung verschiebt die darunter angedockten Geschwister (die Foto-Toolbar).
        // Das eigene SuspendLayout half dabei nicht, denn angeordnet wird im ÜBERGEORDNETEN Panel.
        // Ohne das anschließende Neuzeichnen blieb der verlassene Bereich stehen – die Toolbar war
        // dann kurzzeitig doppelt zu sehen (alte und neue Position übereinander).
        var parent = Parent;
        parent?.SuspendLayout();
        Height = scaledHeight;
        if (parent is null) { return; }
        parent.ResumeLayout(true);  // Anordnung in einem einzigen Durchgang
        parent.Invalidate(true);    // freigewordenen Bereich samt Kind-Controls verwerfen
        parent.Update();            // und sofort neu zeichnen, nicht erst bei der nächsten Leerlaufnachricht
    }

    protected override void OnPaint(PaintEventArgs pe)
    {
        if (Image is null)
        {
            base.OnPaint(pe);
            return;
        }
        var img = Image;
        var pbRect = ClientRectangle;
        Rectangle destRect; // Ziel-Rechteck für das Bild
        if (img.Width < pbRect.Width) { destRect = new Rectangle(0, 0, img.Width, img.Height); }
        else
        {
            var w = pbRect.Width;  // HINWEIS: Diese Logik muss exakt dieselbe sein wie in UpdateHeightForImage
            var h = (int)((double)img.Height * pbRect.Width / img.Width);
            destRect = new Rectangle(0, 0, w, h);
        }
        pe.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        pe.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        pe.Graphics.SmoothingMode = SmoothingMode.HighQuality;
        pe.Graphics.DrawImage(img, destRect);
    }
}
