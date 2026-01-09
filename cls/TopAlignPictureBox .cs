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
            UpdateHeightForImage();
            Invalidate(); // neu zeichnen
        }
    }


    private void UpdateHeightForImage()
    {
        if (base.Image is null)
        {
            Height = 0; // Höhe zurücksetzen, wenn kein Bild gesetzt ist    
            return;
        }
        var img = base.Image;
        var pbWidth = Width;
        if (pbWidth == 0) { return; } // Wenn die Box noch keine Breite hat (z.B. beim Initialisieren), nichts tun
        var scaledHeight = img.Width < pbWidth ? img.Height : (int)((double)img.Height * pbWidth / img.Width); // verbesserte Präzision durch Double-Berechnung
        if (Height != scaledHeight)
        {
            SuspendLayout();
            Height = scaledHeight;
            ResumeLayout();
        }
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        ////UpdateHeightForImage(); // Invalidate() wird durch Resize ausgelöst 
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
