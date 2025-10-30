using System.ComponentModel;
using System.Drawing.Drawing2D;
using Adressen.Properties;

namespace Adressen.cls;
internal class TagControl : Control
{
    private Rectangle deleteRect;
    private bool isHoveringPanel = false;
    private bool isHoveringDelete = false;
    private string _membership = string.Empty;

    private readonly Image deleteImage = Resources.delete12;
    private readonly int deleteButtonLogicalSize = 16;

    public event EventHandler? DeleteClick;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
    [Browsable(true)]
    [Bindable(true)]
    public string Membership
    {
        get => _membership;
        set => _membership = value;
    }

    public TagControl()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
        Text = "Tag";
        BackColor = Color.Transparent; // wird sonst als rechtesckiger Hintergrund gezeichnet
        ForeColor = SystemColors.ControlText;
        Margin = new Padding(2); //Padding = new Padding(0, 0, 0, 0); // Padding.Left wird für die Textposition verwendet
        Font = new Font("Segoe UI", 10);
        UpdateSize();
    }

    protected override void OnTextChanged(EventArgs e)
    {
        base.OnTextChanged(e);
        UpdateSize();
    }

    private void UpdateSize()
    {
        var panelHeight = 23;
        var radius = panelHeight / 2.0f;
        var textSize = TextRenderer.MeasureText(Text, Font);
        var panelWidth = Padding.Left + textSize.Width + (deleteButtonLogicalSize / 2) + (int)Math.Ceiling(radius) - 4; // bestimmt Abstand zwischen Text und Delete-Button
        Size = new Size(panelWidth, panelHeight);
        var centerX = Width - radius - 1.5f;  // minus 2.0f Feinjustierung
        var centerY = (Height / 2.0f) - 0.5f; // minus 0.5f Feinjustierung
        var rectX = (int)Math.Round(centerX - (deleteButtonLogicalSize / 2.0f));  // Bounding-Box (16x16) um den Mittelpunkt herum berechnen
        var rectY = (int)Math.Round(centerY - (deleteButtonLogicalSize / 2.0f));
        deleteRect = new Rectangle(rectX, rectY, deleteButtonLogicalSize, deleteButtonLogicalSize);
    }

    private static GraphicsPath GetTagPath(float w, float h)
    {
        w -= 1; // Anpassung der Breite
        h -= 1; // Anpassung der Höhe
        if (w <= 0 || h <= 0) { return new GraphicsPath(); }
        var r = h / 2.0f; // Radius ist die halbe Höhe
        var path = new GraphicsPath();
        path.AddLine(0, 0, w - r, 0);  // obere Linie
        path.AddArc(w - h, 0, h, h, 270, 180);  // rechter Halbkreis-Bogen
        path.AddLine(w - r, h, 0, h);  // Untere Linie
        path.CloseFigure();  // linke Linie entsteht durch das Schließen
        return path;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        float w = ClientRectangle.Width;
        float h = ClientRectangle.Height;
        if (w <= 1 || h <= 1) { return; }
        using var path = GetTagPath(w, h);
        using (var brush = new SolidBrush(Color.Honeydew)) { g.FillPath(brush, path); }  // Hintergrund
        var textY = ((int)(h - TextRenderer.MeasureText(Text, Font).Height) / 2) - 1;
        TextRenderer.DrawText(g, Text, Font, new Point(Padding.Left, textY), ForeColor);
        if (isHoveringPanel)
        {
            if (isHoveringDelete)
            {
                using var hoverBrush = new SolidBrush(Color.LightPink);
                g.FillEllipse(hoverBrush, deleteRect);
            }
            var imgX = deleteRect.X + ((deleteRect.Width - deleteImage.Width) / 2);
            var imgY = deleteRect.Y + ((deleteRect.Height - deleteImage.Height) / 2);
            g.DrawImage(deleteImage, imgX, imgY);
        }
        else
        {
            var eyeletRadius = 2f;
            var centerX = deleteRect.X + deleteRect.Width / 2f;
            var centerY = deleteRect.Y + deleteRect.Height / 2f;
            var eyeletRect = new RectangleF(centerX - eyeletRadius, centerY - eyeletRadius, eyeletRadius * 2f, eyeletRadius * 2f);
            g.FillEllipse(Brushes.White, eyeletRect);
            g.DrawEllipse(Pens.Gray, eyeletRect); // Pens.Black ist statisch und muss nicht disposed (via 'using') werden.
        }
        using var pen = new Pen(SystemColors.ControlDarkDark, 1);  // Rand HotTrack
        g.DrawPath(pen, path);
    }

    private static bool IsPointInCircle(Point mousePos, Rectangle circleBounds)
    {
        var cX = circleBounds.X + circleBounds.Width / 2.0f;
        var cY = circleBounds.Y + circleBounds.Height / 2.0f;
        var r = circleBounds.Width / 2.0f; // Radius (8)
        var dX = mousePos.X - cX;
        var dY = mousePos.Y - cY;
        return (dX * dX + dY * dY) <= (r * r);
    }

    protected override void OnMouseEnter(EventArgs e)
    {
        base.OnMouseEnter(e);
        isHoveringPanel = true;
        Invalidate();
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        isHoveringPanel = false;
        isHoveringDelete = false;
        Invalidate();
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (!isHoveringPanel) { return; }
        var overButton = IsPointInCircle(e.Location, deleteRect);
        if (overButton != isHoveringDelete)
        {
            isHoveringDelete = overButton;
            Invalidate(); // Neu zeichnen (für Hover-Effekt)
        }
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (isHoveringPanel && IsPointInCircle(e.Location, deleteRect)) { DeleteClick?.Invoke(this, EventArgs.Empty); }
    }
}

