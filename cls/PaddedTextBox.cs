using System.ComponentModel;

namespace Adressen.cls;

public class PaddedTextBox : TextBox
{
    private int _cachedSpaceWidth = -1; // -1: muss berechnet werden; ApplyEditControlsFont() (siehe Settings-Dialog) aktualisiert die Schriftart und invalidiert damit diesen Cache

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        ApplyPadding();  // Native Placeholder – kein Flackern, kein WndProc-Eingriff
    }

    private void ApplyPadding()
    {
        var p = GetSpaceWidth();

        if (Multiline)
        {
            // 1. EM_SETMARGINS anwenden, damit die Links-/Rechts-Ausrichtung exakt mit den einzeiligen TextBoxen übereinstimmt
            _ = NativeMethods.SendMessage(Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (p << 16) | (p & 0xFFFF));

            // 2. Das durch die Margins intern angepasste Rechteck abrufen
            var rect = new NativeMethods.RECT();
            _ = NativeMethods.SendMessage(Handle, NativeMethods.EM_GETRECT, IntPtr.Zero, ref rect);

            // 3. Nur die vertikale Ausrichtung korrigieren, die perfekt ausgerichteten horizontalen Werte beibehalten
            rect.Top = 2;
            rect.Bottom = ClientSize.Height;

            _ = NativeMethods.SendMessage(Handle, NativeMethods.EM_SETRECT, IntPtr.Zero, ref rect);
        }
        else
        {
            _ = NativeMethods.SendMessage(Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (p << 16) | (p & 0xFFFF));
        }
    }

    private int GetSpaceWidth()
    {
        if (_cachedSpaceWidth == -1) { _cachedSpaceWidth = TextRenderer.MeasureText(AppSettings.TextBoxPaddingChar.ToString(), Font, Size.Empty, TextFormatFlags.NoPadding).Width; }
        return _cachedSpaceWidth;
    }

    protected override void OnFontChanged(EventArgs e)
    {
        base.OnFontChanged(e);
        _cachedSpaceWidth = -1;  // Cache invalidieren
        ApplyPadding();  // Layout neu anwenden, da die Schriftgröße sich geändert hat
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == NativeMethods.WM_LBUTTONDBLCLK)
        {
            HandleCustomDoubleClick(m.LParam);
            return;
        }
        base.WndProc(ref m);
        if (m.Msg == NativeMethods.WM_SETFONT) { ApplyPadding(); }
    }

    private void HandleCustomDoubleClick(IntPtr lParam)
    {
        if (TextLength == 0) { return; }

        var x = (int)(lParam.ToInt64() & 0xFFFF);
        var y = (int)((lParam.ToInt64() >> 16) & 0xFFFF);
        var clickLocation = new Point(x, y);
        var index = GetCharIndexFromPosition(clickLocation);
        if (index < 0 || index >= TextLength) { return; }
        var text = Text;

        static int GetCharClass(char c)
        {
            if (char.IsWhiteSpace(c) || char.IsControl(c)) { return 0; }
            if (char.IsLetterOrDigit(c) || c == '_') { return 1; }
            return 2;
        }

        var targetClass = GetCharClass(text[index]);
        var start = index;
        var end = index;
        while (start > 0 && GetCharClass(text[start - 1]) == targetClass) { start--; }
        while (end < TextLength - 1 && GetCharClass(text[end + 1]) == targetClass) { end++; }
        var length = end - start + 1;
        if (length > 0) { Select(start, length); }
    }
}

public class PaddedMaskedTextBox : MaskedTextBox
{
    private int _innerMargin = -1;  // -1 bedeutet: "Nutze die dynamisch berechnete Breite des AppSettings.TextBoxPadding-Zeichens"
    private int _cachedPaddingWidth = -1;
    private string _customPlaceholder = string.Empty;

    [Category("Appearance")]
    [Description("Gibt den inneren Abstand (links und rechts) des Textes an.")]
    public int InnerMargin
    {
        get
        {
            // Wenn kein benutzerdefinierter Rand gesetzt wurde, berechne ihn anhand des Zeichens
            if (_innerMargin == -1) { return GetPaddingWidth(); }
            return _innerMargin;
        }
        set
        {
            if (_innerMargin != value)
            {
                _innerMargin = value;
                ApplyMargins();
            }
        }
    }

    //// --- Alias-Eigenschaften ---
    //[Browsable(false)]
    //public int LeftInnerMargin
    //{
    //    get => InnerMargin;
    //    set => InnerMargin = value;
    //}

    //[Browsable(false)]
    //public int RightInnerMargin
    //{
    //    get => InnerMargin;
    //    set => InnerMargin = value;
    //}

    private bool ShouldSerializeInnerMargin() => _innerMargin != -1;
    private void ResetInnerMargin() => InnerMargin = -1;
    //private static bool ShouldSerializeLeftInnerMargin() => false;
    //private static bool ShouldSerializeRightInnerMargin() => false;

    [Category("Appearance")]
    [Description("Der Text, der angezeigt wird, wenn das Steuerelement leer ist.")]
    [DefaultValue("")]
    [Localizable(true)]
    public string PlaceholderText
    {
        get => _customPlaceholder;
        set
        {
            if (_customPlaceholder != value)
            {
                _customPlaceholder = value ?? string.Empty;
                Invalidate();
            }
        }
    }

    private int GetPaddingWidth()
    {
        if (_cachedPaddingWidth == -1)
        {
            var text = AppSettings.TextBoxPaddingChar.ToString();
            _cachedPaddingWidth = TextRenderer.MeasureText(text, Font, Size.Empty, TextFormatFlags.NoPadding).Width;
        }
        return _cachedPaddingWidth;
    }

    protected override void OnFontChanged(EventArgs e)
    {
        base.OnFontChanged(e);
        _cachedPaddingWidth = -1; // Cache bei Font-Änderung leeren
        ApplyMargins();
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        ApplyMargins();
    }

    private void ApplyMargins()
    {
        if (!IsHandleCreated) { return; }
        var margin = InnerMargin;
        var lParam = (margin << 16) | (margin & 0xFFFF);
        var wParam = NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN;
        _ = NativeMethods.SendMessage(Handle, NativeMethods.EM_SETMARGINS, wParam, lParam);
    }

    protected override void WndProc(ref Message m)
    {
        // 1. Abfangen (Early Return)
        if (m.Msg == NativeMethods.WM_PASTE)
        {
            if (Mask != null && Mask.Replace(@"\", "") == "00.00.0000" && Clipboard.ContainsText())
            {
                var clipboardText = Clipboard.GetText().Trim();
                if (DateOnly.TryParse(clipboardText, out var parsedDate))
                {
                    Text = parsedDate.ToString("dd.MM.yyyy");
                    return;
                }
            }
        }
        if (m.Msg == NativeMethods.WM_LBUTTONDBLCLK)
        {
            HandleCustomDoubleClick(m.LParam);
            return;
        }

        // 2. delegieren (Base-Processing)
        base.WndProc(ref m);
        
        // 3. nachbearbeiten (Post-Processing)
        if (m.Msg == NativeMethods.WM_SETFONT) { ApplyMargins(); }
        if (m.Msg == NativeMethods.WM_PAINT) { DrawPlaceholderIfNeeded(); }
    }


    private void DrawPlaceholderIfNeeded()
    {
        if (Focused || string.IsNullOrEmpty(_customPlaceholder)) { return; }

        var isEmpty = MaskedTextProvider?.AssignedEditPositionCount == 0;
        if (!isEmpty) { return; }
        if (!IsHandleCreated || IsDisposed || Disposing || ClientSize.Width <= 0 || ClientSize.Height <= 0) { return; }

        try
        {
            using var g = Graphics.FromHwnd(Handle);

            var editRect = new NativeMethods.RECT();
            NativeMethods.SendMessage(Handle, NativeMethods.EM_GETRECT, IntPtr.Zero, ref editRect);

            // Masken-Zeichen weglöschen
            using var bgBrush = new SolidBrush(BackColor);
            g.FillRectangle(bgBrush, editRect.Left, editRect.Top, editRect.Right - editRect.Left, editRect.Bottom - editRect.Top);

            var targetLeft = Math.Max(editRect.Left, InnerMargin + 1);
            var targetRight = Math.Min(editRect.Right, ClientSize.Width - InnerMargin);
            var rect = new Rectangle(targetLeft, editRect.Top, Math.Max(0, targetRight - targetLeft), editRect.Bottom - editRect.Top);

            if (rect.Width <= 0 || rect.Height <= 0) { return; }

            var flags = TextFormatFlags.NoPadding       // GenericTypographic → kein Glyph-Overhang-Padding
                      | TextFormatFlags.SingleLine       // StringFormatFlags.NoWrap
                      | TextFormatFlags.EndEllipsis      // StringTrimming.EllipsisCharacter
                      | (Multiline
                            ? TextFormatFlags.Top            // StringAlignment.Near
                            : TextFormatFlags.VerticalCenter); // StringAlignment.Center

            TextRenderer.DrawText(g, _customPlaceholder, Font, rect, SystemColors.GrayText, flags);
        }
        catch (ArgumentException) { }
    }

    private void HandleCustomDoubleClick(IntPtr lParam)
    {
        var text = MaskedTextProvider != null ? MaskedTextProvider.ToDisplayString() : Text;
        if (string.IsNullOrEmpty(text)) { return; }

        var x = (int)(lParam.ToInt64() & 0xFFFF);
        var y = (int)((lParam.ToInt64() >> 16) & 0xFFFF);
        var clickLocation = new Point(x, y);
        var index = GetCharIndexFromPosition(clickLocation);

        if (index < 0 || index >= text.Length) { return; }

        static int GetCharClass(char c)
        {
            if (char.IsWhiteSpace(c) || char.IsControl(c)) { return 0; }
            if (char.IsLetterOrDigit(c) || c == '_') { return 1; }
            return 2;
        }

        var targetClass = GetCharClass(text[index]);
        var start = index;
        var end = index;

        while (start > 0 && GetCharClass(text[start - 1]) == targetClass) { start--; }
        while (end < text.Length - 1 && GetCharClass(text[end + 1]) == targetClass) { end++; }

        var length = end - start + 1;
        if (length > 0) { Select(start, length); }
    }
}