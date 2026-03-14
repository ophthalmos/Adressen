using System.ComponentModel;

namespace Adressen.cls;

public class PaddedTextBox : TextBox
{
    private int _leftMargin = AppSettings.TextBoxPadding;
    private int _rightMargin = AppSettings.TextBoxPadding;
    private string _customPlaceholder = string.Empty;

    [Category("Appearance")]
    [Description("Gibt den inneren linken Abstand des Textes an.")]
    public int LeftInnerMargin
    {
        get => _leftMargin;
        set
        {
            if (_leftMargin != value)
            {
                _leftMargin = value;
                ApplyMargins();
            }
        }
    }

    [Category("Appearance")]
    [Description("Gibt den inneren rechten Abstand des Textes an.")]
    public int RightInnerMargin
    {
        get => _rightMargin;
        set
        {
            if (_rightMargin != value)
            {
                _rightMargin = value;
                ApplyMargins();
            }
        }
    }

    private bool ShouldSerializeLeftInnerMargin() => _leftMargin != AppSettings.TextBoxPadding;
    private bool ShouldSerializeRightInnerMargin() => _rightMargin != AppSettings.TextBoxPadding;

    private void ResetRightInnerMargin() => RightInnerMargin = AppSettings.TextBoxPadding;
    private void ResetLeftInnerMargin() => LeftInnerMargin = AppSettings.TextBoxPadding;

    [Category("Appearance")]
    [Description("Der Text, der angezeigt wird, wenn das Steuerelement leer ist.")]
    [DefaultValue("")]
    [Localizable(true)]
    public new string PlaceholderText
    {
        get => _customPlaceholder;
        set
        {
            if (_customPlaceholder != value)
            {
                _customPlaceholder = value ?? string.Empty;
                base.PlaceholderText = string.Empty;
                Invalidate();
            }
        }
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        ApplyMargins();
        if (Multiline && ScrollBars == ScrollBars.Vertical)
        {
            AppendText(" ");
        }
    }

    protected override void OnEnter(EventArgs e)
    {
        base.OnEnter(e);
        ApplyMargins();
    }

    protected override void OnLeave(EventArgs e)
    {
        base.OnLeave(e);
        Invalidate();
    }

    protected override void OnTextChanged(EventArgs e)
    {
        base.OnTextChanged(e);
        ApplyMargins();
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        ApplyMargins();
    }

    protected override void OnFontChanged(EventArgs e)
    {
        base.OnFontChanged(e);
        ApplyMargins();
    }

    private void ApplyMargins()
    {
        if (!IsHandleCreated) { return; }
        if (Multiline)  // EM_SETMARGINS funktioniert bei mehrzeiligen Textboxen nicht korrekt, da es die Scrollbar mitverschiebt.
        {
            var rect = new NativeMethods.RECT
            {
                Left = LeftInnerMargin,
                Top = 2, // Ein Pixel Luft nach oben sieht in WinForms immer etwas sauberer aus
                Right = ClientSize.Width - RightInnerMargin,
                Bottom = ClientSize.Height
            };
            NativeMethods.SendMessage(Handle, NativeMethods.EM_SETRECT, IntPtr.Zero, ref rect);
        }
        else  // Für einzeilige Textboxen funktioniert EM_SETMARGINS fehlerfrei
        {
            var lParam = (_rightMargin << 16) | (_leftMargin & 0xFFFF);
            var wParam = NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN;
            NativeMethods.SendMessage(Handle, NativeMethods.EM_SETMARGINS, wParam, lParam);
        }
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == NativeMethods.WM_LBUTTONDBLCLK)
        {
            HandleCustomDoubleClick(m.LParam);
            return;
        }

        if (m.Msg == NativeMethods.WM_PAINT && TextLength == 0 && !Focused && !string.IsNullOrEmpty(_customPlaceholder))
        {
            // 1. ZUERST das native Control regulär zeichnen lassen!
            // Dadurch wird der (ausgegraute) Scrollbalken von Windows initialisiert und gerendert.
            base.WndProc(ref m);

            // Wenn das Control minimiert oder noch nicht sichtbar ist, ignorieren
            if (ClientSize.Width <= 0 || ClientSize.Height <= 0) { return; }

            // 2. Platzhalter über den Textbereich zeichnen
            using var gScreen = Graphics.FromHwnd(Handle);
            var editRect = new NativeMethods.RECT();
            NativeMethods.SendMessage(Handle, NativeMethods.EM_GETRECT, IntPtr.Zero, ref editRect);

            var targetLeft = Math.Max(editRect.Left, LeftInnerMargin + 1);
            var targetRight = Math.Min(editRect.Right, ClientSize.Width - RightInnerMargin);
            var rect = new Rectangle(targetLeft, editRect.Top, Math.Max(0, targetRight - targetLeft), editRect.Bottom - editRect.Top);

            gScreen.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            using var brush = new SolidBrush(Color.LightGray);

            using var format = new StringFormat(StringFormat.GenericTypographic)
            {
                LineAlignment = Multiline ? StringAlignment.Near : StringAlignment.Center,
                FormatFlags = StringFormatFlags.NoWrap,
                Trimming = StringTrimming.EllipsisCharacter
            };

            rect.X += 1;
            gScreen.DrawString(_customPlaceholder, Font, brush, rect, format);
            return;
        }

        // Für alle anderen Nachrichten ganz normales Verhalten
        base.WndProc(ref m);
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
    private int _leftMargin = AppSettings.TextBoxPadding;
    private int _rightMargin = AppSettings.TextBoxPadding;

    [Category("Appearance")]
    [Description("Gibt den inneren linken Abstand des Textes an.")]
    public int LeftInnerMargin
    {
        get => _leftMargin;
        set
        {
            if (_leftMargin != value)
            {
                _leftMargin = value;
                ApplyMargins();
            }
        }
    }

    [Category("Appearance")]
    [Description("Gibt den inneren rechten Abstand des Textes an.")]
    public int RightInnerMargin
    {
        get => _rightMargin;
        set
        {
            if (_rightMargin != value)
            {
                _rightMargin = value;
                ApplyMargins();
            }
        }
    }

    private bool ShouldSerializeLeftInnerMargin() => _leftMargin != AppSettings.TextBoxPadding;
    private void ResetLeftInnerMargin() => LeftInnerMargin = AppSettings.TextBoxPadding;
    private bool ShouldSerializeRightInnerMargin() => _rightMargin != AppSettings.TextBoxPadding;
    private void ResetRightInnerMargin() => RightInnerMargin = AppSettings.TextBoxPadding;

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        ApplyMargins();
    }

    protected override void OnFontChanged(EventArgs e)
    {
        base.OnFontChanged(e);
        ApplyMargins();
    }

    private void ApplyMargins()
    {
        if (IsHandleCreated)
        {
            var lParam = (_rightMargin << 16) | (_leftMargin & 0xFFFF);
            var wParam = NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN;
            NativeMethods.SendMessage(Handle, NativeMethods.EM_SETMARGINS, wParam, lParam);
        }
    }
}