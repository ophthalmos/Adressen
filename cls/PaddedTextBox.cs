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

    protected override void WndProc(ref Message m)
    {
        // 1. Dpppelklick separat behandeln
        if (m.Msg == NativeMethods.WM_LBUTTONDBLCLK)
        {
            HandleCustomDoubleClick(m.LParam);
            return;
        }
        // 2. Basis IMMER ausführen
        base.WndProc(ref m);
        // 3. Nach dem Setzen des Fonts zwingen wir unsere Margins auf
        if (m.Msg == NativeMethods.WM_SETFONT) { ApplyMargins(); }
        // 4. Custom Placeholder zeichnen NACHDEM das Control fertig gezeichnet ist
        if (m.Msg == NativeMethods.WM_PAINT && TextLength == 0 && !Focused && !string.IsNullOrEmpty(_customPlaceholder))
        {
            if (ClientSize.Width <= 0 || ClientSize.Height <= 0) { return; }
            using var gScreen = Graphics.FromHwnd(Handle);
            var editRect = new NativeMethods.RECT();
            NativeMethods.SendMessage(Handle, NativeMethods.EM_GETRECT, IntPtr.Zero, ref editRect);
            var targetLeft = Math.Max(editRect.Left, LeftInnerMargin + 1);
            var targetRight = Math.Min(editRect.Right, ClientSize.Width - RightInnerMargin);
            var rect = new Rectangle(targetLeft, editRect.Top, Math.Max(0, targetRight - targetLeft), editRect.Bottom - editRect.Top);

            gScreen.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            using var brush = new SolidBrush(Color.LightGray); // Ideal wäre SystemColors.GrayText

            using var format = new StringFormat(StringFormat.GenericTypographic)
            {
                LineAlignment = Multiline ? StringAlignment.Near : StringAlignment.Center,
                FormatFlags = StringFormatFlags.NoWrap,
                Trimming = StringTrimming.EllipsisCharacter
            };

            rect.X += 1;
            gScreen.DrawString(_customPlaceholder, Font, brush, rect, format);
            // KEIN RETURN HIER! Die Event-Kette ist beendet und Windows räumt sauber auf.
        }
    }

    private void ApplyMargins()
    {
        if (!IsHandleCreated) { return; }

        if (Multiline)
        {
            var rect = new NativeMethods.RECT
            {
                Left = LeftInnerMargin,
                Top = 2,
                Right = ClientSize.Width - RightInnerMargin,
                Bottom = ClientSize.Height
            };
            NativeMethods.SendMessage(Handle, NativeMethods.EM_SETRECT, IntPtr.Zero, ref rect);
        }
        else
        {
            var lParam = (_rightMargin << 16) | (_leftMargin & 0xFFFF);
            var wParam = NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN;
            NativeMethods.SendMessage(Handle, NativeMethods.EM_SETMARGINS, wParam, lParam);
        }
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

    protected override void WndProc(ref Message m)
    {
        // 1. Doppelklick separat behandeln
        if (m.Msg == NativeMethods.WM_LBUTTONDBLCLK)
        {
            HandleCustomDoubleClick(m.LParam);
            return;
        }

        // 2. Basis IMMER ausführen, damit das Control (und die Maske) gezeichnet wird
        base.WndProc(ref m);

        // 3. Nach dem Setzen des Fonts zwingen wir unsere Margins auf
        if (m.Msg == NativeMethods.WM_SETFONT) { ApplyMargins(); }

        // 4. Custom Placeholder zeichnen NACHDEM das Control fertig gezeichnet ist
        if (m.Msg == NativeMethods.WM_PAINT && !Focused && !string.IsNullOrEmpty(_customPlaceholder))
        {
            // Bei MaskedTextBox prüfen wir, ob tatsächliche Eingaben vorliegen, nicht nur die Maske
            var isEmpty = MaskedTextProvider == null ? TextLength == 0 : MaskedTextProvider.AssignedEditPositionCount == 0;

            if (isEmpty)
            {
                if (ClientSize.Width <= 0 || ClientSize.Height <= 0) { return; }

                using var gScreen = Graphics.FromHwnd(Handle);
                var editRect = new NativeMethods.RECT();
                NativeMethods.SendMessage(Handle, NativeMethods.EM_GETRECT, IntPtr.Zero, ref editRect);

                // WICHTIG: Die von der Basisklasse gezeichnete leere Maske mit der Hintergrundfarbe "radieren"
                using var bgBrush = new SolidBrush(BackColor);
                gScreen.FillRectangle(bgBrush, editRect.Left, editRect.Top, editRect.Right - editRect.Left, editRect.Bottom - editRect.Top);

                var targetLeft = Math.Max(editRect.Left, LeftInnerMargin + 1);
                var targetRight = Math.Min(editRect.Right, ClientSize.Width - RightInnerMargin);
                var rect = new Rectangle(targetLeft, editRect.Top, Math.Max(0, targetRight - targetLeft), editRect.Bottom - editRect.Top);
                if (rect.Width <= 0 || rect.Height <= 0) { return; }
                gScreen.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                using var brush = new SolidBrush(Color.LightGray); // SystemColors.GrayText ist ebenfalls eine gute Wahl

                using var format = new StringFormat(StringFormat.GenericTypographic)
                {
                    LineAlignment = Multiline ? StringAlignment.Near : StringAlignment.Center,
                    FormatFlags = StringFormatFlags.NoWrap,
                    Trimming = StringTrimming.EllipsisCharacter
                };

                rect.X += 1;
                gScreen.DrawString(_customPlaceholder, Font, brush, rect, format);
            }
        }
    }

    private void ApplyMargins()
    {
        if (!IsHandleCreated) { return; }
        var lParam = (_rightMargin << 16) | (_leftMargin & 0xFFFF);
        var wParam = NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN;
        NativeMethods.SendMessage(Handle, NativeMethods.EM_SETMARGINS, wParam, lParam);
    }

    private void HandleCustomDoubleClick(IntPtr lParam)
    {
        // Exakt den sichtbaren Text (inkl. Platzhalter und Trennzeichen) abrufen
        var text = MaskedTextProvider != null ? MaskedTextProvider.ToDisplayString() : Text;
        if (string.IsNullOrEmpty(text)) { return; }
        var x = (int)(lParam.ToInt64() & 0xFFFF);
        var y = (int)((lParam.ToInt64() >> 16) & 0xFFFF);
        var clickLocation = new Point(x, y);
        var index = GetCharIndexFromPosition(clickLocation);

        // Prüfen, ob der Index im Bereich des Display-Strings liegt
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
