namespace Adressen.cls;

internal static class FormStateManager
{
    public static void SetInnerMargins(this TextBoxBase textBox, int left, int right)
    {
        NativeMethods.SendMessage(textBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_RIGHTMARGIN, right << 16);
        NativeMethods.SendMessage(textBox.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN, left);
    }

    public static void SetPlaceholder(this TextBoxBase control, string text) => _ = NativeMethods.SendMessage(control.Handle, NativeMethods.EM_SETCUEBANNER, 0, text);

    public static void RestoreWindowBounds(Form form, WindowPlacement? placement, bool isMaximized = false)
    {
        if (isMaximized)
        {
            form.WindowState = FormWindowState.Maximized;
            return;
        }
        if (placement == null) { return; }
        form.StartPosition = FormStartPosition.Manual;
        form.WindowState = FormWindowState.Normal;
        var targetRect = new Rectangle(placement.X, placement.Y, placement.Width, placement.Height);
        var screen = Screen.FromRectangle(targetRect);  // Screen.FromRectangle ist robuster als FromPoint, da es prüft, wo der größte Teil des Fensters liegt.
        var workArea = screen.WorkingArea;
        var width = Math.Max(targetRect.Width, form.MinimumSize.Width);  // nicht größer als Bildschirm, aber nicht kleiner als MinimumSize
        var height = Math.Max(targetRect.Height, form.MinimumSize.Height);
        width = Math.Min(width, workArea.Width);
        height = Math.Min(height, workArea.Height);
        targetRect.Width = width;
        targetRect.Height = height;
        if (targetRect.Right > workArea.Right) { targetRect.X = workArea.Right - targetRect.Width; }
        if (targetRect.Left < workArea.Left) { targetRect.X = workArea.Left; }
        if (targetRect.Bottom > workArea.Bottom) { targetRect.Y = workArea.Bottom - targetRect.Height; }
        if (targetRect.Top < workArea.Top) { targetRect.Y = workArea.Top; }
        form.DesktopBounds = targetRect;
    }

    internal static bool RowIsVisible(DataGridView dgv, DataGridViewRow row)
    {
        if (dgv.FirstDisplayedCell == null) { return false; }
        var firstVisibleRowIndex = dgv.FirstDisplayedCell.RowIndex;
        var lastVisibleRowIndex = firstVisibleRowIndex + dgv.DisplayedRowCount(false) - 1;
        return row.Index >= firstVisibleRowIndex && row.Index <= lastVisibleRowIndex;
    }

}
