using System.ComponentModel;
using System.Globalization;

namespace Adressen.cls;

public partial class YearSlider : UserControl
{
    public event EventHandler? YearChanged;

    /// <summary>Feuert bei jeder Textänderung in der TextBox, auch bei partieller Eingabe.</summary>
    public event EventHandler? RawTextChanged;

    public YearSlider()
    {
        InitializeComponent();
        SetStyle(ControlStyles.Selectable, true);
        TabStop = true;
        _year = null;  // Default: keine Jahreszahl anzeigen
        UpdateTextBoxFromYear();

        btnPrev.Click += (_, __) => DecrementYear();
        btnNext.Click += (_, __) => IncrementYear();
        txtYear.KeyPress += TxtYear_KeyPress;       // nur Ziffern erlauben
        txtYear.KeyDown += TxtYear_KeyDown;         // Enter/Escape
        txtYear.Leave += TxtYear_Leave;             // beim Verlassen validieren
    }

    private int? _year = null;
    private int _minYear = 1900;
    private int _maxYear = 2100;

    [Browsable(true)]
    [DefaultValue(1900)]
    public int MinYear
    {
        get => _minYear;
        set
        {
            _minYear = value;
            if (_year.HasValue && _year.Value < value) { Year = value; }
        }
    }

    [Browsable(true)]
    [DefaultValue(2100)]
    public int MaxYear
    {
        get => _maxYear;
        set
        {
            _maxYear = value;
            if (_year.HasValue && _year.Value > value) { Year = value; }
        }
    }

    [Browsable(true)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public int? Year
    {
        get => _year;
        set
        {
            var newValue = value.HasValue ? Math.Max(MinYear, Math.Min(MaxYear, value.Value)) : (int?)null;
            if (newValue == _year) { return; }
            _year = newValue;
            UpdateTextBoxFromYear();
            YearChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private void UpdateTextBoxFromYear()
    {
        var pos = txtYear.SelectionStart;  // Cursorposition vor dem Setzen merken
        txtYear.Text = _year.HasValue ? _year.Value.ToString(CultureInfo.InvariantCulture) : string.Empty;
        if (txtYear.Text.Length > 0) { txtYear.SelectionStart = Math.Min(pos, txtYear.Text.Length); }
    }

    /// <summary>Gibt den aktuellen TextBox-Inhalt zurück, ohne Validierung oder Clamp.
    /// Ermöglicht partielle Jahressuche (z.B. "19" findet alle Jahre mit "19").</summary>
    public string RawText => txtYear.Text.Trim();

    /// <summary>Setzt die TextBox direkt auf einen beliebigen Text, ohne Year-Validierung.
    /// Für ApplyCriteria: Wiederherstellen von Suchkriterien inkl. partieller Eingaben.</summary>
    public void SetSearchText(string text) => txtYear.Text = text;

    /// <summary>
    /// Optionaler Startwert für <see cref="IncrementYear"/> wenn noch kein Jahr gesetzt ist.
    /// Wird von außen gesetzt (z.B. yearSlider2.DefaultYear = yearSlider1.Year).
    /// <see cref="DecrementYear"/> ignoriert diesen Wert und nutzt immer das aktuelle Jahr.
    /// </summary>
    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public int? DefaultYear { get; set; }

    /// <summary>Hintergrundfarbe ausschließlich der inneren TextBox.</summary>
    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Color TextBackColor
    {
        get => txtYear.BackColor;
        set { if (Enabled) { txtYear.BackColor = value; } }
    }

    // Hilfsmethoden für konsistente Inkremente/Dekremente
    private void IncrementYear()
    {
        if (!_year.HasValue) { Year = Math.Max(MinYear, Math.Min(MaxYear, DefaultYear ?? DateTime.Now.Year)); }
        else { Year = Math.Min(MaxYear, _year.Value + 1); }
    }

    private void DecrementYear()
    {
        if (!_year.HasValue) { Year = Math.Max(MinYear, Math.Min(MaxYear, DateTime.Now.Year)); }
        else { Year = Math.Max(MinYear, _year.Value - 1); }
    }

    // Public helper
    public void ClearYear() => Year = null;
    public void SetYear(int y) => Year = y;

    // TextBox: nur Ziffern, Backspace erlauben
    private void TxtYear_KeyPress(object? sender, KeyPressEventArgs e)
    {
        if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) { e.Handled = true; }
    }

    // Enter = übernehmen, Escape = verwerfen (zurück auf vorherigen Wert)
    private readonly string _textBeforeEdit = string.Empty;

    private void TxtYear_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            CommitTextBox();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
        else if (e.KeyCode == Keys.Escape)
        {
            txtYear.Text = _textBeforeEdit;  // Rücksetzen auf vorherigen Text
            txtYear.SelectAll();
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Up)
        {
            IncrementYear();
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Down)
        {
            DecrementYear();
            e.Handled = true;
        }
    }

    private void TxtYear_TextChanged(object sender, EventArgs e)
    {
        var txt = txtYear.Text.Trim();
        if (string.IsNullOrEmpty(txt))
        {
            Year = null;
        }
        else if (txt.Length == 4 && int.TryParse(txt, out var parsed))
        {
            if (parsed < MinYear) { parsed = MinYear; }
            if (parsed > MaxYear) { parsed = MaxYear; }
            Year = parsed; // ruft YearChanged auf, falls Wert sich ändert
        }
        // partielle Eingabe (1–3 Ziffern): Year bleibt null, Text bleibt erhalten
        RawTextChanged?.Invoke(this, EventArgs.Empty);
    }

    private void TxtYear_Leave(object? sender, EventArgs e) => CommitTextBox();

    private void CommitTextBox()
    {
        var txt = txtYear.Text.Trim();
        if (string.IsNullOrEmpty(txt))
        {
            Year = null;  // leeres Feld = keine Auswahl
            return;
        }
        if (int.TryParse(txt, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        {
            if (txt.Length >= 4)
            {
                // vollständige Jahresangabe: Clamp und übernehmen
                parsed = Math.Max(MinYear, Math.Min(MaxYear, parsed));
                Year = parsed;
                UpdateTextBoxFromYear();
            }
            // partielle Eingabe (1–3 Ziffern): Text und Year=null bleiben erhalten
        }
        else { UpdateTextBoxFromYear(); }  // ungültig (nicht-numerisch): zurücksetzen
    }

    // ProcessCmdKey wird vor ProcessDialogKey aufgerufen und gilt auch für Child-Controls
    // (btnPrev/btnNext), bei denen OnKeyDown des UserControls nicht feuert.
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (!txtYear.Focused)
        {
            var step = keyData.HasFlag(Keys.Shift) ? 10 : keyData.HasFlag(Keys.Control) ? 100 : 1;
            switch (keyData & Keys.KeyCode)
            {
                case Keys.Up:
                case Keys.Right:
                    Year = _year.HasValue
                        ? Math.Min(MaxYear, _year.Value + step)
                        : Math.Max(MinYear, Math.Min(MaxYear, DefaultYear ?? DateTime.Now.Year));
                    return true;
                case Keys.Down:
                case Keys.Left:
                    Year = _year.HasValue
                        ? Math.Max(MinYear, _year.Value - step)
                        : Math.Max(MinYear, Math.Min(MaxYear, DateTime.Now.Year));
                    return true;
                case Keys.Delete:
                    ClearYear();
                    return true;
            }
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        Focus();
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        base.OnMouseWheel(e);
        if (e.Delta > 0) { IncrementYear(); }
        else if (e.Delta < 0) { DecrementYear(); }
    }

    protected override void OnEnabledChanged(EventArgs e)
    {
        base.OnEnabledChanged(e);
        btnPrev.Enabled = Enabled;
        btnNext.Enabled = Enabled;
        txtYear.ReadOnly = !Enabled;
        txtYear.BackColor = Enabled ? SystemColors.Window : SystemColors.Control;
        Invalidate();
    }
}
