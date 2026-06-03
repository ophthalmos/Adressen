using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Net;

namespace Adressen.cls;

public class IPv4AddressControl : UserControl
{
    private readonly TextBox[] _oct = new TextBox[4];

    public IPv4AddressControl()
    {
        SuspendLayout();
        BorderStyle = BorderStyle.Fixed3D;
        BackColor = SystemColors.Window;
        Height = 23;
        Width = 160;

        var x = 4;

        for (var i = 0; i < 4; i++)
        {
            var idx = i; // Closure-Variable, damit die Event-Handler den richtigen Index verwenden können
            _oct[i] = new TextBox
            {
                Left = x,
                Top = 2,
                Width = 30,
                Margin = new Padding(0),
                MaxLength = 3,
                BorderStyle = BorderStyle.None,
                TextAlign = HorizontalAlignment.Center,
                BackColor = SystemColors.Window,
            };
            _oct[i].KeyPress += (_, e) => OnOctKeyPress(idx, e);
            _oct[i].KeyDown += (_, e) => OnOctKeyDown(idx, e);
            _oct[i].TextChanged += (_, _) => OnOctTextChanged(idx);
            _oct[i].Enter += (_, _) => _oct[idx].SelectAll();
            _oct[i].Leave += (_, _) => ClampOctet(idx);
            Controls.Add(_oct[i]);
            x += 30;  // X-Koordinate um die Breite der Textbox verschieben

            if (i < 3)  // Punkte zwischen den Oktetten hinzufügen
            {
                Controls.Add(new Label
                {
                    Text = ".",
                    Left = x,
                    Top = 2,
                    Padding = new Padding(0),
                    Margin = new Padding(0),
                    AutoSize = true,
                    BackColor = Color.Transparent
                });
                x += 10;  // X-Koordinate für den nächsten Durchlauf um den Platz für den Punkt verschieben
            }
        }

        ResumeLayout();
    }

    // ── Tastatur-Handling ────────────────────────────────────────────────────
    private void OnOctKeyPress(int idx, KeyPressEventArgs e)
    {
        if (e.KeyChar is '.' or ' ')
        {
            if (idx < 3) { _oct[idx + 1].Focus(); }

            e.Handled = true;
        }
        else if (!char.IsDigit(e.KeyChar) && e.KeyChar != '\b') { e.Handled = true; }
    }

    private void OnOctKeyDown(int idx, KeyEventArgs e)
    {
        var tb = _oct[idx];

        if (e.KeyCode == Keys.Back && tb.Text.Length == 0 && idx > 0)
        {
            _oct[idx - 1].Focus();
            _oct[idx - 1].SelectAll();
            e.Handled = true;
        }
        // Pfeil Links: Wenn Cursor ganz links ist und es nicht das erste Feld ist
        else if (e.KeyCode == Keys.Left && tb.SelectionStart == 0 && idx > 0)
        {
            _oct[idx - 1].Focus();
            _oct[idx - 1].SelectionStart = _oct[idx - 1].Text.Length;
            e.Handled = true;
        }
        // Pfeil Rechts: Wenn Cursor ganz rechts ist und es nicht das letzte Feld ist
        else if (e.KeyCode == Keys.Right && tb.SelectionStart == tb.Text.Length && idx < 3)
        {
            _oct[idx + 1].Focus();
            _oct[idx + 1].SelectionStart = 0;
            e.Handled = true;
        }
    }

    private void OnOctTextChanged(int idx)
    {
        if (_oct[idx].Text.Length == 3 && idx < 3 && int.TryParse(_oct[idx].Text, out var v) && v <= 255)
        {
            _oct[idx + 1].Focus();
            _oct[idx + 1].SelectAll();
        }
    }

    private void ClampOctet(int idx)
    {
        if (string.IsNullOrEmpty(_oct[idx].Text)) { return; }
        _oct[idx].Text = int.TryParse(_oct[idx].Text, out var v) ? Math.Clamp(v, 0, 255).ToString() : "0";
    }

    // ── Öffentliche API ──────────────────────────────────────────────────────

    [Category("Daten")]
    [DefaultValue("0.0.0.0")]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
    public string Address
    {
        get => string.Join(".", _oct.Select(t => t.Text.Length > 0 ? t.Text : "0"));
        set
        {
            var parts = (value ?? "0.0.0.0").Split('.');

            for (var i = 0; i < 4; i++)
            {
                if (i < parts.Length) { _oct[i].Text = parts[i]; }
                else { _oct[i].Text = "0"; }
            }
        }
    }

    public bool TryGetIPAddress([NotNullWhen(true)] out IPAddress? result)
    {
        return IPAddress.TryParse(Address, out result);
    }

    [Browsable(false)]
    public bool IsValid => IPAddress.TryParse(Address, out _);
}