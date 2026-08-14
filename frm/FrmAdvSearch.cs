using Adressen.cls;
using Adressen.Properties;

namespace Adressen.frm;

public partial class FrmAdvSearch : Form
{
    public bool RefineSearch => cbRefineSearch.Checked;

    public FrmAdvSearch(string colorScheme = "", bool isFilterActive = false)
    {
        InitializeComponent();
        SetColorScheme(colorScheme);
        btnResetDate.Image = Resources.delete12;
        tbVorname.BackColor = Color.LightYellow;
        cbRefineSearch.Enabled = isFilterActive;
        if (isFilterActive) { cbRefineSearch.Checked = true; }

        _ = NativeMethods.SendMessage(tbVorname.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
        _ = NativeMethods.SendMessage(tbNachname.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
        _ = NativeMethods.SendMessage(tbNickname.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
        _ = NativeMethods.SendMessage(tbTitel.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
        _ = NativeMethods.SendMessage(tbAnrede.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
        _ = NativeMethods.SendMessage(tbUnternehmen.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
        _ = NativeMethods.SendMessage(tbStrasse.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
        _ = NativeMethods.SendMessage(tbOrt.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
        _ = NativeMethods.SendMessage(tbPLZvon.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
        _ = NativeMethods.SendMessage(tbPLZbis.Handle, NativeMethods.EM_SETMARGINS, NativeMethods.EC_LEFTMARGIN | NativeMethods.EC_RIGHTMARGIN, (2 << 16) | (2 & 0xFFFF));
    }

    private void SetColorScheme(string colorScheme)
    {
        panel.BackColor = groupBox1.BackColor = groupBox2.BackColor = groupBox3.BackColor = colorScheme switch
        {
            "blue" => SystemColors.InactiveBorder,
            "pale" => SystemColors.ControlLightLight,
            "dark" => SystemColors.Control,
            _ => SystemColors.ButtonFace,
        };
    }

    // ── Öffentliche API ──────────────────────────────────────────────────────

    /// <summary>Liest alle Formularfelder aus und gibt ein fertiges Kriterien-Objekt zurück.</summary>
    public AdvancedSearchCriteria BuildCriteria() => new()
    {
        Vorname = tbVorname.Text.AsNullIfEmpty(),
        Nachname = tbNachname.Text.AsNullIfEmpty(),
        Nickname = tbNickname.Text.AsNullIfEmpty(),
        Praefix = tbTitel.Text.AsNullIfEmpty(),
        Anrede = tbAnrede.Text.AsNullIfEmpty(),
        Unternehmen = tbUnternehmen.Text.AsNullIfEmpty(),
        Strasse = tbStrasse.Text.AsNullIfEmpty(),
        Ort = tbOrt.Text.AsNullIfEmpty(),
        PLZvon = tbPLZvon.Text.AsNullIfEmpty(),
        PLZbis = tbPLZbis.Enabled ? tbPLZbis.Text.AsNullIfEmpty() : null,
        GeburtsjahrVon = yearSlider1.RawText.AsNullIfEmpty(),
        GeburtsjahrBis = yearSlider2.Enabled ? yearSlider2.Year : null,
        Mode = rbContains.Checked ? SearchMode.Contains : rbStartwith.Checked ? SearchMode.StartsWith : SearchMode.Exact,
        Logic = rbAND.Checked ? SearchLogic.And : SearchLogic.Or,
    };

    /// <summary>Befüllt das Formular mit bestehenden Kriterien (für Wiederöffnen mit letzten Werten).</summary>
    public void ApplyCriteria(AdvancedSearchCriteria c)
    {
        tbVorname.Text = c.Vorname ?? string.Empty;
        tbNachname.Text = c.Nachname ?? string.Empty;
        tbNickname.Text = c.Nickname ?? string.Empty;
        tbTitel.Text = c.Praefix ?? string.Empty;
        tbAnrede.Text = c.Anrede ?? string.Empty;
        tbUnternehmen.Text = c.Unternehmen ?? string.Empty;
        tbStrasse.Text = c.Strasse ?? string.Empty;
        tbOrt.Text = c.Ort ?? string.Empty;
        tbPLZvon.Text = c.PLZvon ?? string.Empty;
        tbPLZbis.Text = c.PLZbis ?? string.Empty;
        yearSlider1.SetSearchText(c.GeburtsjahrVon ?? string.Empty);
        yearSlider2.Year = c.GeburtsjahrBis;

        // Enabled-Zustände aus den wiederhergestellten Werten ableiten
        TbPLZvon_TextChanged(tbPLZvon, EventArgs.Empty);
        YearSlider1_RawTextChanged(yearSlider1, EventArgs.Empty);

        (rbContains.Checked, rbStartwith.Checked, rbExact.Checked) = c.Mode switch
        {
            SearchMode.StartsWith => (false, true, false),
            SearchMode.Exact => (false, false, true),
            _ => (true, false, false),
        };
        rbAND.Checked = c.Logic == SearchLogic.And;
        rbOR.Checked = c.Logic == SearchLogic.Or;
    }

    // ── Event-Handler ────────────────────────────────────────────────────────

    private void BtnResetDate_Click(object sender, EventArgs e) => yearSlider1.ClearYear();

    private void BtnReset_Click(object? sender, EventArgs e)
    {
        tbVorname.Clear();
        tbNachname.Clear();
        tbNickname.Clear();
        tbTitel.Clear();
        tbAnrede.Clear();
        tbUnternehmen.Clear();
        tbStrasse.Clear();
        tbOrt.Clear();
        tbPLZvon.Clear();
        tbPLZbis.Clear();
        yearSlider1.ClearYear();   // löst RawTextChanged aus → yearSlider2 wird geleert und disabled
        rbContains.Checked = true;
        rbAND.Checked = true;
        tbVorname.Focus();
    }

    private void YearSlider1_RawTextChanged(object sender, EventArgs e)
    {
        var hasText = !string.IsNullOrWhiteSpace(yearSlider1.RawText);
        if (!hasText) { yearSlider2.ClearYear(); }
        yearSlider2.Enabled = yearSlider1.Year.HasValue;
        btnResetDate.Enabled = hasText;  // auch bei partieller Eingabe löschbar
        yearSlider2.DefaultYear = yearSlider1.Year;
    }

    private void TbPLZvon_TextChanged(object sender, EventArgs e) => tbPLZbis.Enabled = labelBis.Enabled = !string.IsNullOrWhiteSpace(tbPLZvon.Text);

    private async void TextBox_Enter(object sender, EventArgs e)
    {
        if (sender is TextBox tb) { tb.BackColor = Color.LightYellow; }
    }

    private void TextBox_Leave(object sender, EventArgs e)
    {
        if (sender is TextBox tb) { tb.BackColor = Color.White; }
    }

    private void YearText_Enter(object sender, EventArgs e)
    {
        yearSlider1.TextBackColor = yearSlider2.TextBackColor = Color.LightYellow;
    }

    private void YearText_Leave(object sender, EventArgs e)
    {
        yearSlider1.TextBackColor = yearSlider2.TextBackColor = SystemColors.Window;
    }

}
