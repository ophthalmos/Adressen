using System.Text.RegularExpressions;
using Adressen.cls;

namespace Adressen; // WICHTIG: Muss exakt wie im Designer heißen

public partial class FrmCopyScheme : Form
{
    private readonly AppSettings _settings;
    private readonly Dictionary<string, string> _addBookDict;
    private readonly Font _tabFont = new("Segoe UI", 10.0f, FontStyle.Bold, GraphicsUnit.Point);

    // Helper-Property für Zugriff auf die Textbox des aktuellen Tabs
    private TextBox CurrentPatternBox
    {
        get
        {
            if (tabControl.SelectedTab == tabPage6) { return tbPattern6; }
            if (tabControl.SelectedTab == tabPage5) { return tbPattern5; }
            if (tabControl.SelectedTab == tabPage4) { return tbPattern4; }
            if (tabControl.SelectedTab == tabPage3) { return tbPattern3; }
            if (tabControl.SelectedTab == tabPage2) { return tbPattern2; }
            return tbPattern1;
        }
    }

    // Neuer Konstruktor: Nimmt nur Settings & Dictionary
    internal FrmCopyScheme(AppSettings settings, Dictionary<string, string> addressDict)
    {
        InitializeComponent();
        _settings = settings;
        _addBookDict = addressDict;

        // --- Farben anwenden ---
        panelLeft.BackColor = _settings.ColorScheme switch
        {
            "blue" => SystemColors.GradientInactiveCaption,
            "pale" => SystemColors.ControlLightLight,
            "dark" => SystemColors.ControlDark,
            _ => SystemColors.Control
        };

        foreach (TabPage tabPage in tabControl.TabPages)
        {
            tabPage.BackColor = _settings.ColorScheme switch
            {
                "blue" => SystemColors.InactiveBorder,
                "pale" => SystemColors.ControlLightLight,
                "dark" => SystemColors.AppWorkspace,
                _ => SystemColors.ButtonFace
            };
        }

        // --- Combobox füllen ---
        cbxFields.Items.AddRange([.. _addBookDict.Keys]);
        if (cbxFields.Items.Count > 0) { cbxFields.SelectedIndex = 0; }
        Utils.AdjustComboBoxDropDownWidth(cbxFields);
        // --- Textboxen aus Settings füllen ---
        tbPattern1.Lines = _settings.CopyPattern1 ?? [];
        tbPattern2.Lines = _settings.CopyPattern2 ?? [];
        tbPattern3.Lines = _settings.CopyPattern3 ?? [];
        tbPattern4.Lines = _settings.CopyPattern4 ?? [];
        tbPattern5.Lines = _settings.CopyPattern5 ?? [];
        tbPattern6.Lines = _settings.CopyPattern6 ?? [];

        // --- Letzten aktiven Tab setzen ---
        if (_settings.CopyPatternIndex >= 0 && _settings.CopyPatternIndex < tabControl.TabCount)
        {
            tabControl.SelectedIndex = _settings.CopyPatternIndex;
        }

        // --- Tooltips initialisieren ---
        UpdateAllTooltips();
    }

    private void FrmCopyScheme_Load(object sender, EventArgs e) => UpdateCurrentTabInfo();

    private void FrmCopyScheme_Shown(object sender, EventArgs e)
    {
        tbPattern1.Select(tbPattern1.Text.Length, 0);
        btnCopy.Focus();
        Utils.MoveCursorToControl(btnCopy);
    }

    // WICHTIG: Dies ist der Button "Text in Zwischenablage kopieren"
    // Da er im Designer DialogResult = OK hat, dient er gleichzeitig als "Speichern & Schließen".
    private void BtnCopy_Click(object sender, EventArgs e)
    {
        Utils.SetClipboardText(tbResult.Text.Trim());

        // 2. Änderungen zurück in das Settings-Objekt schreiben
        _settings.CopyPattern1 = tbPattern1.Lines;
        _settings.CopyPattern2 = tbPattern2.Lines;
        _settings.CopyPattern3 = tbPattern3.Lines;
        _settings.CopyPattern4 = tbPattern4.Lines;
        _settings.CopyPattern5 = tbPattern5.Lines;
        _settings.CopyPattern6 = tbPattern6.Lines;
        _settings.CopyPatternIndex = tabControl.SelectedIndex;

        // Form schließt sich automatisch wegen btnCopy.DialogResult = OK
    }

    private void BtnInsert_Click(object sender, EventArgs e)
    {
        var tbPattern = CurrentPatternBox; // Nutzt den Helper oben
        var textToInsert = $"[{cbxFields.Text}]"; // var textToInsert = cbxFields.Text;
        var cursorPosition = tbPattern.SelectionStart;

        // Logik: Leerzeichen automatisch einfügen
        while (cursorPosition < tbPattern.Text.Length && !char.IsWhiteSpace(tbPattern.Text[cursorPosition])) { cursorPosition++; }
        if (cursorPosition > 0 && !char.IsWhiteSpace(tbPattern.Text[cursorPosition - 1])) { textToInsert = " " + textToInsert; }

        tbPattern.Text = tbPattern.Text.Insert(cursorPosition, textToInsert);
        tbPattern.SelectionStart = cursorPosition + textToInsert.Length;
        tbPattern.Focus();
    }

    // Wird für alle tbPatternX TextChanged Events aufgerufen
    private void TbPattern_TextChanged(object sender, EventArgs e) => UpdateCurrentTabInfo();

    private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabControl.Visible && tabControl.Focused) { UpdateCurrentTabInfo(); }
    }

    // Aktualisiert Tooltip und Scrollbars für den aktuellen Tab
    private void UpdateCurrentTabInfo()
    {
        if (!tabControl.Visible || tabControl.SelectedTab == null) { return; }

        var tbPattern = CurrentPatternBox;

        if (string.IsNullOrEmpty(tbPattern.Text))
        {
            tbResult.Clear();
            tabControl.SelectedTab.ToolTipText = string.Empty;
        }
        else
        {
            tbResult.Lines = UsePattern(tbPattern.Lines);
            tabControl.SelectedTab.ToolTipText = tbResult.Text.Trim();

            var textSize = TextRenderer.MeasureText(tbPattern.Text, tbPattern.Font,
                new Size(tbPattern.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue),
                TextFormatFlags.LeftAndRightPadding | TextFormatFlags.TextBoxControl);

            tbPattern.ScrollBars = textSize.Height > tbPattern.Height ? ScrollBars.Vertical : ScrollBars.None;
        }
    }

    // Aktualisiert alle Tooltips (beim Start nötig)
    private void UpdateAllTooltips()
    {
        tabPage1.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern1.Lines)).Trim();
        tabPage2.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern2.Lines)).Trim();
        tabPage3.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern3.Lines)).Trim();
        tabPage4.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern4.Lines)).Trim();
        tabPage5.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern5.Lines)).Trim();
        tabPage6.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern6.Lines)).Trim();
    }

    private string[] UsePattern(string[] pattern)
    {
        if (pattern == null)
        {
            return [];
        }

        var result = new string[pattern.Length];

        for (var i = 0; i < pattern.Length; i++)
        {
            var line = pattern[i];

            // Das Pattern @"\[([^\]]+)\]" sucht nach:
            // \[       -> einer öffnenden eckigen Klammer
            // ([^\]]+) -> (Gruppe 1) beliebig vielen Zeichen, die KEINE schließende eckige Klammer sind
            // \]       -> einer schließenden eckigen Klammer
            result[i] = Regex.Replace(line, @"\[([^\]]+)\]", match =>
            {
                // match.Groups[1].Value enthält den reinen Key OHNE die Klammern (z.B. "Vorname")
                var key = match.Groups[1].Value;

                if (_addBookDict.TryGetValue(key, out var value)) { return value; }

                // Wenn der Key nicht im Dictionary existiert (z.B. wenn der Nutzer manuell [Hallo] tippt), 
                // bleibt der Originaltext inklusive Klammern unangetastet stehen.
                return match.Value;
            });
        }

        return result;
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape) { Close(); return true; }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
    {
        if (sender is not TabControl tabControlSender) { return; }

        var g = e.Graphics;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

        var tabPage = tabControlSender.TabPages[e.Index];
        var tabBounds = tabControlSender.GetTabRect(e.Index);

        if (e.State == DrawItemState.Selected) { g.FillRectangle(Brushes.Gray, e.Bounds); }
        else { e.DrawBackground(); }

        var textColor = e.State == DrawItemState.Selected ? Color.White : tabControlSender.ForeColor;
        using var textBrush = new SolidBrush(textColor);

        using var stringFlags = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center,
            FormatFlags = StringFormatFlags.NoWrap,
            Trimming = StringTrimming.EllipsisCharacter
        };
        tabBounds.Inflate(-2, -2);
        g.DrawString(tabPage.Text, _tabFont, textBrush, tabBounds, stringFlags);
    }

    private void StatusStrip_Paint(object sender, PaintEventArgs e)
    {
        if (sender is not StatusStrip strip) { return; }
        var splitX = panelLeft.Width;  // Die Grenze ist exakt die Breite des linken Panels
        using var brush = new SolidBrush(panelLeft.BackColor);  // Den linken Teil mit der Farbe von panelLeft übermalen
        e.Graphics.FillRectangle(brush, 0, 0, splitX, strip.Height);  // Wir füllen das Rechteck von (0,0) bis (splitX, Höhe)
    }

    private void LblGoogleSearch_Click(object sender, EventArgs e)
    {
        var searchText = tbResult.Text.Replace("\r", " ").Replace("\n", " ").Trim();
        if (string.IsNullOrWhiteSpace(searchText)) { return; }
        var query = Uri.EscapeDataString(searchText);
        var url = $"https://www.google.com/search?q={query}";
        Utils.StartLink(Handle, url);
        BtnCopy_Click(sender, e); 
        DialogResult = DialogResult.OK;
        //Close();
    }
}
