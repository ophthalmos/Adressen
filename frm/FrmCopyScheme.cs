using System.Text.RegularExpressions;
using Adressen.cls;

namespace Adressen; 

public partial class FrmCopyScheme : Form
{
    private readonly AppSettings _settings;
    private readonly Dictionary<string, string> _addBookDict;
    private readonly Font _tabFont = new("Segoe UI", 10.0f, FontStyle.Bold, GraphicsUnit.Point);


    private TextBox CurrentPatternBox  // Helper-Property für Zugriff auf die Textbox des aktuellen Tabs
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

    internal FrmCopyScheme(AppSettings settings, Dictionary<string, string> addressDict)
    {
        InitializeComponent();
        _settings = settings;
        _addBookDict = addressDict;

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

        cbxFields.Items.AddRange([.. _addBookDict.Keys]);
        if (cbxFields.Items.Count > 0) { cbxFields.SelectedIndex = 0; }
        Utils.AdjustComboBoxDropDownWidth(cbxFields);
        tbPattern1.Lines = _settings.CopyPattern1 ?? [];
        tbPattern2.Lines = _settings.CopyPattern2 ?? [];
        tbPattern3.Lines = _settings.CopyPattern3 ?? [];
        tbPattern4.Lines = _settings.CopyPattern4 ?? [];
        tbPattern5.Lines = _settings.CopyPattern5 ?? [];
        tbPattern6.Lines = _settings.CopyPattern6 ?? [];
        if (_settings.CopyPatternIndex >= 0 && _settings.CopyPatternIndex < tabControl.TabCount)        {            tabControl.SelectedIndex = _settings.CopyPatternIndex;        }
        UpdateAllTooltips();
    }

    private void FrmCopyScheme_Load(object sender, EventArgs e) => UpdateCurrentTabInfo();

    private void FrmCopyScheme_Shown(object sender, EventArgs e)
    {
        tbPattern1.Select(tbPattern1.Text.Length, 0);
        btnCopy.Focus();
        Utils.MoveCursorToControl(btnCopy);
    }

  
    private void BtnCopy_Click(object sender, EventArgs e)  // Button "Text in Zwischenablage kopieren" = "Speichern & Schließen"
    {
        Utils.SetClipboardText(tbResult.Text.Trim());
        _settings.CopyPattern1 = tbPattern1.Lines;
        _settings.CopyPattern2 = tbPattern2.Lines;
        _settings.CopyPattern3 = tbPattern3.Lines;
        _settings.CopyPattern4 = tbPattern4.Lines;
        _settings.CopyPattern5 = tbPattern5.Lines;
        _settings.CopyPattern6 = tbPattern6.Lines;
        _settings.CopyPatternIndex = tabControl.SelectedIndex;
    }  // Form schließt sich automatisch wegen btnCopy.DialogResult = OK

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

    private void TbPattern_TextChanged(object sender, EventArgs e) => UpdateCurrentTabInfo();

    private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabControl.Visible && tabControl.Focused) { UpdateCurrentTabInfo(); }
    }

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
            var tooltipText = tbResult.Text.Trim().Replace("\t", "    ");  // WinForms-Tooltips brechen oft bei Tab-Zeichen ab.
            tabControl.SelectedTab.ToolTipText = tooltipText;
            var textSize = TextRenderer.MeasureText(tbPattern.Text, tbPattern.Font,
                new Size(tbPattern.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue),
                TextFormatFlags.LeftAndRightPadding | TextFormatFlags.TextBoxControl);
            if (textSize.Height > tbPattern.Height) { tbPattern.ScrollBars = ScrollBars.Vertical; }
            else { tbPattern.ScrollBars = ScrollBars.None; }
        }
    }

    private void UpdateAllTooltips()
    {
        string GetCleanTooltip(string[] lines)  // Lokale Hilfsfunktion für saubere Tooltip-Texte ohne Tabulatoren
        {
            var text = string.Join(Environment.NewLine, UsePattern(lines)).Trim();
            return text.Replace("\t", "    ");
        }
        tabPage1.ToolTipText = GetCleanTooltip(tbPattern1.Lines);
        tabPage2.ToolTipText = GetCleanTooltip(tbPattern2.Lines);
        tabPage3.ToolTipText = GetCleanTooltip(tbPattern3.Lines);
        tabPage4.ToolTipText = GetCleanTooltip(tbPattern4.Lines);
        tabPage5.ToolTipText = GetCleanTooltip(tbPattern5.Lines);
        tabPage6.ToolTipText = GetCleanTooltip(tbPattern6.Lines);
    }

    private string[] UsePattern(string[] pattern)
    {
        if (pattern == null) { return []; }
        var result = new string[pattern.Length];
        for (var i = 0; i < pattern.Length; i++)
        {
            var line = pattern[i];
            line = Regex.Replace(line, @"\[([^\]]+)\]", match =>
            {
                var key = match.Groups[1].Value;
                if (_addBookDict.TryGetValue(key, out var value)) { return value ?? string.Empty; }  // Falls der Wert null ist, leeren String zurückgeben, damit keine "null"-Strings im Text landen.
                return match.Value;
            });
            line = Regex.Replace(line, @"(,\s*){2,}", ", ");  // Mehrfache Kommas (auch getrennt durch Leerzeichen) zu einem einzigen Komma zusammenfassen
            line = Regex.Replace(line, @" {2,}", " ");  // Mehrfache aufeinanderfolgende normale Leerzeichen zu einem einzigen zusammenfassen
            line = line.Trim(' ', ',');  // Führende und nachfolgende Leerzeichen sowie Kommas entfernen (z.B. wenn das erste Feld leer war)
            result[i] = line;
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
