using Adressen.cls;
using System.Data;
using System.Text.RegularExpressions;

namespace Adressen;

public partial class FrmCopyScheme : Form
{
    public string[] GetPattern1() => tbPattern1.Lines;
    public string[] GetPattern2() => tbPattern2.Lines;
    public string[] GetPattern3() => tbPattern3.Lines;
    public string[] GetPattern4() => tbPattern4.Lines;
    public string[] GetPattern5() => tbPattern5.Lines;
    public string[] GetPattern6() => tbPattern6.Lines;
    public int PatternIndex => tabControl.SelectedIndex;  // nullbasiert

    private readonly Dictionary<string, string> addBookDict = [];

    public FrmCopyScheme(string colorScheme, Dictionary<string, string> addressDict, int patternIndex, string[] pattern1, string[] pattern2, string[] pattern3, string[] pattern4, string[] pattern5, string[] pattern6)
    {
        InitializeComponent();
        panel.BackColor = colorScheme switch { "blue" => SystemColors.GradientInactiveCaption, "pale" => SystemColors.ControlLightLight, "dark" => SystemColors.ControlDark, _ => SystemColors.Control, };
        foreach (TabPage tabPage in tabControl.TabPages)
        {
            tabPage.BackColor = colorScheme switch { "blue" => SystemColors.InactiveBorder, "pale" => SystemColors.ControlLightLight, "dark" => SystemColors.AppWorkspace, _ => SystemColors.ButtonFace, };
        }
        foreach (var key in addressDict.Keys) { cbxFields.Items.Add(key); }
        addBookDict = addressDict;
        cbxFields.SelectedIndex = 0;
        tabControl.SelectedIndex = patternIndex;
        tbPattern1.Lines = pattern1 ?? [];
        tbPattern2.Lines = pattern2 ?? [];
        tbPattern3.Lines = pattern3 ?? [];
        tbPattern4.Lines = pattern4 ?? [];
        tbPattern5.Lines = pattern5 ?? [];
        tbPattern6.Lines = pattern6 ?? [];
        tabPage1.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern1.Lines)).Trim();
        tabPage2.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern2.Lines)).Trim();
        tabPage3.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern3.Lines)).Trim();
        tabPage4.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern4.Lines)).Trim();
        tabPage5.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern5.Lines)).Trim();
        tabPage6.ToolTipText = string.Join(Environment.NewLine, UsePattern(tbPattern6.Lines)).Trim();
    }

    private void FrmCopyScheme_Load(object sender, EventArgs e)
    {
        TbPattern_TextChanged(sender, EventArgs.Empty);
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape) { Close(); return true; }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void BtnInsert_Click(object sender, EventArgs e)
    {
        var tbPattern = tabControl.SelectedTab == tabPage6 ? tbPattern6 : tabControl.SelectedTab == tabPage5 ? tbPattern5 : tabControl.SelectedTab == tabPage4 ? tbPattern4
            : tabControl.SelectedTab == tabPage3 ? tbPattern3 : tabControl.SelectedTab == tabPage2 ? tbPattern2 : tbPattern1; // kein Using verwenden!
        var textToInsert = cbxFields.Text;
        var cursorPosition = tbPattern.SelectionStart;
        while (cursorPosition < tbPattern.Text.Length && !char.IsWhiteSpace(tbPattern.Text[cursorPosition])) { cursorPosition++; }
        if (cursorPosition > 0 && !char.IsWhiteSpace(tbPattern.Text[cursorPosition - 1])) { textToInsert = " " + textToInsert; }
        tbPattern.Text = tbPattern.Text.Insert(cursorPosition, textToInsert);  // String an der Cursorposition einfügen
        tbPattern.SelectionStart = cursorPosition + textToInsert.Length;  // Cursorposition aktualisieren
    }

    private void TbPattern_TextChanged(object sender, EventArgs e)
    {
        if (!tabControl.Visible || tabControl.SelectedTab == null) { return; } // Sicherstellen, dass tabControl und SelectedTab nicht null sind    
        var tbPattern = tabControl.SelectedTab == tabPage6 ? tbPattern6 : tabControl.SelectedTab == tabPage5 ? tbPattern5 : tabControl.SelectedTab == tabPage4 ? tbPattern4
            : tabControl.SelectedTab == tabPage3 ? tbPattern3 : tabControl.SelectedTab == tabPage2 ? tbPattern2 : tbPattern1; // kein Using verwenden!
        if (string.IsNullOrEmpty(tbPattern.Text))
        {
            tbResult.Clear();
            tabControl.SelectedTab.ToolTipText = string.Empty;
        }
        else
        {
            tbResult.Lines = UsePattern(tbPattern.Lines); // result;
            tabControl.SelectedTab.ToolTipText = tbResult.Text.Trim();
            tbPattern.ScrollBars = TextRenderer.MeasureText(tbPattern.Text, tbPattern.Font,
                new Size(tbPattern.Width - SystemInformation.VerticalScrollBarWidth, int.MaxValue),
                TextFormatFlags.LeftAndRightPadding | TextFormatFlags.TextBoxControl).Height > tbPattern.Height ? ScrollBars.Vertical : ScrollBars.None;
        }
    }

    private string[] UsePattern(string[] pattern)
    {
        if (pattern == null) { return []; }
        var result = new string[pattern.Length];
        for (var i = 0; i < pattern.Length; i++)
        {
            var line = pattern[i];
            var words = Regex.Matches(line, @"\b\w+\b").Cast<Match>().Select(m => addBookDict.TryGetValue(m.Value, out var value) ? value : null);
            result[i] = string.Join(" ", words).Trim();
        }
        return result;
    }

    private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabControl.Visible && tabControl.Focused) { TbPattern_TextChanged(sender, EventArgs.Empty); }
    }

    private void BtnCopy_Click(object sender, EventArgs e) => Utils.SetClipboardText(tbResult.Text.Trim());

    private void FrmCopyScheme_Shown(object sender, EventArgs e)
    {
        tbPattern1.Select(tbPattern1.Text.Length, 0);
        btnCopy.Focus();
    }

    private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
    {
        if (sender is TabControl tabControlSender)
        {
            using var g = e.Graphics;
            Brush textBrush;
            e.DrawBackground();
            var tabPage = tabControlSender.TabPages[e.Index]; // Get the item from the collection.
            if (e.State == DrawItemState.Selected)
            {
                textBrush = new SolidBrush(Color.White);
                g.FillRectangle(Brushes.Gray, e.Bounds);
            }
            else
            {
                textBrush = new SolidBrush(e.ForeColor);
                e.DrawBackground();
            }
            using var tabFont = new Font("Segoe UI", 10.0f, FontStyle.Bold, GraphicsUnit.Point);
            using var stringFlags = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center }; // Draw string. Center the text.
            var tabBounds = tabControlSender.GetTabRect(e.Index); // Get the real bounds for the tab rectangle.
            tabBounds.Inflate(-2, -2); // Inflate the rectangle to make it smaller.
            g.DrawString(tabPage.Text, tabFont, textBrush, tabBounds, new StringFormat(stringFlags));
        }
    }

}
