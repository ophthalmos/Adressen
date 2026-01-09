using System.Data;

namespace Adressen.frm;

public partial class FrmGroupFilter : Form
{
    private readonly Dictionary<string, (CheckBox Include, CheckBox Exclude)> _groupControls = [];
    public List<string> IncludedGroups { get; private set; } = [];
    public List<string> ExcludedGroups { get; private set; } = [];

    public FrmGroupFilter(SortedSet<string> groupList)
    {
        InitializeComponent();
        tableLayoutPanel.SuspendLayout();
        tableLayoutPanel.RowCount = 0; // Bestehende Zeilen entfernen
        tableLayoutPanel.RowStyles.Clear();
        //tableLayoutPanel.Controls.Add(new Label { Text = "Einschluss", Dock = DockStyle.Fill }, 1, 0);
        //tableLayoutPanel.Controls.Add(new Label { Text = "Ausschluss", Dock = DockStyle.Fill }, 2, 0);

        foreach (var groupName in groupList.OrderBy(g => g))
        {
            tableLayoutPanel.RowCount++;
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var text = groupName == "★" ? $"{groupName} (Favoriten)" : groupName;
            var groupLabel = new Label { Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            var includeCheckBox = new CheckBoxNoFocus { Text = "", Dock = DockStyle.Fill, CheckAlign = ContentAlignment.MiddleCenter };
            var excludeCheckBox = new CheckBoxNoFocus { Text = "", Dock = DockStyle.Fill, CheckAlign = ContentAlignment.MiddleCenter };
            includeCheckBox.CheckedChanged += (s, e) =>  // Checkboxen dürfen nicht gleichzeitig aktiviert sein
            {
                if (includeCheckBox.Checked) { excludeCheckBox.Checked = false; }
            };
            excludeCheckBox.CheckedChanged += (s, e) =>  // Checkboxen dürfen nicht gleichzeitig aktiviert sein 
            {
                if (excludeCheckBox.Checked) { includeCheckBox.Checked = false; }
            };
            var newRowIndex = tableLayoutPanel.RowCount - 1;
            tableLayoutPanel.Controls.Add(groupLabel, 0, newRowIndex);
            tableLayoutPanel.Controls.Add(includeCheckBox, 1, newRowIndex);
            tableLayoutPanel.Controls.Add(excludeCheckBox, 2, newRowIndex);
            _groupControls.Add(groupName, (includeCheckBox, excludeCheckBox)); // Steuerlemente für spätere Abfrage im Dictionary speichern 
        }
        tableLayoutPanel.RowCount++;
        _ = tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // machte alle Zeilen davor so klein wie nötig (AutoSize), füllt den Rest des Platzes auf
        tableLayoutPanel.ResumeLayout();
    }

    private void ButtonFilter_Click(object sender, EventArgs e)
    {
        IncludedGroups = [.. _groupControls.Where(kvp => kvp.Value.Include.Checked).Select(kvp => kvp.Key)]; // die ausgewählten Gruppen in …
        ExcludedGroups = [.. _groupControls.Where(kvp => kvp.Value.Exclude.Checked).Select(kvp => kvp.Key)]; // die öffentlichen Listen füllen
    }

    private void FrmGroupFilter_Load(object sender, EventArgs e)
    {
        if (panelParent.VerticalScroll.Visible)  // Neuer Right-Wert
        {
            labelHeader.Padding = new Padding(labelHeader.Padding.Left, labelHeader.Padding.Top, SystemInformation.VerticalScrollBarWidth - 2, labelHeader.Padding.Bottom);
            labelHeader.Text = "Einschluss  Ausschluss";
        }
    }

    private void ButtonAll_Click(object sender, EventArgs e)
    {
        buttonAll.Tag ??= 1;
        if (buttonAll.Tag is int foo && foo == 0)
        {
            foreach (var (Include, Exclude) in _groupControls.Values)
            {
                Include.Checked = false;
                Exclude.Checked = false;
            }
            buttonAll.Tag = 1;
        }
        else if (buttonAll.Tag is int bar && bar == 1)
        {
            foreach (var (Include, Exclude) in _groupControls.Values)
            {
                Include.Checked = true;
                Exclude.Checked = false;
            }
            buttonAll.Tag = 2;
        }
        else
        {
            foreach (var (Include, Exclude) in _groupControls.Values)
            {
                Include.Checked = false;
                Exclude.Checked = true;
            }
            buttonAll.Tag = 0;
        }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.F9)
        {
            buttonAll.Focus();
            ButtonAll_Click(buttonAll, EventArgs.Empty);
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

}

public class CheckBoxNoFocus : CheckBox  // Verhindert das Anzeigen des Fokusrands um die Checkbox (obwohl kein Text vorhanden ist) 
{
    protected override bool ShowFocusCues => false;
}
