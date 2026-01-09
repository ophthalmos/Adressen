namespace Adressen;

public partial class FrmBirthdays : Form
{
    public int SelectionIndex => listView.SelectedIndices.Count > 0 ? listView.SelectedIndices[0] : -1;
    public int BirthdayRemindLimit => (int)beforeNumUpDown.Value; // >= 0 ? (int)numericUpDown.Value : -1;
    public int BirthdayRemindAfter => (int)afterNumUpDown.Value; // >= 0 ? (int)numericUpDown.Value : -1;

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool BirthdayAutoShow
    {
        get => chkBxBirthdayAutoShow.Checked;
        set => chkBxBirthdayAutoShow.Checked = value;
    }

    private readonly List<int> _birthdayTodayList = [];
    private readonly Image partyHat = Properties.Resources.FavoriteStar16;
    private readonly bool isLocal;
    private readonly int _initialIndex = -1;

    public FrmBirthdays(string colorScheme, List<(DateOnly Datum, string Name, int Alter, int Tage, string Id)> geburtstage, int reminderBefore, int reminderAfter, bool localAdr) // , string colorSheme
    {
        InitializeComponent();
        isLocal = localAdr;
        BackColor = colorScheme switch
        {
            "blue" => SystemColors.GradientInactiveCaption,
            "pale" => SystemColors.ControlLightLight,
            "dark" => SystemColors.ControlDark,
            _ => SystemColors.Control,
        };
        beforeNumUpDown.Value = reminderBefore <= beforeNumUpDown.Maximum && reminderBefore >= beforeNumUpDown.Minimum ? reminderBefore : beforeNumUpDown.Value;
        afterNumUpDown.Value = reminderAfter <= afterNumUpDown.Maximum && reminderAfter >= afterNumUpDown.Minimum ? reminderAfter : afterNumUpDown.Value;
        var nextBirthdayIndex = -1; // -1 bedeutet "noch nicht gefunden"
        var minDays = int.MaxValue; // Startwert so hoch wie möglich setzen
        var index = 0;
        foreach (var info in geburtstage)
        {
            var item = new ListViewItem(info.Datum.ToShortDateString());
            item.SubItems.Add(info.Name);
            item.SubItems.Add(info.Alter.ToString());
            item.SubItems.Add(info.Tage.ToString());
            if (info.Tage >= 0 && info.Tage < minDays)
            {
                minDays = info.Tage; // Kleinste Tagesdifferenz speichern
                nextBirthdayIndex = index; // Zugehörigen Index speichern
            }

            if (info.Tage == 0) { _birthdayTodayList.Add(index); }
            listView.Items.Add(item);
            index++;
        }
        if (nextBirthdayIndex != -1)
        {
            _initialIndex = nextBirthdayIndex;
            AcceptButton = btnShowAddress;
        }
        else if (listView.Items.Count > 0)
        {
            AcceptButton = btnShowAddress;
            listView.Items[0].Selected = true;
            listView.Items[0].Focused = true;
            listView.EnsureVisible(0);
        }
    }

    private void ListView_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (listView.SelectedIndices.Count > 0)
        {
            btnShowAddress.Enabled = true;
            AcceptButton = btnShowAddress;
        }
        else
        {
            btnShowAddress.Enabled = false;
            AcceptButton = btnCancel;
        }
    }

    private void ListView_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e) => e.DrawDefault = true;

    private void ListView_DrawItem(object sender, DrawListViewItemEventArgs e) => e.DrawDefault = true;

    private void ListView_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
    {
        var isSelected = e.Item != null && e.Item.Selected;
        if (e.SubItem != null)
        {
            e.Graphics.FillRectangle(new SolidBrush(isSelected && isLocal ? Color.FromArgb(176, 125, 71) : isSelected ? SystemColors.Highlight : e.SubItem.BackColor), e.Bounds);
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, listView.Font, e.Bounds, isSelected ? Color.White : e.SubItem.ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
        }
        if (e.ColumnIndex == 1)
        {
            if (_birthdayTodayList.Contains(e.ItemIndex))
            {
                var bildX = e.Bounds.Right - partyHat.Width - 4; // am rechten Rand der Zelle
                var bildY = e.Bounds.Top + (e.Bounds.Height - partyHat.Height) / 2;
                e.Graphics.DrawImage(partyHat, bildX, bildY);
            }
        }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape) { Close(); return true; }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void ListView_MouseDoubleClick(object sender, MouseEventArgs e) => DialogResult = DialogResult.OK;

    private void FrmBirthdays_Shown(object sender, EventArgs e)
    {
        BringToFront();
        Activate();
        if (_initialIndex >= 0 && _initialIndex < listView.Items.Count)
        {
            listView.Focus();
            var item = listView.Items[_initialIndex];
            item.Selected = true;
            item.Focused = true;
            listView.FocusedItem = item;
            listView.EnsureVisible(_initialIndex);
        }
        else { listView.Focus(); }
    }

    private void ListView_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Space && listView.SelectedIndices.Count > 0) { DialogResult = DialogResult.OK; }
    }
}


