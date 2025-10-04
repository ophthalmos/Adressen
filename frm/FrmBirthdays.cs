namespace Adressen;
public partial class FrmBirthdays : Form
{
    public int SelectionIndex => listView.SelectedIndices.Count > 0 ? listView.SelectedIndices[0] : -1;
    public int BirthdayRemindLimit => (int)numericUpDown.Value; // >= 0 ? (int)numericUpDown.Value : -1;

    private readonly List<int> _birthdayTodayList = [];
    private readonly Image partyHat = Properties.Resources.FavoriteStar16;
    private readonly bool isLocal;

    public FrmBirthdays(string colorScheme, List<(DateTime Datum, string Name, int Alter, int Tage, string Id)> geburtstage, int reminderDays, bool localAdr) // , string colorSheme
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
        numericUpDown.Value = reminderDays <= numericUpDown.Maximum && reminderDays >= numericUpDown.Minimum ? reminderDays : numericUpDown.Value;
        var index = 0;
        foreach (var info in geburtstage)
        {
            var item = new ListViewItem(info.Datum.ToShortDateString());
            item.SubItems.Add(info.Name);
            item.SubItems.Add(info.Alter.ToString());
            item.SubItems.Add(info.Tage.ToString());
            if (info.Tage == 0) { _birthdayTodayList.Add(index); }
            listView.Items.Add(item);
            index++;
        }
        if (listView.Items.Count > 0)
        {
            listView.Items[0].Selected = true;
            listView.Items[0].Focused = true;
            listView.EnsureVisible(0); // Scrollt zur ersten Zeile
            AcceptButton = btnShowAddress; // Setzt die Schaltfläche, die bei Enter gedrückt wird
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
    }

    private void ListView_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Space && listView.SelectedIndices.Count > 0) { DialogResult = DialogResult.OK; }
    }
}


