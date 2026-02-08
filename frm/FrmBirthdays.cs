using Adressen.cls; // Namespace für AppSettings

namespace Adressen;

public partial class FrmBirthdays : Form
{
    // Diese Property bleibt, da sie reiner UI-Status ist (welche Zeile gewählt wurde)
    public int SelectionIndex => listView.SelectedIndices.Count > 0 ? listView.SelectedIndices[0] : -1;

    private readonly AppSettings _settings;
    private readonly List<int> _birthdayTodayList = [];
    private readonly Image partyHat = Properties.Resources.FavoriteStar16; // oder Ihr Bildname
    private readonly bool _isLocal; // Umbenannt in _isLocal (Naming Convention)
    private readonly int _initialIndex = -1;

    // Konstruktor nimmt jetzt AppSettings direkt entgegen
    public FrmBirthdays(AppSettings settings, List<(DateOnly Datum, string Name, int Alter, int Tage, string Id)> geburtstage, bool localAdr)
    {
        InitializeComponent();
        _settings = settings;
        _isLocal = localAdr;

        // --- 1. Design & Farben ---
        ApplyColorScheme();

        // --- 2. Data Binding ---
        InitializeDataBindings();

        // --- 3. Listen-Logik (unverändert, aber sauberer) ---
        var nextBirthdayIndex = -1;
        var minDays = int.MaxValue;

        listView.BeginUpdate(); // Performance-Boost beim Füllen
        for (var i = 0; i < geburtstage.Count; i++)
        {
            var info = geburtstage[i];
            var item = new ListViewItem(info.Datum.ToShortDateString());
            item.SubItems.Add(info.Name);
            item.SubItems.Add(info.Alter.ToString());
            item.SubItems.Add(info.Tage.ToString());

            // Logik für "Nächsten Geburtstag" suchen
            if (info.Tage >= 0 && info.Tage < minDays)
            {
                minDays = info.Tage;
                nextBirthdayIndex = i;
            }

            if (info.Tage == 0) { _birthdayTodayList.Add(i); }
            listView.Items.Add(item);
        }
        listView.EndUpdate();

        // --- 4. Fokus setzen ---
        if (nextBirthdayIndex != -1)
        {
            _initialIndex = nextBirthdayIndex;
            AcceptButton = btnShowAddress;
        }
        else if (listView.Items.Count > 0)
        {
            AcceptButton = btnShowAddress;
            listView.Items[0].Selected = true;
            listView.EnsureVisible(0);
        }
    }

    private void ApplyColorScheme()
    {
        BackColor = _settings.ColorScheme switch
        {
            "blue" => SystemColors.GradientInactiveCaption,
            "pale" => SystemColors.ControlLightLight,
            "dark" => SystemColors.ControlDark,
            _ => SystemColors.Control,
        };
    }

    private void InitializeDataBindings()
    {
        // Numerische Felder direkt an Settings binden
        beforeNumUpDown.DataBindings.Add("Value", _settings, nameof(AppSettings.BirthdayRemindLimit), false, DataSourceUpdateMode.OnPropertyChanged);
        afterNumUpDown.DataBindings.Add("Value", _settings, nameof(AppSettings.BirthdayRemindAfter), false, DataSourceUpdateMode.OnPropertyChanged);

        // CheckBox dynamisch binden:
        // Wenn Adressen-Modus (isLocal) -> Binde an BirthdayAddressShow
        // Wenn Kontakte-Modus (!isLocal) -> Binde an BirthdayContactShow
        var targetProperty = _isLocal ? nameof(AppSettings.BirthdayAddressShow) : nameof(AppSettings.BirthdayContactShow);
        chkBxBirthdayAutoShow.DataBindings.Add("Checked", _settings, targetProperty, false, DataSourceUpdateMode.OnPropertyChanged);
    }

    // --- Event Handler ---

    private void ListView_SelectedIndexChanged(object sender, EventArgs e)
    {
        var hasSelection = listView.SelectedIndices.Count > 0;
        btnShowAddress.Enabled = hasSelection;
        AcceptButton = hasSelection ? btnShowAddress : btnCancel;
    }

    private void ListView_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e) => e.DrawDefault = true;
    private void ListView_DrawItem(object sender, DrawListViewItemEventArgs e) => e.DrawDefault = true;

    private void ListView_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
    {
        // Verhindert Fehler, wenn Item null ist
        if (e.Item == null || e.SubItem == null) { return; }

        var isSelected = e.Item.Selected;

        // Hintergrund zeichnen
        using (var backBrush = new SolidBrush(isSelected && _isLocal ? Color.FromArgb(176, 125, 71) : isSelected ? SystemColors.Highlight : e.SubItem.BackColor))
        {
            e.Graphics.FillRectangle(backBrush, e.Bounds);
        }

        // Text zeichnen
        var textColor = isSelected ? Color.White : e.SubItem.ForeColor;
        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, listView.Font, e.Bounds, textColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);

        // Partyhut zeichnen (Spalte 1 = Name)
        if (e.ColumnIndex == 1 && _birthdayTodayList.Contains(e.ItemIndex))
        {
            var bildX = e.Bounds.Right - partyHat.Width - 4;
            var bildY = e.Bounds.Top + (e.Bounds.Height - partyHat.Height) / 2;
            e.Graphics.DrawImage(partyHat, bildX, bildY);
        }
    }

    private void FrmBirthdays_Shown(object sender, EventArgs e)
    {
        BringToFront(); // Sicherstellen, dass es vorne ist
        Activate();

        if (_initialIndex >= 0 && _initialIndex < listView.Items.Count)
        {
            listView.Focus();
            var item = listView.Items[_initialIndex];
            item.Selected = true;
            item.Focused = true;
            listView.EnsureVisible(_initialIndex);
        }
        else
        {
            listView.Focus();
        }
    }

    private void ListView_MouseDoubleClick(object sender, MouseEventArgs e) => DialogResult = DialogResult.OK;

    private void ListView_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Space && listView.SelectedIndices.Count > 0) {DialogResult = DialogResult.OK; }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape) { Close(); return true; }
        return base.ProcessCmdKey(ref msg, keyData);
    }
}