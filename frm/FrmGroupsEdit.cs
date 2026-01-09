using System.Data;

namespace Adressen.frm;

public partial class FrmGroupsEdit : Form
{
    public Dictionary<string, string> groupNameMap = [];

    public FrmGroupsEdit(Dictionary<string, int> groupDict)
    {
        InitializeComponent();
        var sortedGroups = groupDict.OrderByDescending(kvp => kvp.Value);
        groupNameMap = sortedGroups.ToDictionary(kvp => kvp.Key, kvp => kvp.Key);
        var listViewIndex = 0;
        foreach (var kvp in sortedGroups)
        {
            var item = new ListViewItem($"{kvp.Key} ({kvp.Value})") { Tag = listViewIndex };
            listView.Items.Add(item);
            listViewIndex++;
        }
    }

    private void BtnDelete_Click(object sender, EventArgs e)
    {
        if (listView.SelectedItems.Count > 0)
        {
            var oldGroupName = groupNameMap.Select(kvp => kvp.Key).ToList()[listView.SelectedItems[0].Index];
            groupNameMap[oldGroupName] = string.Empty;
            listView.Items.Remove(listView.SelectedItems[0]);
            btnClose.Enabled = true;
            btnClose.Focus();
        }
    }

    private void BtnEdit_Click(object sender, EventArgs e)
    {
        if (listView.SelectedItems.Count > 0)
        {
            var originalIndex = listView.SelectedItems[0].Tag as int? ?? -1;
            var oldGroupName = groupNameMap.Select(kvp => kvp.Key).ToList()[originalIndex];
            using var frm = new FrmGroupRename(oldGroupName);
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                if (frm.GetText() is string text && !string.IsNullOrEmpty(text))
                {
                    listView.SelectedItems[0].Text = groupNameMap[oldGroupName] = text;
                    btnClose.Enabled = groupNameMap.Any(kvp => kvp.Key != kvp.Value || string.IsNullOrEmpty(kvp.Value));
                    if (btnClose.Enabled) { btnClose.Focus(); }
                }
            }
        }
    }

    private void ListView_SelectedIndexChanged(object sender, EventArgs e) => btnEdit.Enabled = btnDelete.Enabled = listView.SelectedItems.Count > 0;

    private void FrmGroups_Shown(object sender, EventArgs e) => listView.Focus();
}
