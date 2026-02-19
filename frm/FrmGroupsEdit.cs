using System.Data;

namespace Adressen.frm;

public partial class FrmGroupsEdit : Form
{
    public Dictionary<string, string> groupNameMap = [];

    public FrmGroupsEdit(Dictionary<string, int> groupDict)
    {
        InitializeComponent();
        //var sortedGroups = groupDict.OrderByDescending(kvp => kvp.Value);
        //groupNameMap = sortedGroups.ToDictionary(kvp => kvp.Key, kvp => kvp.Key);
        //var listViewIndex = 0;
        //foreach (var kvp in sortedGroups)
        //{
        //    var item = new ListViewItem($"{kvp.Key} ({kvp.Value})") { Tag = listViewIndex };
        //    listView.Items.Add(item);
        //    listViewIndex++;
        //}
        // ★ nach oben, danach nach Anzahl absteigend sortieren
        var sortedGroups = groupDict
            .OrderByDescending(kvp => kvp.Key == "★")
            .ThenByDescending(kvp => kvp.Value)
            .ToList();

        groupNameMap = sortedGroups.ToDictionary(kvp => kvp.Key, kvp => kvp.Key);

        listView.Items.Clear();
        for (var i = 0; i < sortedGroups.Count; i++)
        {
            var kvp = sortedGroups[i];
            var item = new ListViewItem($"{kvp.Key} ({kvp.Value})") { Tag = i };
            listView.Items.Add(item);
        }
    }

    private void BtnDelete_Click(object sender, EventArgs e)
    {
        if (listView.SelectedItems.Count > 0)
        {
            var originalIndex = listView.SelectedItems[0].Tag as int? ?? -1;
            var oldGroupName = groupNameMap.Keys.ElementAt(originalIndex);
            if (oldGroupName == "★") { return; }
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
            var oldGroupName = groupNameMap.Keys.ElementAt(originalIndex);
            if (oldGroupName == "★") { return; }
            using var frm = new FrmGroupRename(oldGroupName);
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                var text = frm.GetText();
                if (!string.IsNullOrEmpty(text))
                {
                    listView.SelectedItems[0].Text = $"{text} (geändert)"; // Optionaler Hinweis
                    groupNameMap[oldGroupName] = text;
                    btnClose.Enabled = true;
                    btnClose.Focus();
                }
            }
        }
    }

    private void ListView_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (listView.SelectedItems.Count > 0)
        {
            // Den Originalnamen über den im Tag gespeicherten Index ermitteln
            var originalIndex = listView.SelectedItems[0].Tag as int? ?? -1;
            var groupNames = groupNameMap.Keys.ToList();

            if (originalIndex >= 0 && originalIndex < groupNames.Count)
            {
                var currentName = groupNames[originalIndex];
                var isSpecialGroup = currentName == "★";

                btnEdit.Enabled = !isSpecialGroup;
                btnDelete.Enabled = !isSpecialGroup;
                return;
            }
        }
        btnEdit.Enabled = btnDelete.Enabled = false;
    }

    private void FrmGroups_Shown(object sender, EventArgs e) => listView.Focus();
}
