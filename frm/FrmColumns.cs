namespace Adressen;

public partial class FrmColumns : Form
{
    public ListView GetColumnList() => listView;
    public void SetColumnList(ListView value) => listView = value;
    private readonly bool[] hideColumnArr; // = new bool[24];

    public FrmColumns(bool[] boolArray)
    {
        InitializeComponent();
        hideColumnArr = boolArray;
        //listView.Items[^1].Text = RessourceName;
    }

    private void BtnStandard_Click(object sender, EventArgs e)
    {
        for (var i = 0; i < listView.Items.Count; i++) { listView.Items[i].Checked = !hideColumnArr[i]; }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        switch (keyData)
        {
            case Keys.Escape:
                {
                    Close();
                    return true;
                }
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }
}
