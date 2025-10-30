using System.Diagnostics;
using System.Drawing.Drawing2D;

namespace Adressen;
public partial class FrmProgSettings : Form
{
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool AskBeforeDelete
    {
        get => ckbAskBeforeDelete.Checked;
        set => ckbAskBeforeDelete.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool ColorSchemeBlue
    {
        get => rbtnBlue.Checked;
        set => rbtnBlue.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool? WordProcProg
    {
        get
        {
            if (rbMSWord.Checked) { return true; }
            else if (rbLibreOffice.Checked) { return false; }
            else { return null; }
        }
        set
        {
            if (value is null) { rbManualSelect.Checked = true; }
            else if (value is true) { rbMSWord.Checked = true; }
            else { rbLibreOffice.Checked = true; }
        }
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool ColorSchemeDark
    {
        get => rbtnDark.Checked;
        set => rbtnDark.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool ColorSchemePale
    {
        get => rbtnPale.Checked;
        set => rbtnPale.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool ContactsAutoload
    {
        get => ckbContactsAutoload.Checked;
        set => ckbContactsAutoload.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool NoFile
    {
        get => rbEmpty.Checked;
        set => rbEmpty.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool ReloadRecent
    {
        get => rbRecent.Checked;
        set => rbRecent.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public string StandardFile
    {
        get => tbStandard.Text;
        set => tbStandard.Text = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool DailyBackup
    {
        get => ckbBackup.Checked;
        set => ckbBackup.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool BackupSuccess
    {
        get => ckbBackupSuccess.Checked;
        set => ckbBackupSuccess.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public string BackupDirectory
    {
        get => tbBackupFolder.Text;
        set => tbBackupFolder.Text = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public string DatabaseFolder
    {
        get => tbDatabaseFolder.Text;
        set => tbDatabaseFolder.Text = value;
    }

    //[System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    //public decimal SuccessDuration
    //{
    //    get => numUpDownSuccess.Value;
    //    set => numUpDownSuccess.Value = value;
    //}
    public NumericUpDown NumUpDownSuccess => numUpDownSuccess;

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool AskBeforeSaveSQL
    {
        get => ckbAskBeforeSaveSQL.Checked;
        set => ckbAskBeforeSaveSQL.Checked = value;
    }

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public bool WatchFolder
    {
        get => ckbWatchFolder.Checked;
        set => ckbWatchFolder.Checked = value;
    }   

    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
    public string LetterDirectory
    {
        get => tbWatchFolder.Text;
        set => tbWatchFolder.Text = value;
    }   

    public FrmProgSettings()
    {
        InitializeComponent(); //tabControl.BackColor = SystemColors.ControlLightLight; // funktioniert nicht
    }

    private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
    {
        var g = e.Graphics; //e.DrawBackground();
        g.SmoothingMode = SmoothingMode.HighQuality; //.AntiAlias;
        var tabPage = tabControl.TabPages[e.Index];
        var tabBounds = tabControl.GetTabRect(e.Index); // Get the real bounds for the tab rectangle.
        Brush textBrush;
        if (e.State == DrawItemState.Selected)
        {
            textBrush = new SolidBrush(Color.White);  // Draw a different background color, and don't paint a focus rectangle.
            g.FillRectangle(SystemBrushes.GradientActiveCaption, e.Bounds);
        }
        else
        {
            textBrush = new SolidBrush(e.ForeColor);
            g.FillRectangle(SystemBrushes.GradientInactiveCaption, e.Bounds);
        }
        var tabFont = new Font("Segoe UI", (float)10.0, FontStyle.Regular, GraphicsUnit.Point);
        using var stringFlags = new StringFormat  // Draw string. Center the text.
        {
            Alignment = StringAlignment.Near,
            LineAlignment = StringAlignment.Center
        };
        g.DrawString(tabPage.Text, tabFont, textBrush, tabBounds, new StringFormat(stringFlags));
        g.FillRectangle(new SolidBrush(Color.FromArgb(250, 250, 251)), new Rectangle(0, (tabBounds.Height * tabControl.TabPages.Count) + 3, tabBounds.Width + 2, tabPage.Height - (tabBounds.Height * tabControl.TabPages.Count))); // TabControl.BackColor
        g.DrawLine(new Pen(Color.FromArgb(240, 244, 249)), 0, tabPage.Height + 4, tabBounds.Width + 2, tabPage.Height + 5);
        g.DrawLine(new Pen(Color.FromArgb(160, 160, 160)), 0, tabPage.Height + 6, tabBounds.Width + 2, tabPage.Height + 6);
        g.DrawLine(new Pen(Color.FromArgb(105, 105, 105)), 0, tabPage.Height + 7, tabBounds.Width + 2, tabPage.Height + 7);
    }

    private void FrmProgSettings_Load(object sender, EventArgs e)
    {
        if (!rbEmpty.Checked && !rbRecent.Checked)
        {
            if (string.IsNullOrEmpty(tbStandard.Text)) { rbEmpty.Checked = true; }
            else { rbStandard.Checked = true; }
        }
        tbStandard.Enabled = btnStandardFile.Enabled = rbStandard.Checked;
        if (string.IsNullOrEmpty(tbBackupFolder.Text)) { btnExplorer.Enabled = ckbBackup.Checked = false; }
        tbBackupFolder.Enabled = btnBackupFolder.Enabled = ckbBackup.Checked;
        tbWatchFolder.Enabled = btnWatchFolder.Enabled = lblWatchFolder.Enabled = ckbWatchFolder.Checked;
    }

    private void BtnDatabaseFolder_Click(object sender, EventArgs e)
    {
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK) { tbDatabaseFolder.Text = folderBrowserDialog.SelectedPath; }
    }

    private void BtnBackupFolder_Click(object sender, EventArgs e)
    {
        folderBrowserDialog.InitialDirectory = Directory.Exists(tbBackupFolder.Text) ? tbBackupFolder.Text : "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK) { tbBackupFolder.Text = folderBrowserDialog.SelectedPath; }
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
            case Keys.Tab:
                tabControl.SelectedIndex = (tabControl.SelectedIndex + 1) % tabControl.TabCount;
                return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void CkbBackup_CheckedChanged(object sender, EventArgs e)
    {
        tbBackupFolder.Enabled = btnBackupFolder.Enabled = ckbBackupSuccess.Enabled = numUpDownSuccess.Enabled = labelMS.Enabled = ckbBackup.Checked;
    }

    private void BtnStandardFile_Click(object sender, EventArgs e)
    {
        openFileDialog.InitialDirectory = !string.IsNullOrEmpty(tbStandard.Text) ? Path.GetDirectoryName(tbStandard.Text) : null;
        openFileDialog.CheckFileExists = true;
        if (openFileDialog.ShowDialog() == DialogResult.OK) { tbStandard.Text = openFileDialog.FileName; }
    }

    private void RbStandard_CheckedChanged(object sender, EventArgs e) => btnStandardFile.Enabled = tbStandard.Enabled = rbStandard.Checked;

    private void BtnExplorer_Click(object sender, EventArgs e)
    {
        if (Directory.Exists(tbBackupFolder.Text))
        {
            using var process = new Process();
            process.StartInfo.FileName = tbBackupFolder.Text;
            process.StartInfo.UseShellExecute = true;
            process.Start();
        }
        else { Console.Beep(); }
    }

    private void TbBackupFolder_TextChanged(object sender, EventArgs e)
    {
        btnExplorer.Enabled = !string.IsNullOrEmpty(tbBackupFolder.Text);
    }

    private void CkbWatchFolder_CheckedChanged(object sender, EventArgs e)
    {
        tbWatchFolder.Enabled = btnWatchFolder.Enabled = lblWatchFolder.Enabled = ckbWatchFolder.Checked;
    }

    private void BtnWatchFolder_Click(object sender, EventArgs e)
    {
        if (folderBrowserDialog.ShowDialog() == DialogResult.OK) { tbWatchFolder.Text = folderBrowserDialog.SelectedPath; }
    }

}

