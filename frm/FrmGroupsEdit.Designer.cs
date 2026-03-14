namespace Adressen.frm;

partial class FrmGroupsEdit
{
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        panelRight = new Panel();
        btnCancel = new Button();
        btnClose = new Button();
        btnDelete = new Button();
        btnEdit = new Button();
        listBox = new ListBox();
        columnHeader = new ColumnHeader();
        statusStrip = new StatusStrip();
        toolStripStatusLabel = new ToolStripStatusLabel();
        panelRight.SuspendLayout();
        statusStrip.SuspendLayout();
        SuspendLayout();
        // 
        // panelRight
        // 
        panelRight.Controls.Add(btnCancel);
        panelRight.Controls.Add(btnClose);
        panelRight.Controls.Add(btnDelete);
        panelRight.Controls.Add(btnEdit);
        panelRight.Dock = DockStyle.Right;
        panelRight.Location = new Point(153, 0);
        panelRight.Name = "panelRight";
        panelRight.Size = new Size(120, 224);
        panelRight.TabIndex = 0;
        // 
        // btnCancel
        // 
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(6, 182);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(102, 30);
        btnCancel.TabIndex = 3;
        btnCancel.Text = "Abbrechen";
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // btnClose
        // 
        btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnClose.DialogResult = DialogResult.OK;
        btnClose.Enabled = false;
        btnClose.Location = new Point(6, 146);
        btnClose.Name = "btnClose";
        btnClose.Size = new Size(102, 30);
        btnClose.TabIndex = 0;
        btnClose.Text = "Speichern";
        btnClose.UseVisualStyleBackColor = true;
        // 
        // btnDelete
        // 
        btnDelete.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnDelete.Enabled = false;
        btnDelete.Location = new Point(6, 48);
        btnDelete.Name = "btnDelete";
        btnDelete.Size = new Size(102, 30);
        btnDelete.TabIndex = 2;
        btnDelete.Text = "Löschen";
        btnDelete.UseVisualStyleBackColor = true;
        btnDelete.Click += BtnDelete_Click;
        // 
        // btnEdit
        // 
        btnEdit.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnEdit.Enabled = false;
        btnEdit.Location = new Point(6, 12);
        btnEdit.Name = "btnEdit";
        btnEdit.Size = new Size(102, 30);
        btnEdit.TabIndex = 1;
        btnEdit.Text = "Umbenennen";
        btnEdit.UseVisualStyleBackColor = true;
        btnEdit.Click += BtnEdit_Click;
        // 
        // listBox
        // 
        listBox.Dock = DockStyle.Fill;
        listBox.DrawMode = DrawMode.OwnerDrawFixed;
        listBox.Location = new Point(0, 0);
        listBox.Name = "listBox";
        listBox.Size = new Size(153, 224);
        listBox.TabIndex = 2;
        listBox.DrawItem += ListBox_DrawItem;
        listBox.SelectedIndexChanged += ListBox_SelectedIndexChanged;
        listBox.SizeChanged += ListBox_SizeChanged;
        // 
        // columnHeader
        // 
        columnHeader.Width = 97;
        // 
        // statusStrip
        // 
        statusStrip.Items.AddRange(new ToolStripItem[] { toolStripStatusLabel });
        statusStrip.Location = new Point(0, 224);
        statusStrip.Name = "statusStrip";
        statusStrip.Size = new Size(273, 22);
        statusStrip.TabIndex = 3;
        statusStrip.Text = "statusStrip1";
        // 
        // toolStripStatusLabel
        // 
        toolStripStatusLabel.Name = "toolStripStatusLabel";
        toolStripStatusLabel.Size = new Size(258, 17);
        toolStripStatusLabel.Spring = true;
        toolStripStatusLabel.Text = "0";
        // 
        // FrmGroupsEdit
        // 
        AcceptButton = btnClose;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = btnCancel;
        ClientSize = new Size(273, 246);
        Controls.Add(listBox);
        Controls.Add(panelRight);
        Controls.Add(statusStrip);
        Font = new Font("Segoe UI", 10F);
        MaximizeBox = false;
        MinimizeBox = false;
        MinimumSize = new Size(289, 285);
        Name = "FrmGroupsEdit";
        ShowIcon = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Gruppen bearbeiten";
        Shown += FrmGroups_Shown;
        panelRight.ResumeLayout(false);
        statusStrip.ResumeLayout(false);
        statusStrip.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private Panel panelRight;
    private ListBox listBox;
    private Button btnDelete;
    private Button btnEdit;
    private Button btnClose;
    private ColumnHeader columnHeader;
    private Button btnCancel;
    private StatusStrip statusStrip;
    private ToolStripStatusLabel toolStripStatusLabel;
}