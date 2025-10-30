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
        listView = new ListView();
        columnHeader = new ColumnHeader();
        panelRight.SuspendLayout();
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
        panelRight.Size = new Size(120, 217);
        panelRight.TabIndex = 0;
        // 
        // btnCancel
        // 
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(6, 174);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(102, 30);
        btnCancel.TabIndex = 3;
        btnCancel.Text = "Abbrechen";
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // btnClose
        // 
        btnClose.DialogResult = DialogResult.OK;
        btnClose.Enabled = false;
        btnClose.Location = new Point(6, 138);
        btnClose.Name = "btnClose";
        btnClose.Size = new Size(102, 30);
        btnClose.TabIndex = 0;
        btnClose.Text = "Ausführen";
        btnClose.UseVisualStyleBackColor = true;
        // 
        // btnDelete
        // 
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
        btnEdit.Enabled = false;
        btnEdit.Location = new Point(6, 12);
        btnEdit.Name = "btnEdit";
        btnEdit.Size = new Size(102, 30);
        btnEdit.TabIndex = 1;
        btnEdit.Text = "Umbenennen";
        btnEdit.UseVisualStyleBackColor = true;
        btnEdit.Click += BtnEdit_Click;
        // 
        // listView
        // 
        listView.Columns.AddRange(new ColumnHeader[] { columnHeader });
        listView.Dock = DockStyle.Fill;
        listView.FullRowSelect = true;
        listView.HeaderStyle = ColumnHeaderStyle.None;
        listView.Location = new Point(0, 0);
        listView.MultiSelect = false;
        listView.Name = "listView";
        listView.Size = new Size(153, 217);
        listView.TabIndex = 2;
        listView.UseCompatibleStateImageBehavior = false;
        listView.View = View.Details;
        listView.SelectedIndexChanged += ListView_SelectedIndexChanged;
        // 
        // columnHeader
        // 
        columnHeader.Width = 97;
        // 
        // FrmGroupsEdit
        // 
        AcceptButton = btnClose;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = btnCancel;
        ClientSize = new Size(273, 217);
        Controls.Add(listView);
        Controls.Add(panelRight);
        Font = new Font("Segoe UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "FrmGroupsEdit";
        ShowIcon = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Gruppen bearbeiten";
        Shown += FrmGroups_Shown;
        panelRight.ResumeLayout(false);
        ResumeLayout(false);
    }

    #endregion

    private Panel panelRight;
    private ListView listView;
    private Button btnDelete;
    private Button btnEdit;
    private Button btnClose;
    private ColumnHeader columnHeader;
    private Button btnCancel;
}