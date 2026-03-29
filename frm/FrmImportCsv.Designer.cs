namespace Adressen.frm;

partial class FrmImportCsv
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
        var dataGridViewCellStyle1 = new DataGridViewCellStyle();
        var dataGridViewCellStyle3 = new DataGridViewCellStyle();
        var dataGridViewCellStyle4 = new DataGridViewCellStyle();
        var dataGridViewCellStyle5 = new DataGridViewCellStyle();
        var dataGridViewCellStyle2 = new DataGridViewCellStyle();
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmImportCsv));
        gbSourceFile = new GroupBox();
        btnBrowse = new Button();
        txtCsvPath = new TextBox();
        lblSourceFile = new Label();
        gbTarget = new GroupBox();
        rbNewDb = new RadioButton();
        rbCurrentDb = new RadioButton();
        gbMapping = new GroupBox();
        dgvMapping = new DataGridView();
        csvCol = new DataGridViewTextBoxColumn();
        exampleCol = new DataGridViewTextBoxColumn();
        comboCol = new DataGridViewComboBoxColumn();
        statusStrip = new StatusStrip();
        progressBar = new ToolStripProgressBar();
        toolStripStatusLabel = new ToolStripStatusLabel();
        btnStartImport = new Button();
        btnCancel = new Button();
        lnkExample = new LinkLabel();
        rbDuplicateSkip = new RadioButton();
        rbDuplicateCreate = new RadioButton();
        gbDuplicate = new GroupBox();
        gbSourceFile.SuspendLayout();
        gbTarget.SuspendLayout();
        gbMapping.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvMapping).BeginInit();
        statusStrip.SuspendLayout();
        gbDuplicate.SuspendLayout();
        SuspendLayout();
        // 
        // gbSourceFile
        // 
        gbSourceFile.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        gbSourceFile.Controls.Add(btnBrowse);
        gbSourceFile.Controls.Add(txtCsvPath);
        gbSourceFile.Controls.Add(lblSourceFile);
        gbSourceFile.Font = new Font("Segoe UI", 9F);
        gbSourceFile.Location = new Point(12, 12);
        gbSourceFile.Name = "gbSourceFile";
        gbSourceFile.Size = new Size(363, 61);
        gbSourceFile.TabIndex = 0;
        gbSourceFile.TabStop = false;
        gbSourceFile.Text = "Quelldatei";
        // 
        // btnBrowse
        // 
        btnBrowse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnBrowse.Font = new Font("Segoe UI", 10F);
        btnBrowse.Location = new Point(321, 23);
        btnBrowse.Name = "btnBrowse";
        btnBrowse.Size = new Size(36, 25);
        btnBrowse.TabIndex = 5;
        btnBrowse.Text = "⚙";
        btnBrowse.UseVisualStyleBackColor = true;
        btnBrowse.Click += BtnBrowse_Click;
        // 
        // txtCsvPath
        // 
        txtCsvPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        txtCsvPath.Font = new Font("Segoe UI", 10F);
        txtCsvPath.Location = new Point(87, 24);
        txtCsvPath.Name = "txtCsvPath";
        txtCsvPath.Size = new Size(228, 25);
        txtCsvPath.TabIndex = 1;
        // 
        // lblSourceFile
        // 
        lblSourceFile.AutoSize = true;
        lblSourceFile.Font = new Font("Segoe UI", 10F);
        lblSourceFile.Location = new Point(6, 27);
        lblSourceFile.Name = "lblSourceFile";
        lblSourceFile.Size = new Size(75, 19);
        lblSourceFile.TabIndex = 0;
        lblSourceFile.Text = "CSV-Datei:";
        // 
        // gbTarget
        // 
        gbTarget.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        gbTarget.Controls.Add(rbNewDb);
        gbTarget.Controls.Add(rbCurrentDb);
        gbTarget.Font = new Font("Segoe UI", 9F);
        gbTarget.Location = new Point(12, 79);
        gbTarget.Name = "gbTarget";
        gbTarget.Size = new Size(363, 81);
        gbTarget.TabIndex = 1;
        gbTarget.TabStop = false;
        gbTarget.Text = "Zieldatenbank";
        // 
        // rbNewDb
        // 
        rbNewDb.AutoSize = true;
        rbNewDb.Checked = true;
        rbNewDb.Font = new Font("Segoe UI", 10F);
        rbNewDb.Location = new Point(6, 53);
        rbNewDb.Name = "rbNewDb";
        rbNewDb.Size = new Size(337, 23);
        rbNewDb.TabIndex = 1;
        rbNewDb.TabStop = true;
        rbNewDb.Text = "Neue Datenbank erstellen und dorthin importieren";
        rbNewDb.UseVisualStyleBackColor = true;
        // 
        // rbCurrentDb
        // 
        rbCurrentDb.AutoSize = true;
        rbCurrentDb.Enabled = false;
        rbCurrentDb.Font = new Font("Segoe UI", 10F);
        rbCurrentDb.Location = new Point(6, 24);
        rbCurrentDb.Name = "rbCurrentDb";
        rbCurrentDb.Size = new Size(291, 23);
        rbCurrentDb.TabIndex = 0;
        rbCurrentDb.Text = "In aktuell geöffnete Datenbank importieren";
        rbCurrentDb.UseVisualStyleBackColor = true;
        // 
        // gbMapping
        // 
        gbMapping.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        gbMapping.Controls.Add(dgvMapping);
        gbMapping.Font = new Font("Segoe UI", 9F);
        gbMapping.Location = new Point(12, 226);
        gbMapping.Name = "gbMapping";
        gbMapping.Size = new Size(363, 283);
        gbMapping.TabIndex = 2;
        gbMapping.TabStop = false;
        gbMapping.Text = "Spaltenzuordnung";
        // 
        // dgvMapping
        // 
        dgvMapping.AllowUserToAddRows = false;
        dgvMapping.AllowUserToDeleteRows = false;
        dgvMapping.AllowUserToResizeColumns = false;
        dgvMapping.AllowUserToResizeRows = false;
        dgvMapping.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        dgvMapping.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        dgvMapping.BackgroundColor = SystemColors.ControlLightLight;
        dgvMapping.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
        dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
        dataGridViewCellStyle1.BackColor = SystemColors.ControlDark;
        dataGridViewCellStyle1.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
        dataGridViewCellStyle1.ForeColor = SystemColors.HighlightText;
        dataGridViewCellStyle1.SelectionBackColor = SystemColors.ControlDark;
        dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
        dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
        dgvMapping.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
        dgvMapping.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvMapping.Columns.AddRange(new DataGridViewColumn[] { csvCol, exampleCol, comboCol });
        dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
        dataGridViewCellStyle3.BackColor = SystemColors.Window;
        dataGridViewCellStyle3.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
        dataGridViewCellStyle3.ForeColor = SystemColors.ControlText;
        dataGridViewCellStyle3.SelectionBackColor = SystemColors.Highlight;
        dataGridViewCellStyle3.SelectionForeColor = SystemColors.HighlightText;
        dataGridViewCellStyle3.WrapMode = DataGridViewTriState.False;
        dgvMapping.DefaultCellStyle = dataGridViewCellStyle3;
        dgvMapping.EditMode = DataGridViewEditMode.EditOnEnter;
        dgvMapping.EnableHeadersVisualStyles = false;
        dgvMapping.Location = new Point(6, 24);
        dgvMapping.MultiSelect = false;
        dgvMapping.Name = "dgvMapping";
        dataGridViewCellStyle4.Alignment = DataGridViewContentAlignment.MiddleLeft;
        dataGridViewCellStyle4.BackColor = SystemColors.Control;
        dataGridViewCellStyle4.Font = new Font("Segoe UI", 10F);
        dataGridViewCellStyle4.ForeColor = SystemColors.WindowText;
        dataGridViewCellStyle4.SelectionBackColor = SystemColors.Highlight;
        dataGridViewCellStyle4.SelectionForeColor = SystemColors.HighlightText;
        dataGridViewCellStyle4.WrapMode = DataGridViewTriState.True;
        dgvMapping.RowHeadersDefaultCellStyle = dataGridViewCellStyle4;
        dgvMapping.RowHeadersVisible = false;
        dataGridViewCellStyle5.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
        dgvMapping.RowsDefaultCellStyle = dataGridViewCellStyle5;
        dgvMapping.ScrollBars = ScrollBars.Vertical;
        dgvMapping.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dgvMapping.Size = new Size(349, 250);
        dgvMapping.TabIndex = 0;
        dgvMapping.CellValueChanged += DgvMapping_CellValueChanged;
        dgvMapping.CurrentCellDirtyStateChanged += DgvMapping_CurrentCellDirtyStateChanged;
        dgvMapping.EditingControlShowing += DgvMapping_EditingControlShowing;
        // 
        // csvCol
        // 
        csvCol.FillWeight = 30F;
        csvCol.HeaderText = "CSV-Spalte";
        csvCol.Name = "csvCol";
        csvCol.ReadOnly = true;
        // 
        // exampleCol
        // 
        dataGridViewCellStyle2.ForeColor = Color.Gray;
        exampleCol.DefaultCellStyle = dataGridViewCellStyle2;
        exampleCol.FillWeight = 30F;
        exampleCol.HeaderText = "Zeileninhalt";
        exampleCol.Name = "exampleCol";
        exampleCol.ReadOnly = true;
        // 
        // comboCol
        // 
        comboCol.FillWeight = 40F;
        comboCol.FlatStyle = FlatStyle.Flat;
        comboCol.HeaderText = "Programmfeld";
        comboCol.Name = "comboCol";
        // 
        // statusStrip
        // 
        statusStrip.Items.AddRange(new ToolStripItem[] { progressBar, toolStripStatusLabel });
        statusStrip.Location = new Point(0, 545);
        statusStrip.Name = "statusStrip";
        statusStrip.Size = new Size(387, 22);
        statusStrip.TabIndex = 3;
        statusStrip.Text = "statusStrip1";
        // 
        // progressBar
        // 
        progressBar.Margin = new Padding(11, 3, 1, 3);
        progressBar.Name = "progressBar";
        progressBar.Size = new Size(100, 16);
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.Visible = false;
        // 
        // toolStripStatusLabel
        // 
        toolStripStatusLabel.Margin = new Padding(11, 3, 0, 2);
        toolStripStatusLabel.Name = "toolStripStatusLabel";
        toolStripStatusLabel.Size = new Size(361, 17);
        toolStripStatusLabel.Spring = true;
        toolStripStatusLabel.Text = " ";
        // 
        // btnStartImport
        // 
        btnStartImport.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnStartImport.Enabled = false;
        btnStartImport.Location = new Point(158, 515);
        btnStartImport.Name = "btnStartImport";
        btnStartImport.Size = new Size(121, 27);
        btnStartImport.TabIndex = 4;
        btnStartImport.Text = "&Import starten…";
        btnStartImport.UseVisualStyleBackColor = true;
        btnStartImport.Click += BtnStartImport_Click;
        // 
        // btnCancel
        // 
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnCancel.Location = new Point(285, 515);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(90, 27);
        btnCancel.TabIndex = 5;
        btnCancel.Text = "&Abbrechen";
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // lnkExample
        // 
        lnkExample.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        lnkExample.AutoSize = true;
        lnkExample.Location = new Point(12, 520);
        lnkExample.Name = "lnkExample";
        lnkExample.Size = new Size(140, 19);
        lnkExample.TabIndex = 6;
        lnkExample.TabStop = true;
        lnkExample.Text = "Beispiel-CSV erstellen";
        lnkExample.LinkClicked += LnkExample_LinkClicked;
        // 
        // rbDuplicateSkip
        // 
        rbDuplicateSkip.AutoSize = true;
        rbDuplicateSkip.Font = new Font("Segoe UI", 10F);
        rbDuplicateSkip.Location = new Point(166, 24);
        rbDuplicateSkip.Name = "rbDuplicateSkip";
        rbDuplicateSkip.Size = new Size(184, 23);
        rbDuplicateSkip.TabIndex = 1;
        rbDuplicateSkip.Text = "Überspringen (ignorieren)";
        rbDuplicateSkip.UseVisualStyleBackColor = true;
        // 
        // rbDuplicateCreate
        // 
        rbDuplicateCreate.AutoSize = true;
        rbDuplicateCreate.Checked = true;
        rbDuplicateCreate.Font = new Font("Segoe UI", 10F);
        rbDuplicateCreate.Location = new Point(6, 24);
        rbDuplicateCreate.Name = "rbDuplicateCreate";
        rbDuplicateCreate.Size = new Size(146, 23);
        rbDuplicateCreate.TabIndex = 0;
        rbDuplicateCreate.TabStop = true;
        rbDuplicateCreate.Text = "Immer neu anlegen";
        rbDuplicateCreate.UseVisualStyleBackColor = true;
        // 
        // gbDuplicate
        // 
        gbDuplicate.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        gbDuplicate.Controls.Add(rbDuplicateSkip);
        gbDuplicate.Controls.Add(rbDuplicateCreate);
        gbDuplicate.Enabled = false;
        gbDuplicate.Font = new Font("Segoe UI", 9F);
        gbDuplicate.Location = new Point(12, 166);
        gbDuplicate.Name = "gbDuplicate";
        gbDuplicate.Size = new Size(363, 54);
        gbDuplicate.TabIndex = 7;
        gbDuplicate.TabStop = false;
        gbDuplicate.Text = "Bei Duplikaten";
        // 
        // FrmImportCsv
        // 
        AcceptButton = btnStartImport;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = btnCancel;
        ClientSize = new Size(387, 567);
        Controls.Add(gbDuplicate);
        Controls.Add(lnkExample);
        Controls.Add(btnCancel);
        Controls.Add(btnStartImport);
        Controls.Add(statusStrip);
        Controls.Add(gbMapping);
        Controls.Add(gbTarget);
        Controls.Add(gbSourceFile);
        Font = new Font("Segoe UI", 10F);
        Icon = (Icon)resources.GetObject("$this.Icon");
        MaximizeBox = false;
        MinimizeBox = false;
        MinimumSize = new Size(403, 606);
        Name = "FrmImportCsv";
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Show;
        StartPosition = FormStartPosition.CenterParent;
        Text = "CSV-Datei importieren";
        gbSourceFile.ResumeLayout(false);
        gbSourceFile.PerformLayout();
        gbTarget.ResumeLayout(false);
        gbTarget.PerformLayout();
        gbMapping.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvMapping).EndInit();
        statusStrip.ResumeLayout(false);
        statusStrip.PerformLayout();
        gbDuplicate.ResumeLayout(false);
        gbDuplicate.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private GroupBox gbSourceFile;
    private TextBox txtCsvPath;
    private Label lblSourceFile;
    private Button btnBrowse;
    private GroupBox gbTarget;
    private RadioButton rbNewDb;
    private RadioButton rbCurrentDb;
    private GroupBox gbMapping;
    private DataGridView dgvMapping;
    private StatusStrip statusStrip;
    private ToolStripProgressBar progressBar;
    private Button btnStartImport;
    private Button btnCancel;
    private LinkLabel lnkExample;
    private ToolStripStatusLabel toolStripStatusLabel;
    private RadioButton rbDuplicateSkip;
    private RadioButton rbDuplicateCreate;
    private GroupBox gbDuplicate;
    private DataGridViewTextBoxColumn csvCol;
    private DataGridViewTextBoxColumn exampleCol;
    private DataGridViewComboBoxColumn comboCol;
}