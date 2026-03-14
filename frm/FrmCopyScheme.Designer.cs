namespace Adressen;

partial class FrmCopyScheme
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
        if (_tabFont != null) { _tabFont.Dispose(); }  // Eigene IDisposable-Ressourcen hier freigeben
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCopyScheme));
        tbPattern1 = new TextBox();
        cbxFields = new ComboBox();
        btnInsert = new Button();
        tabControl = new TabControl();
        tabPage1 = new TabPage();
        tabPage2 = new TabPage();
        tbPattern2 = new TextBox();
        tabPage3 = new TabPage();
        tbPattern3 = new TextBox();
        tabPage4 = new TabPage();
        tbPattern4 = new TextBox();
        tabPage5 = new TabPage();
        tbPattern5 = new TextBox();
        tabPage6 = new TabPage();
        tbPattern6 = new TextBox();
        btnCopy = new Button();
        tbResult = new TextBox();
        panelLeft = new Panel();
        panelRight = new Panel();
        statusStrip = new StatusStrip();
        tabControl.SuspendLayout();
        tabPage1.SuspendLayout();
        tabPage2.SuspendLayout();
        tabPage3.SuspendLayout();
        tabPage4.SuspendLayout();
        tabPage5.SuspendLayout();
        tabPage6.SuspendLayout();
        panelLeft.SuspendLayout();
        panelRight.SuspendLayout();
        SuspendLayout();
        // 
        // tbPattern1
        // 
        tbPattern1.AcceptsReturn = true;
        tbPattern1.AcceptsTab = true;
        tbPattern1.BackColor = Color.Ivory;
        tbPattern1.Dock = DockStyle.Fill;
        tbPattern1.Location = new Point(3, 3);
        tbPattern1.Multiline = true;
        tbPattern1.Name = "tbPattern1";
        tbPattern1.Size = new Size(250, 143);
        tbPattern1.TabIndex = 0;
        tbPattern1.WordWrap = false;
        tbPattern1.TextChanged += TbPattern_TextChanged;
        // 
        // cbxFields
        // 
        cbxFields.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        cbxFields.DropDownStyle = ComboBoxStyle.DropDownList;
        cbxFields.FormattingEnabled = true;
        cbxFields.Location = new Point(36, 169);
        cbxFields.Name = "cbxFields";
        cbxFields.Size = new Size(154, 25);
        cbxFields.TabIndex = 1;
        // 
        // btnInsert
        // 
        btnInsert.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnInsert.Location = new Point(196, 167);
        btnInsert.Name = "btnInsert";
        btnInsert.Size = new Size(90, 27);
        btnInsert.TabIndex = 2;
        btnInsert.Text = "Einfügen ⇑";
        btnInsert.UseVisualStyleBackColor = true;
        btnInsert.Click += BtnInsert_Click;
        // 
        // tabControl
        // 
        tabControl.Alignment = TabAlignment.Left;
        tabControl.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        tabControl.Controls.Add(tabPage1);
        tabControl.Controls.Add(tabPage2);
        tabControl.Controls.Add(tabPage3);
        tabControl.Controls.Add(tabPage4);
        tabControl.Controls.Add(tabPage5);
        tabControl.Controls.Add(tabPage6);
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.ItemSize = new Size(25, 25);
        tabControl.Location = new Point(3, 3);
        tabControl.Multiline = true;
        tabControl.Name = "tabControl";
        tabControl.SelectedIndex = 0;
        tabControl.ShowToolTips = true;
        tabControl.Size = new Size(289, 157);
        tabControl.SizeMode = TabSizeMode.Fixed;
        tabControl.TabIndex = 3;
        tabControl.DrawItem += TabControl_DrawItem;
        tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;
        // 
        // tabPage1
        // 
        tabPage1.Controls.Add(tbPattern1);
        tabPage1.Location = new Point(29, 4);
        tabPage1.Name = "tabPage1";
        tabPage1.Padding = new Padding(3);
        tabPage1.Size = new Size(256, 149);
        tabPage1.TabIndex = 0;
        tabPage1.Text = "1";
        tabPage1.UseVisualStyleBackColor = true;
        // 
        // tabPage2
        // 
        tabPage2.Controls.Add(tbPattern2);
        tabPage2.Location = new Point(29, 4);
        tabPage2.Name = "tabPage2";
        tabPage2.Padding = new Padding(3);
        tabPage2.Size = new Size(256, 149);
        tabPage2.TabIndex = 1;
        tabPage2.Text = "2";
        tabPage2.UseVisualStyleBackColor = true;
        // 
        // tbPattern2
        // 
        tbPattern2.AcceptsReturn = true;
        tbPattern2.AcceptsTab = true;
        tbPattern2.BackColor = Color.Ivory;
        tbPattern2.Dock = DockStyle.Fill;
        tbPattern2.Location = new Point(3, 3);
        tbPattern2.Multiline = true;
        tbPattern2.Name = "tbPattern2";
        tbPattern2.Size = new Size(250, 143);
        tbPattern2.TabIndex = 1;
        tbPattern2.WordWrap = false;
        tbPattern2.TextChanged += TbPattern_TextChanged;
        // 
        // tabPage3
        // 
        tabPage3.Controls.Add(tbPattern3);
        tabPage3.Location = new Point(29, 4);
        tabPage3.Name = "tabPage3";
        tabPage3.Padding = new Padding(3);
        tabPage3.Size = new Size(256, 149);
        tabPage3.TabIndex = 2;
        tabPage3.Text = "3";
        tabPage3.UseVisualStyleBackColor = true;
        // 
        // tbPattern3
        // 
        tbPattern3.AcceptsReturn = true;
        tbPattern3.AcceptsTab = true;
        tbPattern3.BackColor = Color.Ivory;
        tbPattern3.Dock = DockStyle.Fill;
        tbPattern3.Location = new Point(3, 3);
        tbPattern3.Multiline = true;
        tbPattern3.Name = "tbPattern3";
        tbPattern3.Size = new Size(250, 143);
        tbPattern3.TabIndex = 1;
        tbPattern3.WordWrap = false;
        tbPattern3.TextChanged += TbPattern_TextChanged;
        // 
        // tabPage4
        // 
        tabPage4.Controls.Add(tbPattern4);
        tabPage4.Location = new Point(29, 4);
        tabPage4.Name = "tabPage4";
        tabPage4.Padding = new Padding(3);
        tabPage4.Size = new Size(256, 149);
        tabPage4.TabIndex = 3;
        tabPage4.Text = "4";
        tabPage4.UseVisualStyleBackColor = true;
        // 
        // tbPattern4
        // 
        tbPattern4.AcceptsReturn = true;
        tbPattern4.AcceptsTab = true;
        tbPattern4.BackColor = Color.Ivory;
        tbPattern4.Dock = DockStyle.Fill;
        tbPattern4.Location = new Point(3, 3);
        tbPattern4.Multiline = true;
        tbPattern4.Name = "tbPattern4";
        tbPattern4.Size = new Size(250, 143);
        tbPattern4.TabIndex = 1;
        tbPattern4.WordWrap = false;
        tbPattern4.TextChanged += TbPattern_TextChanged;
        // 
        // tabPage5
        // 
        tabPage5.Controls.Add(tbPattern5);
        tabPage5.Location = new Point(29, 4);
        tabPage5.Name = "tabPage5";
        tabPage5.Padding = new Padding(3);
        tabPage5.Size = new Size(256, 149);
        tabPage5.TabIndex = 4;
        tabPage5.Text = "5";
        tabPage5.UseVisualStyleBackColor = true;
        // 
        // tbPattern5
        // 
        tbPattern5.AcceptsReturn = true;
        tbPattern5.AcceptsTab = true;
        tbPattern5.BackColor = Color.Ivory;
        tbPattern5.Dock = DockStyle.Fill;
        tbPattern5.Location = new Point(3, 3);
        tbPattern5.Multiline = true;
        tbPattern5.Name = "tbPattern5";
        tbPattern5.Size = new Size(250, 143);
        tbPattern5.TabIndex = 1;
        tbPattern5.WordWrap = false;
        tbPattern5.TextChanged += TbPattern_TextChanged;
        // 
        // tabPage6
        // 
        tabPage6.Controls.Add(tbPattern6);
        tabPage6.Location = new Point(29, 4);
        tabPage6.Name = "tabPage6";
        tabPage6.Padding = new Padding(3);
        tabPage6.Size = new Size(256, 149);
        tabPage6.TabIndex = 5;
        tabPage6.Text = "6";
        tabPage6.UseVisualStyleBackColor = true;
        // 
        // tbPattern6
        // 
        tbPattern6.AcceptsReturn = true;
        tbPattern6.AcceptsTab = true;
        tbPattern6.BackColor = Color.Ivory;
        tbPattern6.Dock = DockStyle.Fill;
        tbPattern6.Location = new Point(3, 3);
        tbPattern6.Multiline = true;
        tbPattern6.Name = "tbPattern6";
        tbPattern6.Size = new Size(250, 143);
        tbPattern6.TabIndex = 1;
        tbPattern6.WordWrap = false;
        tbPattern6.TextChanged += TbPattern_TextChanged;
        // 
        // btnCopy
        // 
        btnCopy.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnCopy.DialogResult = DialogResult.OK;
        btnCopy.Image = Properties.Resources.clipboard_plus16;
        btnCopy.Location = new Point(3, 167);
        btnCopy.Name = "btnCopy";
        btnCopy.Size = new Size(252, 27);
        btnCopy.TabIndex = 5;
        btnCopy.Text = "Text in Zwischenablage kopieren";
        btnCopy.TextAlign = ContentAlignment.MiddleRight;
        btnCopy.TextImageRelation = TextImageRelation.ImageBeforeText;
        btnCopy.UseVisualStyleBackColor = true;
        btnCopy.Click += BtnCopy_Click;
        // 
        // tbResult
        // 
        tbResult.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        tbResult.BackColor = Color.AliceBlue;
        tbResult.Location = new Point(5, 10);
        tbResult.Multiline = true;
        tbResult.Name = "tbResult";
        tbResult.ReadOnly = true;
        tbResult.Size = new Size(250, 143);
        tbResult.TabIndex = 0;
        tbResult.WordWrap = false;
        // 
        // panelLeft
        // 
        panelLeft.BackColor = SystemColors.ControlLightLight;
        panelLeft.Controls.Add(btnInsert);
        panelLeft.Controls.Add(cbxFields);
        panelLeft.Controls.Add(tabControl);
        panelLeft.Dock = DockStyle.Fill;
        panelLeft.Location = new Point(0, 0);
        panelLeft.Name = "panelLeft";
        panelLeft.Size = new Size(298, 196);
        panelLeft.TabIndex = 6;
        // 
        // panelRight
        // 
        panelRight.Controls.Add(btnCopy);
        panelRight.Controls.Add(tbResult);
        panelRight.Dock = DockStyle.Right;
        panelRight.Location = new Point(298, 0);
        panelRight.Name = "panelRight";
        panelRight.Size = new Size(264, 196);
        panelRight.TabIndex = 7;
        // 
        // statusStrip
        // 
        statusStrip.AutoSize = false;
        statusStrip.BackColor = Color.Transparent;
        statusStrip.BackgroundImageLayout = ImageLayout.None;
        statusStrip.Location = new Point(0, 196);
        statusStrip.Name = "statusStrip";
        statusStrip.Size = new Size(562, 20);
        statusStrip.TabIndex = 8;
        statusStrip.Text = "statusStrip";
        statusStrip.Paint += StatusStrip_Paint;
        // 
        // FrmCopyScheme
        // 
        AcceptButton = btnCopy;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(562, 216);
        Controls.Add(panelLeft);
        Controls.Add(panelRight);
        Controls.Add(statusStrip);
        Font = new Font("Segoe UI", 10F);
        Icon = (Icon)resources.GetObject("$this.Icon");
        MaximizeBox = false;
        MinimizeBox = false;
        MinimumSize = new Size(578, 255);
        Name = "FrmCopyScheme";
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Show;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Kopierschemata";
        Load += FrmCopyScheme_Load;
        Shown += FrmCopyScheme_Shown;
        tabControl.ResumeLayout(false);
        tabPage1.ResumeLayout(false);
        tabPage1.PerformLayout();
        tabPage2.ResumeLayout(false);
        tabPage2.PerformLayout();
        tabPage3.ResumeLayout(false);
        tabPage3.PerformLayout();
        tabPage4.ResumeLayout(false);
        tabPage4.PerformLayout();
        tabPage5.ResumeLayout(false);
        tabPage5.PerformLayout();
        tabPage6.ResumeLayout(false);
        tabPage6.PerformLayout();
        panelLeft.ResumeLayout(false);
        panelRight.ResumeLayout(false);
        panelRight.PerformLayout();
        ResumeLayout(false);
    }

    #endregion

    private TextBox tbPattern1;
    private ComboBox cbxFields;
    private Button btnCopy;
    private Button btnInsert;
    private TextBox tbResult;
    private TabControl tabControl;
    private TabPage tabPage1;
    private TabPage tabPage2;
    private TabPage tabPage3;
    private TextBox tbPattern2;
    private TextBox tbPattern3;
    private Panel panelLeft;
    private TabPage tabPage4;
    private TextBox tbPattern4;
    private TabPage tabPage5;
    private TextBox tbPattern5;
    private TabPage tabPage6;
    private TextBox tbPattern6;
    private Panel panelRight;
    private StatusStrip statusStrip;
}