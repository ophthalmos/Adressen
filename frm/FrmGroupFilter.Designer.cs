namespace Adressen.frm;

partial class FrmGroupFilter
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
        panelBottom = new Panel();
        buttonAll = new Button();
        buttonFilter = new Button();
        buttonCancel = new Button();
        panelTop = new Panel();
        labelHeader = new Label();
        tableLayoutPanel = new TableLayoutPanel();
        panelParent = new Panel();
        panelBottom.SuspendLayout();
        panelTop.SuspendLayout();
        panelParent.SuspendLayout();
        SuspendLayout();
        // 
        // panelBottom
        // 
        panelBottom.BackColor = SystemColors.ControlLight;
        panelBottom.Controls.Add(buttonAll);
        panelBottom.Controls.Add(buttonFilter);
        panelBottom.Controls.Add(buttonCancel);
        panelBottom.Dock = DockStyle.Bottom;
        panelBottom.Location = new Point(0, 395);
        panelBottom.Name = "panelBottom";
        panelBottom.Size = new Size(280, 45);
        panelBottom.TabIndex = 5;
        // 
        // buttonAll
        // 
        buttonAll.Location = new Point(12, 6);
        buttonAll.Name = "buttonAll";
        buttonAll.Size = new Size(46, 27);
        buttonAll.TabIndex = 1;
        buttonAll.Text = "Alle";
        buttonAll.UseVisualStyleBackColor = true;
        buttonAll.Click += ButtonAll_Click;
        // 
        // buttonFilter
        // 
        buttonFilter.DialogResult = DialogResult.OK;
        buttonFilter.Location = new Point(64, 6);
        buttonFilter.Name = "buttonFilter";
        buttonFilter.Size = new Size(114, 27);
        buttonFilter.TabIndex = 0;
        buttonFilter.Text = "Filter anwenden";
        buttonFilter.UseVisualStyleBackColor = true;
        buttonFilter.Click += ButtonFilter_Click;
        // 
        // buttonCancel
        // 
        buttonCancel.DialogResult = DialogResult.Cancel;
        buttonCancel.Location = new Point(184, 6);
        buttonCancel.Name = "buttonCancel";
        buttonCancel.Size = new Size(84, 27);
        buttonCancel.TabIndex = 0;
        buttonCancel.Text = "Abbrechen";
        buttonCancel.UseVisualStyleBackColor = true;
        // 
        // panelTop
        // 
        panelTop.Controls.Add(labelHeader);
        panelTop.Dock = DockStyle.Top;
        panelTop.Location = new Point(0, 0);
        panelTop.Name = "panelTop";
        panelTop.Size = new Size(280, 33);
        panelTop.TabIndex = 0;
        // 
        // labelHeader
        // 
        labelHeader.BackColor = SystemColors.ControlLight;
        labelHeader.Dock = DockStyle.Fill;
        labelHeader.Location = new Point(0, 0);
        labelHeader.Name = "labelHeader";
        labelHeader.Padding = new Padding(8, 8, 2, 8);
        labelHeader.Size = new Size(280, 33);
        labelHeader.TabIndex = 4;
        labelHeader.Text = "Einschluss   Ausschluss";
        labelHeader.TextAlign = ContentAlignment.TopRight;
        // 
        // tableLayoutPanel
        // 
        tableLayoutPanel.AutoScroll = true;
        tableLayoutPanel.AutoSize = true;
        tableLayoutPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        tableLayoutPanel.BackColor = SystemColors.ControlLightLight;
        tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
        tableLayoutPanel.ColumnCount = 3;
        tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48F));
        tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 26F));
        tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 26F));
        tableLayoutPanel.Dock = DockStyle.Top;
        tableLayoutPanel.Location = new Point(0, 0);
        tableLayoutPanel.Name = "tableLayoutPanel";
        tableLayoutPanel.RowCount = 1;
        tableLayoutPanel.RowStyles.Add(new RowStyle());
        tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
        tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
        tableLayoutPanel.Size = new Size(280, 2);
        tableLayoutPanel.TabIndex = 2;
        // 
        // panelParent
        // 
        panelParent.AutoScroll = true;
        panelParent.Controls.Add(tableLayoutPanel);
        panelParent.Dock = DockStyle.Fill;
        panelParent.Location = new Point(0, 33);
        panelParent.Name = "panelParent";
        panelParent.Size = new Size(280, 362);
        panelParent.TabIndex = 3;
        // 
        // FrmGroupFilter
        // 
        AcceptButton = buttonFilter;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = buttonCancel;
        ClientSize = new Size(280, 440);
        Controls.Add(panelParent);
        Controls.Add(panelTop);
        Controls.Add(panelBottom);
        Font = new Font("Segoe UI", 10F);
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "FrmGroupFilter";
        ShowIcon = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Nach Gruppenzugehörigkeit filtern";
        Load += FrmGroupFilter_Load;
        panelBottom.ResumeLayout(false);
        panelTop.ResumeLayout(false);
        panelParent.ResumeLayout(false);
        panelParent.PerformLayout();
        ResumeLayout(false);
    }

    #endregion
    private Panel panelBottom;
    private Button buttonCancel;
    private Panel panelTop;
    private Label labelHeader;
    private Button buttonFilter;
    private TableLayoutPanel tableLayoutPanel;
    private Panel panelParent;
    private Button buttonAll;
}