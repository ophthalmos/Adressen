namespace Adressen.frm;

partial class FrmSinglePrintPreview
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
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSinglePrintPreview));
        toolStrip = new ToolStrip();
        lblDevice = new ToolStripLabel();
        cbPrinter = new ToolStripComboBox();
        lblSource = new ToolStripLabel();
        cbSources = new ToolStripComboBox();
        toolStripSeparator = new ToolStripSeparator();
        btnZoom = new ToolStripButton();
        toolStripSeparator1 = new ToolStripSeparator();
        splitBtnPaperSize = new ToolStripSplitButton();
        splitItemA4 = new ToolStripMenuItem();
        splitItemA5 = new ToolStripMenuItem();
        btnFontSize = new ToolStripButton();
        btnShowNotes = new ToolStripButton();
        toolStripSeparator2 = new ToolStripSeparator();
        btnPrint = new ToolStripButton();
        toolStripSeparator3 = new ToolStripSeparator();
        btnClose = new ToolStripButton();
        statusStrip = new StatusStrip();
        toolStripStatusLabel = new ToolStripStatusLabel();
        printPreviewControl = new Adressen.cls.FlickerFreePrintPreviewControl();
        printDocument = new System.Drawing.Printing.PrintDocument();
        toolStrip.SuspendLayout();
        statusStrip.SuspendLayout();
        SuspendLayout();
        // 
        // toolStrip
        // 
        toolStrip.GripStyle = ToolStripGripStyle.Hidden;
        toolStrip.Items.AddRange(new ToolStripItem[] { lblDevice, cbPrinter, lblSource, cbSources, toolStripSeparator, btnZoom, toolStripSeparator1, splitBtnPaperSize, btnFontSize, btnShowNotes, toolStripSeparator2, btnPrint, toolStripSeparator3, btnClose });
        toolStrip.Location = new Point(0, 0);
        toolStrip.Name = "toolStrip";
        toolStrip.Size = new Size(890, 27);
        toolStrip.TabIndex = 0;
        toolStrip.Resize += ToolStrip_Resize;
        // 
        // lblDevice
        // 
        lblDevice.Font = new Font("Segoe UI", 10F);
        lblDevice.Name = "lblDevice";
        lblDevice.Size = new Size(13, 24);
        lblDevice.Text = " ";
        // 
        // cbPrinter
        // 
        cbPrinter.DropDownStyle = ComboBoxStyle.DropDownList;
        cbPrinter.Font = new Font("Segoe UI", 10F);
        cbPrinter.Name = "cbPrinter";
        cbPrinter.Size = new Size(170, 27);
        cbPrinter.ToolTipText = "Gerät";
        cbPrinter.SelectedIndexChanged += CbPrinter_SelectedIndexChanged;
        cbPrinter.Click += CbPrinter_Click;
        // 
        // lblSource
        // 
        lblSource.Font = new Font("Segoe UI", 10F);
        lblSource.Name = "lblSource";
        lblSource.Size = new Size(13, 24);
        lblSource.Text = " ";
        // 
        // cbSources
        // 
        cbSources.DropDownStyle = ComboBoxStyle.DropDownList;
        cbSources.Font = new Font("Segoe UI", 10F);
        cbSources.Name = "cbSources";
        cbSources.Size = new Size(170, 27);
        cbSources.ToolTipText = "Papierzufuhr";
        cbSources.SelectedIndexChanged += CbSources_SelectedIndexChanged;
        cbSources.Click += CbSources_Click;
        // 
        // toolStripSeparator
        // 
        toolStripSeparator.Name = "toolStripSeparator";
        toolStripSeparator.Size = new Size(6, 27);
        // 
        // btnZoom
        // 
        btnZoom.Font = new Font("Segoe UI", 10F);
        btnZoom.Image = Properties.Resources.ZoomHS16;
        btnZoom.ImageTransparentColor = Color.Magenta;
        btnZoom.Name = "btnZoom";
        btnZoom.Size = new Size(65, 24);
        btnZoom.Text = "Zoom";
        btnZoom.ToolTipText = "Strg+Z";
        btnZoom.Click += BtnZoom_Click;
        // 
        // toolStripSeparator1
        // 
        toolStripSeparator1.Name = "toolStripSeparator1";
        toolStripSeparator1.Size = new Size(6, 27);
        // 
        // splitBtnPaperSize
        // 
        splitBtnPaperSize.DropDownItems.AddRange(new ToolStripItem[] { splitItemA4, splitItemA5 });
        splitBtnPaperSize.Font = new Font("Segoe UI", 10F);
        splitBtnPaperSize.Image = Properties.Resources.docresize16;
        splitBtnPaperSize.ImageTransparentColor = Color.Magenta;
        splitBtnPaperSize.Name = "splitBtnPaperSize";
        splitBtnPaperSize.Size = new Size(85, 24);
        splitBtnPaperSize.Text = "Format";
        splitBtnPaperSize.ToolTipText = "Strg+D";
        splitBtnPaperSize.ButtonClick += SplitBtnPaperSize_ButtonClick;
        // 
        // splitItemA4
        // 
        splitItemA4.DisplayStyle = ToolStripItemDisplayStyle.Text;
        splitItemA4.Name = "splitItemA4";
        splitItemA4.ShortcutKeyDisplayString = "Strg+4";
        splitItemA4.Size = new Size(180, 24);
        splitItemA4.Text = "DIN A4";
        splitItemA4.Click += SplitItemA4_Click;
        // 
        // splitItemA5
        // 
        splitItemA5.DisplayStyle = ToolStripItemDisplayStyle.Text;
        splitItemA5.Name = "splitItemA5";
        splitItemA5.ShortcutKeyDisplayString = "Strg+5";
        splitItemA5.Size = new Size(180, 24);
        splitItemA5.Text = "DIN A5";
        splitItemA5.Click += SplitItemA5_Click;
        // 
        // btnFontSize
        // 
        btnFontSize.Font = new Font("Segoe UI", 10F);
        btnFontSize.Image = Properties.Resources.IncreaseFontSize16;
        btnFontSize.ImageTransparentColor = Color.Magenta;
        btnFontSize.Name = "btnFontSize";
        btnFontSize.Size = new Size(103, 24);
        btnFontSize.Text = "Schriftgröße";
        btnFontSize.ToolTipText = "Strg+G";
        btnFontSize.Click += BtnFontSize_Click;
        // 
        // btnShowNotes
        // 
        btnShowNotes.Font = new Font("Segoe UI", 10F);
        btnShowNotes.Image = Properties.Resources.stickynoteminus16;
        btnShowNotes.ImageTransparentColor = Color.Magenta;
        btnShowNotes.Name = "btnShowNotes";
        btnShowNotes.Size = new Size(76, 24);
        btnShowNotes.Text = "Notizen";
        btnShowNotes.ToolTipText = "Strg+N";
        btnShowNotes.Click += BtnShowNotes_Click;
        // 
        // toolStripSeparator2
        // 
        toolStripSeparator2.Name = "toolStripSeparator2";
        toolStripSeparator2.Size = new Size(6, 27);
        // 
        // btnPrint
        // 
        btnPrint.Font = new Font("Segoe UI", 10F);
        btnPrint.Image = Properties.Resources.printer16;
        btnPrint.ImageTransparentColor = Color.Magenta;
        btnPrint.Name = "btnPrint";
        btnPrint.Size = new Size(80, 24);
        btnPrint.Text = "Drucken";
        btnPrint.ToolTipText = "Enter, Strg+P";
        btnPrint.Click += BtnPrint_Click;
        // 
        // toolStripSeparator3
        // 
        toolStripSeparator3.Name = "toolStripSeparator3";
        toolStripSeparator3.Size = new Size(6, 27);
        // 
        // btnClose
        // 
        btnClose.Alignment = ToolStripItemAlignment.Right;
        btnClose.Font = new Font("Segoe UI", 10F);
        btnClose.Image = Properties.Resources.exit16;
        btnClose.ImageTransparentColor = Color.Magenta;
        btnClose.Name = "btnClose";
        btnClose.Size = new Size(86, 23);
        btnClose.Text = "Schließen";
        btnClose.ToolTipText = "Escape";
        btnClose.Click += BtnClose_Click;
        // 
        // statusStrip
        // 
        statusStrip.Items.AddRange(new ToolStripItem[] { toolStripStatusLabel });
        statusStrip.Location = new Point(0, 789);
        statusStrip.Name = "statusStrip";
        statusStrip.Size = new Size(890, 22);
        statusStrip.TabIndex = 2;
        statusStrip.Text = "statusStrip1";
        // 
        // toolStripStatusLabel
        // 
        toolStripStatusLabel.Name = "toolStripStatusLabel";
        toolStripStatusLabel.Size = new Size(875, 17);
        toolStripStatusLabel.Spring = true;
        toolStripStatusLabel.Text = "Zoom";
        // 
        // printPreviewControl
        // 
        printPreviewControl.AutoZoom = false;
        printPreviewControl.Dock = DockStyle.Fill;
        printPreviewControl.Document = printDocument;
        printPreviewControl.Location = new Point(0, 27);
        printPreviewControl.Name = "printPreviewControl";
        printPreviewControl.Size = new Size(890, 762);
        printPreviewControl.TabIndex = 3;
        printPreviewControl.UseAntiAlias = true;
        printPreviewControl.Zoom = 0.66894781864841746D;
        printPreviewControl.Paint += PrintPreviewControl_Paint;
        // 
        // printDocument
        // 
        printDocument.PrintPage += PrintDocument_PrintPage;
        // 
        // FrmSinglePrintPreview
        // 
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(890, 811);
        Controls.Add(printPreviewControl);
        Controls.Add(statusStrip);
        Controls.Add(toolStrip);
        Font = new Font("Segoe UI", 10F);
        Icon = (Icon)resources.GetObject("$this.Icon");
        MinimizeBox = false;
        MinimumSize = new Size(720, 700);
        Name = "FrmSinglePrintPreview";
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Druckvorschau";
        Load += FrmSinglePrintPreview_Load;
        toolStrip.ResumeLayout(false);
        toolStrip.PerformLayout();
        statusStrip.ResumeLayout(false);
        statusStrip.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private ToolStrip toolStrip;
    private ToolStripButton btnPrint;
    private ToolStripButton btnZoom;
    private ToolStripButton btnClose;
    private ToolStripSeparator toolStripSeparator;
    private StatusStrip statusStrip;
    private ToolStripStatusLabel toolStripStatusLabel;
    private ToolStripSeparator toolStripSeparator1;
    private ToolStripButton btnShowNotes;
    private ToolStripComboBox cbPrinter;
    private ToolStripComboBox cbSources;
    private ToolStripButton btnFontSize;
    private cls.FlickerFreePrintPreviewControl printPreviewControl;
    private System.Drawing.Printing.PrintDocument printDocument;
    private ToolStripSeparator toolStripSeparator2;
    private ToolStripLabel lblDevice;
    private ToolStripLabel lblSource;
    private ToolStripSeparator toolStripSeparator3;
    private ToolStripSplitButton splitBtnPaperSize;
    private ToolStripMenuItem splitItemA4;
    private ToolStripMenuItem splitItemA5;
}