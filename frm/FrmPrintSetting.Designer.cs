using Adressen.cls;

namespace Adressen;

partial class FrmPrintSetting
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
        components = new System.ComponentModel.Container();
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPrintSetting));
        tabControl = new TabControl();
        printerPage = new TabPage();
        gbOrientation = new GroupBox();
        picLandscape = new PictureBox();
        picPortrait = new PictureBox();
        rbLandscape = new RadioButton();
        rbPortrait = new RadioButton();
        gbPrinter = new GroupBox();
        lblSource = new Label();
        lblDevice = new Label();
        cbSources = new ComboBox();
        cbPrinter = new ComboBox();
        formatPage = new TabPage();
        gbText = new GroupBox();
        cbFontSizeRecipient = new ComboBox();
        lblRecipient = new Label();
        cbFontsizeSender = new ComboBox();
        lblFontsize = new Label();
        cbFont = new ComboBox();
        gbFormat = new GroupBox();
        cbPapersize = new ComboBox();
        senderPage = new TabPage();
        tcSender = new TabControl();
        tpSender1 = new TabPage();
        tbSender1 = new TextBox();
        tpSender2 = new TabPage();
        tbSender2 = new TextBox();
        tpSender3 = new TabPage();
        tbSender3 = new TextBox();
        tpSender4 = new TabPage();
        tbSender4 = new TextBox();
        tpSender5 = new TabPage();
        tbSender5 = new TextBox();
        tpSender6 = new TabPage();
        tbSender6 = new TextBox();
        ckbPrintSender = new CheckBox();
        recipientPage = new TabPage();
        lblLineFactor = new Label();
        lblLineHeight = new Label();
        nudLineHeightFactor = new NumericUpDown();
        lblLandRows = new Label();
        lblLandGapFactor = new Label();
        nudLandGapFactor = new NumericUpDown();
        lblZipRows = new Label();
        lblZipGapFactor = new Label();
        nudZipGapFactor = new NumericUpDown();
        ckbLandGROSS = new CheckBox();
        ckbAnredeOberhalb = new CheckBox();
        ckbAnredePrint = new CheckBox();
        ckbLandPrint = new CheckBox();
        tuningPage = new TabPage();
        lblHorizLine = new Label();
        ckbBoldSender = new CheckBox();
        ckbBoldRecipient = new CheckBox();
        lblBold = new Label();
        nudSenderOffsetY = new NumericUpDown();
        label1 = new Label();
        label2 = new Label();
        label3 = new Label();
        nudSenderOffsetX = new NumericUpDown();
        nudRecipOffsetY = new NumericUpDown();
        lblRecipOffsetY = new Label();
        lblRecipOffsetX = new Label();
        lblAddressOffset = new Label();
        nudRecipOffsetX = new NumericUpDown();
        btnSave = new Button();
        btnCancel = new Button();
        printDocument = new System.Drawing.Printing.PrintDocument();
        printPreviewControl = new FlickerFreePrintPreviewControl();
        contextMenuStrip = new ContextMenuStrip(components);
        zoomInToolStripMenuItem = new ToolStripMenuItem();
        zoomOutToolStripMenuItem = new ToolStripMenuItem();
        zoomDefaultToolStripMenuItem = new ToolStripMenuItem();
        statusStrip = new StatusStrip();
        lblZoomStatus = new ToolStripStatusLabel();
        toolStripStatusLabel = new ToolStripStatusLabel();
        rightPanel = new Panel();
        timerDebounce = new System.Windows.Forms.Timer(components);
        tabControl.SuspendLayout();
        printerPage.SuspendLayout();
        gbOrientation.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)picLandscape).BeginInit();
        ((System.ComponentModel.ISupportInitialize)picPortrait).BeginInit();
        gbPrinter.SuspendLayout();
        formatPage.SuspendLayout();
        gbText.SuspendLayout();
        gbFormat.SuspendLayout();
        senderPage.SuspendLayout();
        tcSender.SuspendLayout();
        tpSender1.SuspendLayout();
        tpSender2.SuspendLayout();
        tpSender3.SuspendLayout();
        tpSender4.SuspendLayout();
        tpSender5.SuspendLayout();
        tpSender6.SuspendLayout();
        recipientPage.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)nudLineHeightFactor).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudLandGapFactor).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudZipGapFactor).BeginInit();
        tuningPage.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)nudSenderOffsetY).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudSenderOffsetX).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudRecipOffsetY).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudRecipOffsetX).BeginInit();
        contextMenuStrip.SuspendLayout();
        statusStrip.SuspendLayout();
        rightPanel.SuspendLayout();
        SuspendLayout();
        // 
        // tabControl
        // 
        tabControl.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        tabControl.Controls.Add(printerPage);
        tabControl.Controls.Add(formatPage);
        tabControl.Controls.Add(senderPage);
        tabControl.Controls.Add(recipientPage);
        tabControl.Controls.Add(tuningPage);
        tabControl.Location = new Point(2, 3);
        tabControl.Name = "tabControl";
        tabControl.SelectedIndex = 0;
        tabControl.Size = new Size(324, 201);
        tabControl.TabIndex = 0;
        tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;
        // 
        // printerPage
        // 
        printerPage.BorderStyle = BorderStyle.FixedSingle;
        printerPage.Controls.Add(gbOrientation);
        printerPage.Controls.Add(gbPrinter);
        printerPage.Location = new Point(4, 26);
        printerPage.Name = "printerPage";
        printerPage.Padding = new Padding(3);
        printerPage.Size = new Size(316, 171);
        printerPage.TabIndex = 0;
        printerPage.Text = "Drucker";
        // 
        // gbOrientation
        // 
        gbOrientation.Controls.Add(picLandscape);
        gbOrientation.Controls.Add(picPortrait);
        gbOrientation.Controls.Add(rbLandscape);
        gbOrientation.Controls.Add(rbPortrait);
        gbOrientation.Location = new Point(8, 112);
        gbOrientation.Name = "gbOrientation";
        gbOrientation.Size = new Size(300, 54);
        gbOrientation.TabIndex = 1;
        gbOrientation.TabStop = false;
        gbOrientation.Text = "Ausrichtung";
        // 
        // picLandscape
        // 
        picLandscape.BackgroundImageLayout = ImageLayout.Center;
        picLandscape.Image = Properties.Resources.mail;
        picLandscape.Location = new Point(270, 23);
        picLandscape.Name = "picLandscape";
        picLandscape.Size = new Size(24, 24);
        picLandscape.TabIndex = 3;
        picLandscape.TabStop = false;
        picLandscape.Click += PicLandscape_Click;
        // 
        // picPortrait
        // 
        picPortrait.BackgroundImageLayout = ImageLayout.Center;
        picPortrait.Image = Properties.Resources.vertical;
        picPortrait.Location = new Point(106, 24);
        picPortrait.Name = "picPortrait";
        picPortrait.Size = new Size(24, 24);
        picPortrait.TabIndex = 2;
        picPortrait.TabStop = false;
        picPortrait.Click += PicPortrait_Click;
        // 
        // rbLandscape
        // 
        rbLandscape.AutoSize = true;
        rbLandscape.Checked = true;
        rbLandscape.Location = new Point(171, 24);
        rbLandscape.Name = "rbLandscape";
        rbLandscape.Size = new Size(99, 23);
        rbLandscape.TabIndex = 1;
        rbLandscape.TabStop = true;
        rbLandscape.Text = "Querformat";
        rbLandscape.UseVisualStyleBackColor = true;
        // 
        // rbPortrait
        // 
        rbPortrait.AutoSize = true;
        rbPortrait.Location = new Point(6, 24);
        rbPortrait.Name = "rbPortrait";
        rbPortrait.Size = new Size(100, 23);
        rbPortrait.TabIndex = 0;
        rbPortrait.Text = "Hochformat";
        rbPortrait.UseVisualStyleBackColor = true;
        rbPortrait.CheckedChanged += RbPortrait_CheckedChanged;
        // 
        // gbPrinter
        // 
        gbPrinter.Controls.Add(lblSource);
        gbPrinter.Controls.Add(lblDevice);
        gbPrinter.Controls.Add(cbSources);
        gbPrinter.Controls.Add(cbPrinter);
        gbPrinter.Location = new Point(8, 6);
        gbPrinter.Name = "gbPrinter";
        gbPrinter.Size = new Size(300, 100);
        gbPrinter.TabIndex = 0;
        gbPrinter.TabStop = false;
        gbPrinter.Text = "Drucker";
        // 
        // lblSource
        // 
        lblSource.AutoSize = true;
        lblSource.Location = new Point(6, 64);
        lblSource.Name = "lblSource";
        lblSource.Size = new Size(89, 19);
        lblSource.TabIndex = 3;
        lblSource.Text = "Papierzufuhr:";
        // 
        // lblDevice
        // 
        lblDevice.AutoSize = true;
        lblDevice.Location = new Point(6, 27);
        lblDevice.Name = "lblDevice";
        lblDevice.Size = new Size(46, 19);
        lblDevice.TabIndex = 2;
        lblDevice.Text = "Gerät:";
        // 
        // cbSources
        // 
        cbSources.DropDownStyle = ComboBoxStyle.DropDownList;
        cbSources.FormattingEnabled = true;
        cbSources.Location = new Point(101, 61);
        cbSources.Name = "cbSources";
        cbSources.Size = new Size(193, 25);
        cbSources.TabIndex = 1;
        cbSources.SelectedIndexChanged += CbSources_SelectedIndexChanged;
        // 
        // cbPrinter
        // 
        cbPrinter.DropDownStyle = ComboBoxStyle.DropDownList;
        cbPrinter.FormattingEnabled = true;
        cbPrinter.Location = new Point(101, 24);
        cbPrinter.Name = "cbPrinter";
        cbPrinter.Size = new Size(193, 25);
        cbPrinter.TabIndex = 0;
        cbPrinter.SelectedIndexChanged += CbPrinter_SelectedIndexChanged;
        // 
        // formatPage
        // 
        formatPage.BorderStyle = BorderStyle.FixedSingle;
        formatPage.Controls.Add(gbText);
        formatPage.Controls.Add(gbFormat);
        formatPage.Location = new Point(4, 24);
        formatPage.Name = "formatPage";
        formatPage.Padding = new Padding(3);
        formatPage.Size = new Size(316, 173);
        formatPage.TabIndex = 1;
        formatPage.Text = "Format";
        // 
        // gbText
        // 
        gbText.Controls.Add(cbFontSizeRecipient);
        gbText.Controls.Add(lblRecipient);
        gbText.Controls.Add(cbFontsizeSender);
        gbText.Controls.Add(lblFontsize);
        gbText.Controls.Add(cbFont);
        gbText.Location = new Point(8, 72);
        gbText.Name = "gbText";
        gbText.Size = new Size(300, 96);
        gbText.TabIndex = 1;
        gbText.TabStop = false;
        gbText.Text = "Schriftart und Schriftgröße";
        // 
        // cbFontSizeRecipient
        // 
        cbFontSizeRecipient.DropDownStyle = ComboBoxStyle.DropDownList;
        cbFontSizeRecipient.FormattingEnabled = true;
        cbFontSizeRecipient.Items.AddRange(new object[] { "10", "12", "14", "16", "18", "20", "22", "24" });
        cbFontSizeRecipient.Location = new Point(234, 62);
        cbFontSizeRecipient.Name = "cbFontSizeRecipient";
        cbFontSizeRecipient.Size = new Size(60, 25);
        cbFontSizeRecipient.TabIndex = 5;
        cbFontSizeRecipient.SelectedIndexChanged += GenericControl_ValueChanged;
        // 
        // lblRecipient
        // 
        lblRecipient.AutoSize = true;
        lblRecipient.Location = new Point(150, 65);
        lblRecipient.Name = "lblRecipient";
        lblRecipient.Size = new Size(78, 19);
        lblRecipient.TabIndex = 4;
        lblRecipient.Text = "Empfänger:";
        // 
        // cbFontsizeSender
        // 
        cbFontsizeSender.DropDownStyle = ComboBoxStyle.DropDownList;
        cbFontsizeSender.FormattingEnabled = true;
        cbFontsizeSender.Items.AddRange(new object[] { "10", "12", "14", "16", "18", "20", "22", "24" });
        cbFontsizeSender.Location = new Point(84, 62);
        cbFontsizeSender.Name = "cbFontsizeSender";
        cbFontsizeSender.Size = new Size(60, 25);
        cbFontsizeSender.TabIndex = 3;
        cbFontsizeSender.SelectedIndexChanged += GenericControl_ValueChanged;
        // 
        // lblFontsize
        // 
        lblFontsize.AutoSize = true;
        lblFontsize.Location = new Point(6, 65);
        lblFontsize.Name = "lblFontsize";
        lblFontsize.Size = new Size(70, 19);
        lblFontsize.TabIndex = 2;
        lblFontsize.Text = "Absender:";
        // 
        // cbFont
        // 
        cbFont.DropDownStyle = ComboBoxStyle.DropDownList;
        cbFont.FormattingEnabled = true;
        cbFont.Location = new Point(6, 24);
        cbFont.Name = "cbFont";
        cbFont.Size = new Size(288, 25);
        cbFont.TabIndex = 0;
        cbFont.SelectedIndexChanged += GenericControl_ValueChanged;
        // 
        // gbFormat
        // 
        gbFormat.Controls.Add(cbPapersize);
        gbFormat.Location = new Point(8, 6);
        gbFormat.Name = "gbFormat";
        gbFormat.Size = new Size(300, 60);
        gbFormat.TabIndex = 0;
        gbFormat.TabStop = false;
        gbFormat.Text = "Format (Höhe × Breite)";
        // 
        // cbPapersize
        // 
        cbPapersize.DropDownStyle = ComboBoxStyle.DropDownList;
        cbPapersize.FormattingEnabled = true;
        cbPapersize.Location = new Point(6, 24);
        cbPapersize.Name = "cbPapersize";
        cbPapersize.Size = new Size(288, 25);
        cbPapersize.TabIndex = 0;
        cbPapersize.SelectedIndexChanged += CbPapersize_SelectedIndexChanged;
        // 
        // senderPage
        // 
        senderPage.BorderStyle = BorderStyle.FixedSingle;
        senderPage.Controls.Add(tcSender);
        senderPage.Controls.Add(ckbPrintSender);
        senderPage.Location = new Point(4, 24);
        senderPage.Name = "senderPage";
        senderPage.Size = new Size(316, 173);
        senderPage.TabIndex = 3;
        senderPage.Text = "Absender";
        // 
        // tcSender
        // 
        tcSender.Alignment = TabAlignment.Left;
        tcSender.Controls.Add(tpSender1);
        tcSender.Controls.Add(tpSender2);
        tcSender.Controls.Add(tpSender3);
        tcSender.Controls.Add(tpSender4);
        tcSender.Controls.Add(tpSender5);
        tcSender.Controls.Add(tpSender6);
        tcSender.Dock = DockStyle.Top;
        tcSender.DrawMode = TabDrawMode.OwnerDrawFixed;
        tcSender.ItemSize = new Size(23, 23);
        tcSender.Location = new Point(0, 0);
        tcSender.Multiline = true;
        tcSender.Name = "tcSender";
        tcSender.SelectedIndex = 0;
        tcSender.Size = new Size(314, 143);
        tcSender.SizeMode = TabSizeMode.Fixed;
        tcSender.TabIndex = 2;
        tcSender.DrawItem += TcSender_DrawItem;
        tcSender.SelectedIndexChanged += GenericControl_ValueChanged;
        // 
        // tpSender1
        // 
        tpSender1.Controls.Add(tbSender1);
        tpSender1.Location = new Point(27, 4);
        tpSender1.Name = "tpSender1";
        tpSender1.Padding = new Padding(3);
        tpSender1.Size = new Size(283, 135);
        tpSender1.TabIndex = 0;
        tpSender1.Text = "1";
        tpSender1.UseVisualStyleBackColor = true;
        // 
        // tbSender1
        // 
        tbSender1.AcceptsReturn = true;
        tbSender1.AcceptsTab = true;
        tbSender1.BackColor = SystemColors.Info;
        tbSender1.Dock = DockStyle.Fill;
        tbSender1.ForeColor = SystemColors.InfoText;
        tbSender1.Location = new Point(3, 3);
        tbSender1.Multiline = true;
        tbSender1.Name = "tbSender1";
        tbSender1.Size = new Size(277, 129);
        tbSender1.TabIndex = 0;
        tbSender1.TextChanged += TbSender_TextChanged;
        // 
        // tpSender2
        // 
        tpSender2.Controls.Add(tbSender2);
        tpSender2.Location = new Point(27, 4);
        tpSender2.Name = "tpSender2";
        tpSender2.Padding = new Padding(3);
        tpSender2.Size = new Size(283, 135);
        tpSender2.TabIndex = 1;
        tpSender2.Text = "2";
        tpSender2.UseVisualStyleBackColor = true;
        // 
        // tbSender2
        // 
        tbSender2.AcceptsReturn = true;
        tbSender2.AcceptsTab = true;
        tbSender2.BackColor = SystemColors.Info;
        tbSender2.Dock = DockStyle.Fill;
        tbSender2.ForeColor = SystemColors.InfoText;
        tbSender2.Location = new Point(3, 3);
        tbSender2.Multiline = true;
        tbSender2.Name = "tbSender2";
        tbSender2.Size = new Size(277, 129);
        tbSender2.TabIndex = 1;
        tbSender2.TextChanged += TbSender_TextChanged;
        // 
        // tpSender3
        // 
        tpSender3.Controls.Add(tbSender3);
        tpSender3.Location = new Point(27, 4);
        tpSender3.Name = "tpSender3";
        tpSender3.Padding = new Padding(3);
        tpSender3.Size = new Size(283, 135);
        tpSender3.TabIndex = 2;
        tpSender3.Text = "3";
        tpSender3.UseVisualStyleBackColor = true;
        // 
        // tbSender3
        // 
        tbSender3.AcceptsReturn = true;
        tbSender3.AcceptsTab = true;
        tbSender3.BackColor = SystemColors.Info;
        tbSender3.Dock = DockStyle.Fill;
        tbSender3.ForeColor = SystemColors.InfoText;
        tbSender3.Location = new Point(3, 3);
        tbSender3.Multiline = true;
        tbSender3.Name = "tbSender3";
        tbSender3.Size = new Size(277, 129);
        tbSender3.TabIndex = 2;
        tbSender3.TextChanged += TbSender_TextChanged;
        // 
        // tpSender4
        // 
        tpSender4.Controls.Add(tbSender4);
        tpSender4.Location = new Point(27, 4);
        tpSender4.Name = "tpSender4";
        tpSender4.Padding = new Padding(3);
        tpSender4.Size = new Size(283, 135);
        tpSender4.TabIndex = 3;
        tpSender4.Text = "4";
        tpSender4.UseVisualStyleBackColor = true;
        // 
        // tbSender4
        // 
        tbSender4.AcceptsReturn = true;
        tbSender4.AcceptsTab = true;
        tbSender4.BackColor = SystemColors.Info;
        tbSender4.Dock = DockStyle.Fill;
        tbSender4.ForeColor = SystemColors.InfoText;
        tbSender4.Location = new Point(3, 3);
        tbSender4.Multiline = true;
        tbSender4.Name = "tbSender4";
        tbSender4.Size = new Size(277, 129);
        tbSender4.TabIndex = 5;
        tbSender4.TextChanged += TbSender_TextChanged;
        // 
        // tpSender5
        // 
        tpSender5.Controls.Add(tbSender5);
        tpSender5.Location = new Point(27, 4);
        tpSender5.Name = "tpSender5";
        tpSender5.Padding = new Padding(3);
        tpSender5.Size = new Size(283, 135);
        tpSender5.TabIndex = 4;
        tpSender5.Text = "5";
        tpSender5.UseVisualStyleBackColor = true;
        // 
        // tbSender5
        // 
        tbSender5.AcceptsReturn = true;
        tbSender5.AcceptsTab = true;
        tbSender5.BackColor = SystemColors.Info;
        tbSender5.Dock = DockStyle.Fill;
        tbSender5.ForeColor = SystemColors.InfoText;
        tbSender5.Location = new Point(3, 3);
        tbSender5.Multiline = true;
        tbSender5.Name = "tbSender5";
        tbSender5.Size = new Size(277, 129);
        tbSender5.TabIndex = 4;
        tbSender5.TextChanged += TbSender_TextChanged;
        // 
        // tpSender6
        // 
        tpSender6.Controls.Add(tbSender6);
        tpSender6.Location = new Point(27, 4);
        tpSender6.Name = "tpSender6";
        tpSender6.Padding = new Padding(3);
        tpSender6.Size = new Size(283, 135);
        tpSender6.TabIndex = 5;
        tpSender6.Text = "6";
        tpSender6.UseVisualStyleBackColor = true;
        // 
        // tbSender6
        // 
        tbSender6.AcceptsReturn = true;
        tbSender6.AcceptsTab = true;
        tbSender6.BackColor = SystemColors.Info;
        tbSender6.Dock = DockStyle.Fill;
        tbSender6.ForeColor = SystemColors.InfoText;
        tbSender6.Location = new Point(3, 3);
        tbSender6.Multiline = true;
        tbSender6.Name = "tbSender6";
        tbSender6.Size = new Size(277, 129);
        tbSender6.TabIndex = 3;
        tbSender6.TextChanged += TbSender_TextChanged;
        // 
        // ckbPrintSender
        // 
        ckbPrintSender.AutoSize = true;
        ckbPrintSender.Checked = true;
        ckbPrintSender.CheckState = CheckState.Checked;
        ckbPrintSender.Location = new Point(32, 146);
        ckbPrintSender.Name = "ckbPrintSender";
        ckbPrintSender.Size = new Size(274, 23);
        ckbPrintSender.TabIndex = 1;
        ckbPrintSender.Text = "Absendertext auf Briefumschlag drucken";
        ckbPrintSender.UseVisualStyleBackColor = true;
        ckbPrintSender.CheckedChanged += GenericControl_ValueChanged;
        // 
        // recipientPage
        // 
        recipientPage.Controls.Add(lblLineFactor);
        recipientPage.Controls.Add(lblLineHeight);
        recipientPage.Controls.Add(nudLineHeightFactor);
        recipientPage.Controls.Add(lblLandRows);
        recipientPage.Controls.Add(lblLandGapFactor);
        recipientPage.Controls.Add(nudLandGapFactor);
        recipientPage.Controls.Add(lblZipRows);
        recipientPage.Controls.Add(lblZipGapFactor);
        recipientPage.Controls.Add(nudZipGapFactor);
        recipientPage.Controls.Add(ckbLandGROSS);
        recipientPage.Controls.Add(ckbAnredeOberhalb);
        recipientPage.Controls.Add(ckbAnredePrint);
        recipientPage.Controls.Add(ckbLandPrint);
        recipientPage.Location = new Point(4, 24);
        recipientPage.Name = "recipientPage";
        recipientPage.Size = new Size(316, 173);
        recipientPage.TabIndex = 4;
        recipientPage.Text = "Empfänger";
        recipientPage.UseVisualStyleBackColor = true;
        // 
        // lblLineFactor
        // 
        lblLineFactor.AutoSize = true;
        lblLineFactor.Location = new Point(239, 76);
        lblLineFactor.Name = "lblLineFactor";
        lblLineFactor.Size = new Size(45, 19);
        lblLineFactor.TabIndex = 30;
        lblLineFactor.Text = "Zeilen";
        // 
        // lblLineHeight
        // 
        lblLineHeight.AutoSize = true;
        lblLineHeight.Location = new Point(8, 76);
        lblLineHeight.Name = "lblLineHeight";
        lblLineHeight.Size = new Size(163, 19);
        lblLineHeight.TabIndex = 29;
        lblLineHeight.Text = "Genereller Zeilenabstand:";
        // 
        // nudLineHeightFactor
        // 
        nudLineHeightFactor.DecimalPlaces = 2;
        nudLineHeightFactor.Increment = new decimal(new int[] { 5, 0, 0, 131072 });
        nudLineHeightFactor.Location = new Point(178, 74);
        nudLineHeightFactor.Maximum = new decimal(new int[] { 30, 0, 0, 65536 });
        nudLineHeightFactor.Minimum = new decimal(new int[] { 5, 0, 0, 65536 });
        nudLineHeightFactor.Name = "nudLineHeightFactor";
        nudLineHeightFactor.Size = new Size(55, 25);
        nudLineHeightFactor.TabIndex = 28;
        nudLineHeightFactor.TextAlign = HorizontalAlignment.Center;
        nudLineHeightFactor.Value = new decimal(new int[] { 15, 0, 0, 65536 });
        nudLineHeightFactor.ValueChanged += GenericControl_ValueChanged;
        // 
        // lblLandRows
        // 
        lblLandRows.AutoSize = true;
        lblLandRows.Location = new Point(239, 142);
        lblLandRows.Name = "lblLandRows";
        lblLandRows.Size = new Size(45, 19);
        lblLandRows.TabIndex = 27;
        lblLandRows.Text = "Zeilen";
        // 
        // lblLandGapFactor
        // 
        lblLandGapFactor.AutoSize = true;
        lblLandGapFactor.Location = new Point(8, 142);
        lblLandGapFactor.Name = "lblLandGapFactor";
        lblLandGapFactor.Size = new Size(130, 19);
        lblLandGapFactor.TabIndex = 26;
        lblLandGapFactor.Text = "Zusätzlich vor Land:";
        // 
        // nudLandGapFactor
        // 
        nudLandGapFactor.DecimalPlaces = 2;
        nudLandGapFactor.Increment = new decimal(new int[] { 5, 0, 0, 131072 });
        nudLandGapFactor.Location = new Point(178, 140);
        nudLandGapFactor.Maximum = new decimal(new int[] { 2, 0, 0, 0 });
        nudLandGapFactor.Name = "nudLandGapFactor";
        nudLandGapFactor.Size = new Size(55, 25);
        nudLandGapFactor.TabIndex = 25;
        nudLandGapFactor.TextAlign = HorizontalAlignment.Center;
        nudLandGapFactor.Value = new decimal(new int[] { 3, 0, 0, 65536 });
        nudLandGapFactor.ValueChanged += GenericControl_ValueChanged;
        // 
        // lblZipRows
        // 
        lblZipRows.AutoSize = true;
        lblZipRows.Location = new Point(239, 109);
        lblZipRows.Name = "lblZipRows";
        lblZipRows.Size = new Size(45, 19);
        lblZipRows.TabIndex = 24;
        lblZipRows.Text = "Zeilen";
        // 
        // lblZipGapFactor
        // 
        lblZipGapFactor.AutoSize = true;
        lblZipGapFactor.Location = new Point(8, 109);
        lblZipGapFactor.Name = "lblZipGapFactor";
        lblZipGapFactor.Size = new Size(149, 19);
        lblZipGapFactor.TabIndex = 23;
        lblZipGapFactor.Text = "Zusätzlich vor PLZ/Ort:";
        // 
        // nudZipGapFactor
        // 
        nudZipGapFactor.DecimalPlaces = 2;
        nudZipGapFactor.Increment = new decimal(new int[] { 5, 0, 0, 131072 });
        nudZipGapFactor.Location = new Point(178, 107);
        nudZipGapFactor.Maximum = new decimal(new int[] { 2, 0, 0, 0 });
        nudZipGapFactor.Name = "nudZipGapFactor";
        nudZipGapFactor.Size = new Size(55, 25);
        nudZipGapFactor.TabIndex = 22;
        nudZipGapFactor.TextAlign = HorizontalAlignment.Center;
        nudZipGapFactor.Value = new decimal(new int[] { 3, 0, 0, 65536 });
        nudZipGapFactor.ValueChanged += GenericControl_ValueChanged;
        // 
        // ckbLandGROSS
        // 
        ckbLandGROSS.AutoSize = true;
        ckbLandGROSS.Checked = true;
        ckbLandGROSS.CheckState = CheckState.Checked;
        ckbLandGROSS.Location = new Point(178, 40);
        ckbLandGROSS.Name = "ckbLandGROSS";
        ckbLandGROSS.Size = new Size(130, 23);
        ckbLandGROSS.TabIndex = 21;
        ckbLandGROSS.Text = "Großbuchstaben";
        ckbLandGROSS.UseVisualStyleBackColor = true;
        ckbLandGROSS.CheckedChanged += GenericControl_ValueChanged;
        // 
        // ckbAnredeOberhalb
        // 
        ckbAnredeOberhalb.AutoSize = true;
        ckbAnredeOberhalb.Enabled = false;
        ckbAnredeOberhalb.Location = new Point(178, 9);
        ckbAnredeOberhalb.Name = "ckbAnredeOberhalb";
        ckbAnredeOberhalb.Size = new Size(85, 23);
        ckbAnredeOberhalb.TabIndex = 20;
        ckbAnredeOberhalb.Text = "Oberhalb";
        ckbAnredeOberhalb.UseVisualStyleBackColor = true;
        ckbAnredeOberhalb.CheckedChanged += GenericControl_ValueChanged;
        // 
        // ckbAnredePrint
        // 
        ckbAnredePrint.AutoSize = true;
        ckbAnredePrint.Location = new Point(12, 9);
        ckbAnredePrint.Name = "ckbAnredePrint";
        ckbAnredePrint.Size = new Size(144, 23);
        ckbAnredePrint.TabIndex = 17;
        ckbAnredePrint.Text = "Anrede hinzufügen";
        ckbAnredePrint.UseVisualStyleBackColor = true;
        ckbAnredePrint.CheckedChanged += GenericControl_ValueChanged;
        // 
        // ckbLandPrint
        // 
        ckbLandPrint.AutoSize = true;
        ckbLandPrint.Checked = true;
        ckbLandPrint.CheckState = CheckState.Checked;
        ckbLandPrint.Location = new Point(12, 40);
        ckbLandPrint.Name = "ckbLandPrint";
        ckbLandPrint.Size = new Size(130, 23);
        ckbLandPrint.TabIndex = 18;
        ckbLandPrint.Text = "Land hinzufügen";
        ckbLandPrint.UseVisualStyleBackColor = true;
        ckbLandPrint.CheckedChanged += GenericControl_ValueChanged;
        // 
        // tuningPage
        // 
        tuningPage.BorderStyle = BorderStyle.FixedSingle;
        tuningPage.Controls.Add(lblHorizLine);
        tuningPage.Controls.Add(ckbBoldSender);
        tuningPage.Controls.Add(ckbBoldRecipient);
        tuningPage.Controls.Add(lblBold);
        tuningPage.Controls.Add(nudSenderOffsetY);
        tuningPage.Controls.Add(label1);
        tuningPage.Controls.Add(label2);
        tuningPage.Controls.Add(label3);
        tuningPage.Controls.Add(nudSenderOffsetX);
        tuningPage.Controls.Add(nudRecipOffsetY);
        tuningPage.Controls.Add(lblRecipOffsetY);
        tuningPage.Controls.Add(lblRecipOffsetX);
        tuningPage.Controls.Add(lblAddressOffset);
        tuningPage.Controls.Add(nudRecipOffsetX);
        tuningPage.Location = new Point(4, 24);
        tuningPage.Name = "tuningPage";
        tuningPage.Size = new Size(316, 173);
        tuningPage.TabIndex = 2;
        tuningPage.Text = "Tuning";
        // 
        // lblHorizLine
        // 
        lblHorizLine.BorderStyle = BorderStyle.Fixed3D;
        lblHorizLine.Location = new Point(3, 126);
        lblHorizLine.Name = "lblHorizLine";
        lblHorizLine.Size = new Size(310, 2);
        lblHorizLine.TabIndex = 13;
        // 
        // ckbBoldSender
        // 
        ckbBoldSender.AutoSize = true;
        ckbBoldSender.Location = new Point(227, 136);
        ckbBoldSender.Name = "ckbBoldSender";
        ckbBoldSender.Size = new Size(86, 23);
        ckbBoldSender.TabIndex = 12;
        ckbBoldSender.Text = "Absender";
        ckbBoldSender.UseVisualStyleBackColor = true;
        ckbBoldSender.CheckedChanged += GenericControl_ValueChanged;
        // 
        // ckbBoldRecipient
        // 
        ckbBoldRecipient.AutoSize = true;
        ckbBoldRecipient.Location = new Point(132, 136);
        ckbBoldRecipient.Name = "ckbBoldRecipient";
        ckbBoldRecipient.Size = new Size(94, 23);
        ckbBoldRecipient.TabIndex = 11;
        ckbBoldRecipient.Text = "Empfänger";
        ckbBoldRecipient.UseVisualStyleBackColor = true;
        ckbBoldRecipient.CheckedChanged += GenericControl_ValueChanged;
        // 
        // lblBold
        // 
        lblBold.AutoSize = true;
        lblBold.Location = new Point(8, 138);
        lblBold.Name = "lblBold";
        lblBold.Size = new Size(114, 19);
        lblBold.TabIndex = 10;
        lblBold.Text = "Text fett drucken:";
        // 
        // nudSenderOffsetY
        // 
        nudSenderOffsetY.Location = new Point(253, 88);
        nudSenderOffsetY.Minimum = new decimal(new int[] { 20, 0, 0, int.MinValue });
        nudSenderOffsetY.Name = "nudSenderOffsetY";
        nudSenderOffsetY.Size = new Size(55, 25);
        nudSenderOffsetY.TabIndex = 9;
        nudSenderOffsetY.TextAlign = HorizontalAlignment.Center;
        nudSenderOffsetY.ValueChanged += GenericControl_ValueChanged;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(191, 90);
        label1.Name = "label1";
        label1.Size = new Size(56, 19);
        label1.TabIndex = 8;
        label1.Text = "vertikal:";
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(9, 90);
        label2.Name = "label2";
        label2.Size = new Size(114, 19);
        label2.TabIndex = 7;
        label2.Text = "Offset horizontal:";
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(8, 68);
        label3.Name = "label3";
        label3.Size = new Size(298, 19);
        label3.TabIndex = 6;
        label3.Text = "Die Absenderadresse wird links oben eingefügt.";
        // 
        // nudSenderOffsetX
        // 
        nudSenderOffsetX.Location = new Point(129, 88);
        nudSenderOffsetX.Minimum = new decimal(new int[] { 20, 0, 0, int.MinValue });
        nudSenderOffsetX.Name = "nudSenderOffsetX";
        nudSenderOffsetX.Size = new Size(55, 25);
        nudSenderOffsetX.TabIndex = 5;
        nudSenderOffsetX.TextAlign = HorizontalAlignment.Center;
        nudSenderOffsetX.ValueChanged += GenericControl_ValueChanged;
        // 
        // nudRecipOffsetY
        // 
        nudRecipOffsetY.Location = new Point(253, 27);
        nudRecipOffsetY.Maximum = new decimal(new int[] { 200, 0, 0, 0 });
        nudRecipOffsetY.Minimum = new decimal(new int[] { 200, 0, 0, int.MinValue });
        nudRecipOffsetY.Name = "nudRecipOffsetY";
        nudRecipOffsetY.Size = new Size(55, 25);
        nudRecipOffsetY.TabIndex = 4;
        nudRecipOffsetY.TextAlign = HorizontalAlignment.Center;
        nudRecipOffsetY.ValueChanged += GenericControl_ValueChanged;
        // 
        // lblRecipOffsetY
        // 
        lblRecipOffsetY.AutoSize = true;
        lblRecipOffsetY.Location = new Point(191, 29);
        lblRecipOffsetY.Name = "lblRecipOffsetY";
        lblRecipOffsetY.Size = new Size(56, 19);
        lblRecipOffsetY.TabIndex = 3;
        lblRecipOffsetY.Text = "vertikal:";
        // 
        // lblRecipOffsetX
        // 
        lblRecipOffsetX.AutoSize = true;
        lblRecipOffsetX.Location = new Point(9, 29);
        lblRecipOffsetX.Name = "lblRecipOffsetX";
        lblRecipOffsetX.Size = new Size(114, 19);
        lblRecipOffsetX.TabIndex = 2;
        lblRecipOffsetX.Text = "Offset horizontal:";
        // 
        // lblAddressOffset
        // 
        lblAddressOffset.AutoSize = true;
        lblAddressOffset.Location = new Point(8, 7);
        lblAddressOffset.Name = "lblAddressOffset";
        lblAddressOffset.Size = new Size(280, 19);
        lblAddressOffset.TabIndex = 1;
        lblAddressOffset.Text = "Die Empfängeradresse wird mittig eingefügt.";
        // 
        // nudRecipOffsetX
        // 
        nudRecipOffsetX.Location = new Point(129, 27);
        nudRecipOffsetX.Maximum = new decimal(new int[] { 200, 0, 0, 0 });
        nudRecipOffsetX.Minimum = new decimal(new int[] { 200, 0, 0, int.MinValue });
        nudRecipOffsetX.Name = "nudRecipOffsetX";
        nudRecipOffsetX.Size = new Size(55, 25);
        nudRecipOffsetX.TabIndex = 0;
        nudRecipOffsetX.TextAlign = HorizontalAlignment.Center;
        nudRecipOffsetX.ValueChanged += GenericControl_ValueChanged;
        // 
        // btnSave
        // 
        btnSave.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnSave.DialogResult = DialogResult.OK;
        btnSave.Image = Properties.Resources.printer24;
        btnSave.Location = new Point(6, 206);
        btnSave.Name = "btnSave";
        btnSave.Size = new Size(166, 32);
        btnSave.TabIndex = 1;
        btnSave.Text = " Umschlag drucken";
        btnSave.TextAlign = ContentAlignment.MiddleRight;
        btnSave.TextImageRelation = TextImageRelation.ImageBeforeText;
        btnSave.UseVisualStyleBackColor = true;
        btnSave.Click += BtnSave_Click;
        // 
        // btnCancel
        // 
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnCancel.DialogResult = DialogResult.OK;
        btnCancel.Location = new Point(178, 207);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(144, 32);
        btnCancel.TabIndex = 2;
        btnCancel.Text = "Speichern/Schließen";
        btnCancel.TextImageRelation = TextImageRelation.ImageBeforeText;
        btnCancel.UseVisualStyleBackColor = true;
        // 
        // printDocument
        // 
        printDocument.BeginPrint += PrintDocument_BeginPrint;
        printDocument.PrintPage += PrintDocument_PrintPage;
        // 
        // printPreviewControl
        // 
        printPreviewControl.AutoZoom = false;
        printPreviewControl.ContextMenuStrip = contextMenuStrip;
        printPreviewControl.Dock = DockStyle.Fill;
        printPreviewControl.Location = new Point(0, 0);
        printPreviewControl.Name = "printPreviewControl";
        printPreviewControl.Size = new Size(440, 242);
        printPreviewControl.TabIndex = 1;
        printPreviewControl.UseAntiAlias = true;
        printPreviewControl.Zoom = 0.5D;
        printPreviewControl.ZoomChanged += PrintPreviewControl_ZoomChanged;
        printPreviewControl.DoubleClick += PrintPreviewControl_DoubleClick;
        // 
        // contextMenuStrip
        // 
        contextMenuStrip.Items.AddRange(new ToolStripItem[] { zoomInToolStripMenuItem, zoomOutToolStripMenuItem, zoomDefaultToolStripMenuItem });
        contextMenuStrip.Name = "contextMenuStrip";
        contextMenuStrip.Size = new Size(194, 70);
        contextMenuStrip.Opening += ContextMenuStrip_Opening;
        // 
        // zoomInToolStripMenuItem
        // 
        zoomInToolStripMenuItem.Image = Properties.Resources.ZoomIn16;
        zoomInToolStripMenuItem.Name = "zoomInToolStripMenuItem";
        zoomInToolStripMenuItem.ShortcutKeyDisplayString = "Strg+＋";
        zoomInToolStripMenuItem.Size = new Size(193, 22);
        zoomInToolStripMenuItem.Text = "Vergrößern";
        zoomInToolStripMenuItem.Click += ZoomInToolStripMenuItem_Click;
        // 
        // zoomOutToolStripMenuItem
        // 
        zoomOutToolStripMenuItem.Image = Properties.Resources.ZoomOut16;
        zoomOutToolStripMenuItem.Name = "zoomOutToolStripMenuItem";
        zoomOutToolStripMenuItem.ShortcutKeyDisplayString = "Strg+‒";
        zoomOutToolStripMenuItem.Size = new Size(193, 22);
        zoomOutToolStripMenuItem.Text = "Verkleinern";
        zoomOutToolStripMenuItem.Click += ZoomOutToolStripMenuItem_Click;
        // 
        // zoomDefaultToolStripMenuItem
        // 
        zoomDefaultToolStripMenuItem.Image = Properties.Resources.ZoomToFit16;
        zoomDefaultToolStripMenuItem.Name = "zoomDefaultToolStripMenuItem";
        zoomDefaultToolStripMenuItem.ShortcutKeyDisplayString = "Strg+0";
        zoomDefaultToolStripMenuItem.Size = new Size(193, 22);
        zoomDefaultToolStripMenuItem.Text = "Standardzoom";
        zoomDefaultToolStripMenuItem.Click += ZoomDefaultToolStripMenuItem_Click;
        // 
        // statusStrip
        // 
        statusStrip.Font = new Font("Segoe UI", 10F);
        statusStrip.GripStyle = ToolStripGripStyle.Visible;
        statusStrip.Items.AddRange(new ToolStripItem[] { lblZoomStatus, toolStripStatusLabel });
        statusStrip.Location = new Point(0, 242);
        statusStrip.Name = "statusStrip";
        statusStrip.Size = new Size(769, 24);
        statusStrip.TabIndex = 3;
        statusStrip.Text = "statusStrip";
        // 
        // lblZoomStatus
        // 
        lblZoomStatus.AutoSize = false;
        lblZoomStatus.BorderSides = ToolStripStatusLabelBorderSides.Right;
        lblZoomStatus.Name = "lblZoomStatus";
        lblZoomStatus.Size = new Size(440, 19);
        lblZoomStatus.Text = "Zoom: 50%";
        lblZoomStatus.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // toolStripStatusLabel
        // 
        toolStripStatusLabel.Font = new Font("Segoe UI", 10F);
        toolStripStatusLabel.Name = "toolStripStatusLabel";
        toolStripStatusLabel.Size = new Size(314, 19);
        toolStripStatusLabel.Spring = true;
        toolStripStatusLabel.Text = "Drucker";
        toolStripStatusLabel.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // rightPanel
        // 
        rightPanel.Controls.Add(tabControl);
        rightPanel.Controls.Add(btnCancel);
        rightPanel.Controls.Add(btnSave);
        rightPanel.Dock = DockStyle.Right;
        rightPanel.Location = new Point(440, 0);
        rightPanel.Name = "rightPanel";
        rightPanel.Size = new Size(329, 242);
        rightPanel.TabIndex = 4;
        // 
        // timerDebounce
        // 
        timerDebounce.Interval = 1000;
        timerDebounce.Tick += TimerDebounce_Tick;
        // 
        // FrmPrintSetting
        // 
        AcceptButton = btnSave;
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        CancelButton = btnCancel;
        ClientSize = new Size(769, 266);
        Controls.Add(printPreviewControl);
        Controls.Add(rightPanel);
        Controls.Add(statusStrip);
        Font = new Font("Segoe UI", 10F);
        Icon = (Icon)resources.GetObject("$this.Icon");
        MaximizeBox = false;
        MinimizeBox = false;
        MinimumSize = new Size(785, 305);
        Name = "FrmPrintSetting";
        ShowInTaskbar = false;
        SizeGripStyle = SizeGripStyle.Show;
        StartPosition = FormStartPosition.CenterParent;
        Text = "Adresse auf Briefumschlag drucken";
        Load += FrmPrintSetting_Load;
        Layout += FrmPrintSetting_Layout;
        tabControl.ResumeLayout(false);
        printerPage.ResumeLayout(false);
        gbOrientation.ResumeLayout(false);
        gbOrientation.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)picLandscape).EndInit();
        ((System.ComponentModel.ISupportInitialize)picPortrait).EndInit();
        gbPrinter.ResumeLayout(false);
        gbPrinter.PerformLayout();
        formatPage.ResumeLayout(false);
        gbText.ResumeLayout(false);
        gbText.PerformLayout();
        gbFormat.ResumeLayout(false);
        senderPage.ResumeLayout(false);
        senderPage.PerformLayout();
        tcSender.ResumeLayout(false);
        tpSender1.ResumeLayout(false);
        tpSender1.PerformLayout();
        tpSender2.ResumeLayout(false);
        tpSender2.PerformLayout();
        tpSender3.ResumeLayout(false);
        tpSender3.PerformLayout();
        tpSender4.ResumeLayout(false);
        tpSender4.PerformLayout();
        tpSender5.ResumeLayout(false);
        tpSender5.PerformLayout();
        tpSender6.ResumeLayout(false);
        tpSender6.PerformLayout();
        recipientPage.ResumeLayout(false);
        recipientPage.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)nudLineHeightFactor).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudLandGapFactor).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudZipGapFactor).EndInit();
        tuningPage.ResumeLayout(false);
        tuningPage.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)nudSenderOffsetY).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudSenderOffsetX).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudRecipOffsetY).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudRecipOffsetX).EndInit();
        contextMenuStrip.ResumeLayout(false);
        statusStrip.ResumeLayout(false);
        statusStrip.PerformLayout();
        rightPanel.ResumeLayout(false);
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private TabControl tabControl;
    private TabPage printerPage;
    private TabPage formatPage;
    private Button btnSave;
    private Button btnCancel;
    private TabPage tuningPage;
    private GroupBox gbPrinter;
    private GroupBox gbOrientation;
    private Label lblSource;
    private Label lblDevice;
    private ComboBox cbSources;
    private RadioButton rbLandscape;
    private RadioButton rbPortrait;
    internal ComboBox cbPrinter;
    private System.Drawing.Printing.PrintDocument printDocument;
    private GroupBox gbFormat;
    private GroupBox gbText;
    internal ComboBox cbPapersize;
    private ComboBox cbFontsizeSender;
    private Label lblFontsize;
    private ComboBox cbFontSizeRecipient;
    private Label lblRecipient;
    private PictureBox picPortrait;
    private PictureBox picLandscape;
    internal ComboBox cbFont;
    private TabPage senderPage;
    private TextBox tbSender1;
    private NumericUpDown nudRecipOffsetX;
    private Label lblRecipOffsetX;
    private Label lblAddressOffset;
    private NumericUpDown nudSenderOffsetY;
    private Label label1;
    private Label label2;
    private Label label3;
    private NumericUpDown nudSenderOffsetX;
    private NumericUpDown nudRecipOffsetY;
    private Label lblRecipOffsetY;
    private CheckBox ckbBoldRecipient;
    private Label lblBold;
    private CheckBox ckbBoldSender;
    private Label lblHorizLine;
    private CheckBox ckbPrintSender;
    private TabControl tcSender;
    private TabPage tpSender1;
    private TabPage tpSender2;
    private TabPage tpSender3;
    private TextBox tbSender2;
    private TextBox tbSender3;
    private TabPage recipientPage;
    private CheckBox ckbAnredePrint;
    private CheckBox ckbLandPrint;
    private FlickerFreePrintPreviewControl printPreviewControl;
    private StatusStrip statusStrip;
    private ToolStripStatusLabel lblZoomStatus;
    private ToolStripStatusLabel toolStripStatusLabel;
    private ContextMenuStrip contextMenuStrip;
    private ToolStripMenuItem zoomInToolStripMenuItem;
    private ToolStripMenuItem zoomOutToolStripMenuItem;
    private ToolStripMenuItem zoomDefaultToolStripMenuItem;
    private TabPage tpSender4;
    private TabPage tpSender5;
    private TabPage tpSender6;
    private TextBox tbSender4;
    private TextBox tbSender5;
    private TextBox tbSender6;
    private Panel rightPanel;
    private System.Windows.Forms.Timer timerDebounce;
    private CheckBox ckbLandGROSS;
    private CheckBox ckbAnredeOberhalb;
    private Label lblZipGapFactor;
    private NumericUpDown nudZipGapFactor;
    private Label lblZipRows;
    private Label lblLandRows;
    private Label lblLandGapFactor;
    private NumericUpDown nudLandGapFactor;
    private Label lblLineFactor;
    private Label lblLineHeight;
    private NumericUpDown nudLineHeightFactor;
}