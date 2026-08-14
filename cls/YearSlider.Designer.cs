namespace Adressen.cls;

partial class YearSlider
{
    /// <summary> 
    /// Erforderliche Designervariable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary> 
    /// Verwendete Ressourcen bereinigen.
    /// </summary>
    /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Vom Komponenten-Designer generierter Code

    /// <summary> 
    /// Erforderliche Methode für die Designerunterstützung. 
    /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
    /// </summary>
    private void InitializeComponent()
    {
        btnPrev = new Button();
        txtYear = new TextBox();
        btnNext = new Button();
        SuspendLayout();
        // 
        // btnPrev
        // 
        btnPrev.FlatStyle = FlatStyle.Flat;
        btnPrev.Image = Properties.Resources.GlyphLeft16x;
        btnPrev.Location = new Point(-1, -1);
        btnPrev.Name = "btnPrev";
        btnPrev.Size = new Size(24, 25);
        btnPrev.TabIndex = 0;
        btnPrev.UseVisualStyleBackColor = true;
        // 
        // txtYear
        // 
        txtYear.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        txtYear.BackColor = SystemColors.Window;
        txtYear.BorderStyle = BorderStyle.None;
        txtYear.Location = new Point(23, 2);
        txtYear.Name = "txtYear";
        txtYear.Size = new Size(52, 18);
        txtYear.TabIndex = 1;
        txtYear.TextAlign = HorizontalAlignment.Center;
        txtYear.TextChanged += TxtYear_TextChanged;
        // 
        // btnNext
        // 
        btnNext.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnNext.FlatStyle = FlatStyle.Flat;
        btnNext.Image = Properties.Resources.GlyphRight16x;
        btnNext.Location = new Point(75, -1);
        btnNext.Name = "btnNext";
        btnNext.Size = new Size(24, 25);
        btnNext.TabIndex = 2;
        btnNext.UseVisualStyleBackColor = true;
        // 
        // YearSlider
        // 
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        BackColor = SystemColors.Window;
        BorderStyle = BorderStyle.FixedSingle;
        Controls.Add(btnNext);
        Controls.Add(txtYear);
        Controls.Add(btnPrev);
        Font = new Font("Segoe UI", 10F);
        Name = "YearSlider";
        Size = new Size(99, 24);
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private Button btnPrev;
    private TextBox txtYear;
    private Button btnNext;
}
