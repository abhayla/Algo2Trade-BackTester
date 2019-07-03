<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_signalCheck
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_signalCheck))
        Me.dgrv_Signal = New System.Windows.Forms.DataGridView()
        Me.lbl_progress = New System.Windows.Forms.Label()
        Me.btn_view = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmb_rule = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.nmrc_signalTimeFrame = New System.Windows.Forms.NumericUpDown()
        Me.uc_date = New Algo2Trade_BackTester.uc_BackTest()
        Me.btn_export = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        CType(Me.dgrv_Signal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nmrc_signalTimeFrame, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgrv_Signal
        '
        Me.dgrv_Signal.AllowUserToAddRows = False
        Me.dgrv_Signal.AllowUserToDeleteRows = False
        Me.dgrv_Signal.AllowUserToResizeRows = False
        Me.dgrv_Signal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgrv_Signal.Cursor = System.Windows.Forms.Cursors.Default
        Me.dgrv_Signal.Location = New System.Drawing.Point(2, 88)
        Me.dgrv_Signal.Margin = New System.Windows.Forms.Padding(2)
        Me.dgrv_Signal.Name = "dgrv_Signal"
        Me.dgrv_Signal.ReadOnly = True
        Me.dgrv_Signal.RowHeadersVisible = False
        Me.dgrv_Signal.RowTemplate.Height = 24
        Me.dgrv_Signal.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgrv_Signal.Size = New System.Drawing.Size(806, 339)
        Me.dgrv_Signal.TabIndex = 10
        '
        'lbl_progress
        '
        Me.lbl_progress.AutoSize = True
        Me.lbl_progress.Location = New System.Drawing.Point(2, 432)
        Me.lbl_progress.Name = "lbl_progress"
        Me.lbl_progress.Size = New System.Drawing.Size(96, 13)
        Me.lbl_progress.TabIndex = 11
        Me.lbl_progress.Text = "Progess Status ....."
        '
        'btn_view
        '
        Me.btn_view.Location = New System.Drawing.Point(715, 13)
        Me.btn_view.Name = "btn_view"
        Me.btn_view.Size = New System.Drawing.Size(75, 23)
        Me.btn_view.TabIndex = 13
        Me.btn_view.Text = "View"
        Me.btn_view.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Choose Rule:"
        '
        'cmb_rule
        '
        Me.cmb_rule.FormattingEnabled = True
        Me.cmb_rule.Items.AddRange(New Object() {"VWALL Rule", "Constricted Range", "Gap & Go", "Gap & Go After Gap Fill"})
        Me.cmb_rule.Location = New System.Drawing.Point(75, 14)
        Me.cmb_rule.Name = "cmb_rule"
        Me.cmb_rule.Size = New System.Drawing.Size(299, 21)
        Me.cmb_rule.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 13)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Signal TimeFrame:"
        '
        'nmrc_signalTimeFrame
        '
        Me.nmrc_signalTimeFrame.Location = New System.Drawing.Point(98, 42)
        Me.nmrc_signalTimeFrame.Name = "nmrc_signalTimeFrame"
        Me.nmrc_signalTimeFrame.Size = New System.Drawing.Size(48, 20)
        Me.nmrc_signalTimeFrame.TabIndex = 17
        '
        'uc_date
        '
        Me.uc_date.Location = New System.Drawing.Point(390, 11)
        Me.uc_date.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.uc_date.Name = "uc_date"
        Me.uc_date.Size = New System.Drawing.Size(300, 28)
        Me.uc_date.TabIndex = 12
        '
        'btn_export
        '
        Me.btn_export.Location = New System.Drawing.Point(733, 427)
        Me.btn_export.Name = "btn_export"
        Me.btn_export.Size = New System.Drawing.Size(75, 23)
        Me.btn_export.TabIndex = 18
        Me.btn_export.Text = "Export"
        Me.btn_export.UseVisualStyleBackColor = True
        '
        'SaveFileDialog1
        '
        '
        'frm_signalCheck
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(811, 450)
        Me.Controls.Add(Me.btn_export)
        Me.Controls.Add(Me.nmrc_signalTimeFrame)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmb_rule)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_view)
        Me.Controls.Add(Me.uc_date)
        Me.Controls.Add(Me.lbl_progress)
        Me.Controls.Add(Me.dgrv_Signal)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frm_signalCheck"
        Me.Text = "Signal Check"
        CType(Me.dgrv_Signal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nmrc_signalTimeFrame, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgrv_Signal As DataGridView
    Friend WithEvents lbl_progress As Label
    Friend WithEvents uc_date As uc_BackTest
    Friend WithEvents btn_view As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents cmb_rule As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents nmrc_signalTimeFrame As NumericUpDown
    Friend WithEvents btn_export As Button
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
End Class
