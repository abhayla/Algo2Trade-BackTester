<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class uc_Indicator
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtpckr_from = New System.Windows.Forms.DateTimePicker()
        Me.dtpckr_to = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmb_inscategory = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txt_instrument = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.chkb_volume = New System.Windows.Forms.CheckBox()
        Me.chkb_low = New System.Windows.Forms.CheckBox()
        Me.chkb_high = New System.Windows.Forms.CheckBox()
        Me.chkb_close = New System.Windows.Forms.CheckBox()
        Me.chkb_open = New System.Windows.Forms.CheckBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 17)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "From Date:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(199, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "To Date:"
        '
        'dtpckr_from
        '
        Me.dtpckr_from.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckr_from.Location = New System.Drawing.Point(87, 13)
        Me.dtpckr_from.Name = "dtpckr_from"
        Me.dtpckr_from.Size = New System.Drawing.Size(108, 22)
        Me.dtpckr_from.TabIndex = 7
        '
        'dtpckr_to
        '
        Me.dtpckr_to.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckr_to.Location = New System.Drawing.Point(266, 14)
        Me.dtpckr_to.Name = "dtpckr_to"
        Me.dtpckr_to.Size = New System.Drawing.Size(108, 22)
        Me.dtpckr_to.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(378, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(139, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Instrument Category:"
        '
        'cmb_inscategory
        '
        Me.cmb_inscategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb_inscategory.FormattingEnabled = True
        Me.cmb_inscategory.Items.AddRange(New Object() {"Cash", "Currency", "Commodity", "Future"})
        Me.cmb_inscategory.Location = New System.Drawing.Point(521, 14)
        Me.cmb_inscategory.Name = "cmb_inscategory"
        Me.cmb_inscategory.Size = New System.Drawing.Size(121, 24)
        Me.cmb_inscategory.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(646, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(119, 17)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Instrument Name:"
        '
        'txt_instrument
        '
        Me.txt_instrument.Location = New System.Drawing.Point(769, 17)
        Me.txt_instrument.Name = "txt_instrument"
        Me.txt_instrument.Size = New System.Drawing.Size(138, 22)
        Me.txt_instrument.TabIndex = 12
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.chkb_volume)
        Me.Panel1.Controls.Add(Me.chkb_low)
        Me.Panel1.Controls.Add(Me.chkb_high)
        Me.Panel1.Controls.Add(Me.chkb_close)
        Me.Panel1.Controls.Add(Me.chkb_open)
        Me.Panel1.Location = New System.Drawing.Point(7, 51)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(372, 36)
        Me.Panel1.TabIndex = 13
        '
        'chkb_volume
        '
        Me.chkb_volume.AutoSize = True
        Me.chkb_volume.Location = New System.Drawing.Point(288, 8)
        Me.chkb_volume.Name = "chkb_volume"
        Me.chkb_volume.Size = New System.Drawing.Size(77, 21)
        Me.chkb_volume.TabIndex = 9
        Me.chkb_volume.Text = "Volume"
        Me.chkb_volume.UseVisualStyleBackColor = True
        '
        'chkb_low
        '
        Me.chkb_low.AutoSize = True
        Me.chkb_low.Location = New System.Drawing.Point(223, 8)
        Me.chkb_low.Name = "chkb_low"
        Me.chkb_low.Size = New System.Drawing.Size(55, 21)
        Me.chkb_low.TabIndex = 8
        Me.chkb_low.Text = "Low"
        Me.chkb_low.UseVisualStyleBackColor = True
        '
        'chkb_high
        '
        Me.chkb_high.AutoSize = True
        Me.chkb_high.Location = New System.Drawing.Point(154, 8)
        Me.chkb_high.Name = "chkb_high"
        Me.chkb_high.Size = New System.Drawing.Size(59, 21)
        Me.chkb_high.TabIndex = 7
        Me.chkb_high.Text = "High"
        Me.chkb_high.UseVisualStyleBackColor = True
        '
        'chkb_close
        '
        Me.chkb_close.AutoSize = True
        Me.chkb_close.Location = New System.Drawing.Point(79, 8)
        Me.chkb_close.Name = "chkb_close"
        Me.chkb_close.Size = New System.Drawing.Size(65, 21)
        Me.chkb_close.TabIndex = 6
        Me.chkb_close.Text = "Close"
        Me.chkb_close.UseVisualStyleBackColor = True
        '
        'chkb_open
        '
        Me.chkb_open.AutoSize = True
        Me.chkb_open.Location = New System.Drawing.Point(4, 8)
        Me.chkb_open.Name = "chkb_open"
        Me.chkb_open.Size = New System.Drawing.Size(65, 21)
        Me.chkb_open.TabIndex = 5
        Me.chkb_open.Text = "Open"
        Me.chkb_open.UseVisualStyleBackColor = True
        '
        'uc_Indicator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.txt_instrument)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmb_inscategory)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dtpckr_to)
        Me.Controls.Add(Me.dtpckr_from)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "uc_Indicator"
        Me.Size = New System.Drawing.Size(919, 95)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents dtpckr_from As DateTimePicker
    Friend WithEvents dtpckr_to As DateTimePicker
    Friend WithEvents Label3 As Label
    Friend WithEvents cmb_inscategory As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txt_instrument As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents chkb_volume As CheckBox
    Friend WithEvents chkb_low As CheckBox
    Friend WithEvents chkb_high As CheckBox
    Friend WithEvents chkb_close As CheckBox
    Friend WithEvents chkb_open As CheckBox
End Class
