<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class uc_BackTest
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
        Me.dtpckr_to = New System.Windows.Forms.DateTimePicker()
        Me.dtpckr_from = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'dtpckr_to
        '
        Me.dtpckr_to.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckr_to.Location = New System.Drawing.Point(281, 6)
        Me.dtpckr_to.Name = "dtpckr_to"
        Me.dtpckr_to.Size = New System.Drawing.Size(108, 22)
        Me.dtpckr_to.TabIndex = 12
        '
        'dtpckr_from
        '
        Me.dtpckr_from.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpckr_from.Location = New System.Drawing.Point(88, 6)
        Me.dtpckr_from.Name = "dtpckr_from"
        Me.dtpckr_from.Size = New System.Drawing.Size(108, 22)
        Me.dtpckr_from.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(207, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 17)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "To Date:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(-1, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "From Date:"
        '
        'uc_BackTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.dtpckr_to)
        Me.Controls.Add(Me.dtpckr_from)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "uc_BackTest"
        Me.Size = New System.Drawing.Size(400, 35)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dtpckr_to As DateTimePicker
    Friend WithEvents dtpckr_from As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
End Class
