﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_BackTest
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_BackTest))
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.cmb_strategy = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_start = New System.Windows.Forms.Button()
        Me.lblProgressStatus = New System.Windows.Forms.Label()
        Me.btn_cancel = New System.Windows.Forms.Button()
        Me.uc_BackTest = New Algo2Trade_BackTester.uc_BackTest()
        Me.grpbxDataSource = New System.Windows.Forms.GroupBox()
        Me.rdbLive = New System.Windows.Forms.RadioButton()
        Me.rdbDatabase = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtSlippageMultiplier = New System.Windows.Forms.TextBox()
        Me.lblSlippageMultiplier = New System.Windows.Forms.Label()
        Me.chkbIncludeSlippage = New System.Windows.Forms.CheckBox()
        Me.grpbxDataSource.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Panel4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Panel4.Location = New System.Drawing.Point(3, 346)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(916, 10)
        Me.Panel4.TabIndex = 11
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Panel5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Panel5.Location = New System.Drawing.Point(3, 1)
        Me.Panel5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(916, 10)
        Me.Panel5.TabIndex = 12
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Panel6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Panel6.Location = New System.Drawing.Point(3, 1)
        Me.Panel6.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(11, 354)
        Me.Panel6.TabIndex = 13
        '
        'Panel7
        '
        Me.Panel7.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Panel7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Panel7.Location = New System.Drawing.Point(911, 1)
        Me.Panel7.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(11, 354)
        Me.Panel7.TabIndex = 14
        '
        'cmb_strategy
        '
        Me.cmb_strategy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmb_strategy.FormattingEnabled = True
        Me.cmb_strategy.Items.AddRange(New Object() {"Fractal With MA Retracement Candle Strategy", "Naughty Boy VWAP Strategy", "In The Trend Strategy", "Gap Fill Strategy", "Gap Fill Without Filling PC Strategy", "Gap Fill Without Filling Neglect GnG Strategy", "Gap Fill Without Filling PHLC Strategy", "Generic Strategy", "Maruti Strategy", "Pair Hazing Strategy"})
        Me.cmb_strategy.Location = New System.Drawing.Point(148, 18)
        Me.cmb_strategy.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmb_strategy.Name = "cmb_strategy"
        Me.cmb_strategy.Size = New System.Drawing.Size(333, 26)
        Me.cmb_strategy.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(123, 18)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Choose Strategy:"
        '
        'btn_start
        '
        Me.btn_start.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_start.Location = New System.Drawing.Point(408, 153)
        Me.btn_start.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn_start.Name = "btn_start"
        Me.btn_start.Size = New System.Drawing.Size(105, 50)
        Me.btn_start.TabIndex = 1
        Me.btn_start.Text = "Start"
        Me.btn_start.UseVisualStyleBackColor = True
        '
        'lblProgressStatus
        '
        Me.lblProgressStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgressStatus.Location = New System.Drawing.Point(19, 300)
        Me.lblProgressStatus.Name = "lblProgressStatus"
        Me.lblProgressStatus.Size = New System.Drawing.Size(888, 39)
        Me.lblProgressStatus.TabIndex = 19
        Me.lblProgressStatus.Text = "Progress Status ................."
        '
        'btn_cancel
        '
        Me.btn_cancel.Location = New System.Drawing.Point(780, 318)
        Me.btn_cancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn_cancel.Name = "btn_cancel"
        Me.btn_cancel.Size = New System.Drawing.Size(127, 26)
        Me.btn_cancel.TabIndex = 20
        Me.btn_cancel.Text = "Cancel Process"
        Me.btn_cancel.UseVisualStyleBackColor = True
        Me.btn_cancel.Visible = False
        '
        'uc_BackTest
        '
        Me.uc_BackTest.Location = New System.Drawing.Point(503, 15)
        Me.uc_BackTest.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.uc_BackTest.Name = "uc_BackTest"
        Me.uc_BackTest.Size = New System.Drawing.Size(403, 34)
        Me.uc_BackTest.TabIndex = 0
        '
        'grpbxDataSource
        '
        Me.grpbxDataSource.Controls.Add(Me.rdbLive)
        Me.grpbxDataSource.Controls.Add(Me.rdbDatabase)
        Me.grpbxDataSource.Location = New System.Drawing.Point(19, 59)
        Me.grpbxDataSource.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxDataSource.Name = "grpbxDataSource"
        Me.grpbxDataSource.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.grpbxDataSource.Size = New System.Drawing.Size(187, 57)
        Me.grpbxDataSource.TabIndex = 21
        Me.grpbxDataSource.TabStop = False
        Me.grpbxDataSource.Text = "Data Source"
        '
        'rdbLive
        '
        Me.rdbLive.AutoSize = True
        Me.rdbLive.Location = New System.Drawing.Point(115, 25)
        Me.rdbLive.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbLive.Name = "rdbLive"
        Me.rdbLive.Size = New System.Drawing.Size(55, 21)
        Me.rdbLive.TabIndex = 1
        Me.rdbLive.TabStop = True
        Me.rdbLive.Text = "Live"
        Me.rdbLive.UseVisualStyleBackColor = True
        '
        'rdbDatabase
        '
        Me.rdbDatabase.AutoSize = True
        Me.rdbDatabase.Location = New System.Drawing.Point(7, 23)
        Me.rdbDatabase.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rdbDatabase.Name = "rdbDatabase"
        Me.rdbDatabase.Size = New System.Drawing.Size(90, 21)
        Me.rdbDatabase.TabIndex = 0
        Me.rdbDatabase.TabStop = True
        Me.rdbDatabase.Text = "Database"
        Me.rdbDatabase.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtSlippageMultiplier)
        Me.GroupBox1.Controls.Add(Me.lblSlippageMultiplier)
        Me.GroupBox1.Controls.Add(Me.chkbIncludeSlippage)
        Me.GroupBox1.Location = New System.Drawing.Point(229, 59)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(384, 57)
        Me.GroupBox1.TabIndex = 22
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Slippage"
        '
        'txtSlippageMultiplier
        '
        Me.txtSlippageMultiplier.Location = New System.Drawing.Point(283, 21)
        Me.txtSlippageMultiplier.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtSlippageMultiplier.Name = "txtSlippageMultiplier"
        Me.txtSlippageMultiplier.Size = New System.Drawing.Size(88, 22)
        Me.txtSlippageMultiplier.TabIndex = 3
        '
        'lblSlippageMultiplier
        '
        Me.lblSlippageMultiplier.AutoSize = True
        Me.lblSlippageMultiplier.Location = New System.Drawing.Point(154, 26)
        Me.lblSlippageMultiplier.Name = "lblSlippageMultiplier"
        Me.lblSlippageMultiplier.Size = New System.Drawing.Size(123, 17)
        Me.lblSlippageMultiplier.TabIndex = 2
        Me.lblSlippageMultiplier.Text = "Slippage Multiplier"
        '
        'chkbIncludeSlippage
        '
        Me.chkbIncludeSlippage.AutoSize = True
        Me.chkbIncludeSlippage.Location = New System.Drawing.Point(7, 25)
        Me.chkbIncludeSlippage.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkbIncludeSlippage.Name = "chkbIncludeSlippage"
        Me.chkbIncludeSlippage.Size = New System.Drawing.Size(134, 21)
        Me.chkbIncludeSlippage.TabIndex = 0
        Me.chkbIncludeSlippage.Text = "Include Slippage"
        Me.chkbIncludeSlippage.UseVisualStyleBackColor = True
        '
        'frm_BackTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(921, 356)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpbxDataSource)
        Me.Controls.Add(Me.btn_cancel)
        Me.Controls.Add(Me.cmb_strategy)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.uc_BackTest)
        Me.Controls.Add(Me.btn_start)
        Me.Controls.Add(Me.lblProgressStatus)
        Me.Controls.Add(Me.Panel7)
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.Panel4)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.Name = "frm_BackTest"
        Me.Text = "BackTest"
        Me.grpbxDataSource.ResumeLayout(False)
        Me.grpbxDataSource.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Panel5 As Panel
    Friend WithEvents Panel6 As Panel
    Friend WithEvents Panel7 As Panel
    Friend WithEvents cmb_strategy As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents uc_BackTest As uc_BackTest
    Friend WithEvents btn_start As Button
    Friend WithEvents lblProgressStatus As Label
    Friend WithEvents btn_cancel As Button
    Friend WithEvents grpbxDataSource As GroupBox
    Friend WithEvents rdbLive As RadioButton
    Friend WithEvents rdbDatabase As RadioButton
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lblSlippageMultiplier As Label
    Friend WithEvents chkbIncludeSlippage As CheckBox
    Friend WithEvents txtSlippageMultiplier As TextBox
End Class
