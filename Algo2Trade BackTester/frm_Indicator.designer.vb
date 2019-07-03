<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Indicator
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Indicator))
        Me.cmb_indicator = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pnl_ema = New System.Windows.Forms.Panel()
        Me.cmb_emafield = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.nmrc_emaPeriod = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.pnl_fractals = New System.Windows.Forms.Panel()
        Me.nmrc_fractalsPeriod = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.pnl_sma = New System.Windows.Forms.Panel()
        Me.cmb_smafields = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.nmrc_smaperiod = New System.Windows.Forms.NumericUpDown()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btn_view = New System.Windows.Forms.Button()
        Me.pnl_smi = New System.Windows.Forms.Panel()
        Me.nmrc_KdoubleSmoothing = New System.Windows.Forms.NumericUpDown()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.nmrc_K = New System.Windows.Forms.NumericUpDown()
        Me.nmrc_Ksmoothing = New System.Windows.Forms.NumericUpDown()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.nmrc_D = New System.Windows.Forms.NumericUpDown()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btn_export = New System.Windows.Forms.Button()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.pnl_atr = New System.Windows.Forms.Panel()
        Me.nmrc_atr = New System.Windows.Forms.NumericUpDown()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.pnl_cci = New System.Windows.Forms.Panel()
        Me.nmrc_cciPeriod = New System.Windows.Forms.NumericUpDown()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.pnl_rsi = New System.Windows.Forms.Panel()
        Me.nmrc_rsiPeriod = New System.Windows.Forms.NumericUpDown()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.pnl_macd = New System.Windows.Forms.Panel()
        Me.nmrc_macdSignal = New System.Windows.Forms.NumericUpDown()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.nmrc_macdSlow = New System.Windows.Forms.NumericUpDown()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.nmrc_macdFast = New System.Windows.Forms.NumericUpDown()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblProgressStatus = New System.Windows.Forms.Label()
        Me.uc_Indicator = New Algo2Trade_BackTester.uc_Indicator()
        Me.pnl_ema.SuspendLayout()
        CType(Me.nmrc_emaPeriod, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnl_fractals.SuspendLayout()
        CType(Me.nmrc_fractalsPeriod, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnl_sma.SuspendLayout()
        CType(Me.nmrc_smaperiod, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnl_smi.SuspendLayout()
        CType(Me.nmrc_KdoubleSmoothing, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nmrc_K, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nmrc_Ksmoothing, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nmrc_D, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnl_atr.SuspendLayout()
        CType(Me.nmrc_atr, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnl_cci.SuspendLayout()
        CType(Me.nmrc_cciPeriod, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnl_rsi.SuspendLayout()
        CType(Me.nmrc_rsiPeriod, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnl_macd.SuspendLayout()
        CType(Me.nmrc_macdSignal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nmrc_macdSlow, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nmrc_macdFast, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmb_indicator
        '
        Me.cmb_indicator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb_indicator.FormattingEnabled = True
        Me.cmb_indicator.Items.AddRange(New Object() {"VWAP", "Fractals", "Simple Moving Average", "Exponential Moving Average", "Stochastic Momentum Index", "Average True Range", "Commodity Channel Index", "Relative Strength Index", "MACD", "Heiken Ashi", "Supertrend", "ATR Trailing Stop", "Fractal1", "Fractal With SMA", "Fractal Retracement Entry Strategy Rule", "ATR Bands", "Naughty Boy VWAP Rule", "VWAP Double Confirmation Rule", "In The Trend Rule", "In The Trend Pearsing Rule", "Swing High Low", "Double U Rule", "MR With One Fractal And MA", "JOYMA4: Now I Should Go Rule", "Higher High Higher Low Pattern Rule", "Bollinger Bands", "Renko Chart", "Trend Intensity Index"})
        Me.cmb_indicator.Location = New System.Drawing.Point(136, 18)
        Me.cmb_indicator.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmb_indicator.Name = "cmb_indicator"
        Me.cmb_indicator.Size = New System.Drawing.Size(277, 24)
        Me.cmb_indicator.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(118, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Choose Indicator:"
        '
        'pnl_ema
        '
        Me.pnl_ema.Controls.Add(Me.cmb_emafield)
        Me.pnl_ema.Controls.Add(Me.Label3)
        Me.pnl_ema.Controls.Add(Me.nmrc_emaPeriod)
        Me.pnl_ema.Controls.Add(Me.Label2)
        Me.pnl_ema.Location = New System.Drawing.Point(9, 48)
        Me.pnl_ema.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnl_ema.Name = "pnl_ema"
        Me.pnl_ema.Size = New System.Drawing.Size(397, 44)
        Me.pnl_ema.TabIndex = 4
        '
        'cmb_emafield
        '
        Me.cmb_emafield.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb_emafield.FormattingEnabled = True
        Me.cmb_emafield.Items.AddRange(New Object() {"Open", "Low", "High", "Close", "Volume"})
        Me.cmb_emafield.Location = New System.Drawing.Point(221, 11)
        Me.cmb_emafield.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmb_emafield.Name = "cmb_emafield"
        Me.cmb_emafield.Size = New System.Drawing.Size(121, 24)
        Me.cmb_emafield.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(145, 14)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "EMA Field:"
        '
        'nmrc_emaPeriod
        '
        Me.nmrc_emaPeriod.Location = New System.Drawing.Point(87, 12)
        Me.nmrc_emaPeriod.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_emaPeriod.Name = "nmrc_emaPeriod"
        Me.nmrc_emaPeriod.Size = New System.Drawing.Size(52, 22)
        Me.nmrc_emaPeriod.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 17)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "EMA Period:"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DodgerBlue
        Me.Panel1.Location = New System.Drawing.Point(1, 1)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1337, 11)
        Me.Panel1.TabIndex = 5
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DodgerBlue
        Me.Panel2.Location = New System.Drawing.Point(1, 535)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1337, 11)
        Me.Panel2.TabIndex = 6
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DodgerBlue
        Me.Panel3.Location = New System.Drawing.Point(1, 1)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(11, 543)
        Me.Panel3.TabIndex = 7
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.DodgerBlue
        Me.Panel4.Location = New System.Drawing.Point(1329, 1)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(11, 543)
        Me.Panel4.TabIndex = 8
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataGridView1.Location = New System.Drawing.Point(17, 98)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(1307, 398)
        Me.DataGridView1.TabIndex = 9
        '
        'pnl_fractals
        '
        Me.pnl_fractals.Controls.Add(Me.nmrc_fractalsPeriod)
        Me.pnl_fractals.Controls.Add(Me.Label5)
        Me.pnl_fractals.Location = New System.Drawing.Point(9, 48)
        Me.pnl_fractals.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnl_fractals.Name = "pnl_fractals"
        Me.pnl_fractals.Size = New System.Drawing.Size(397, 44)
        Me.pnl_fractals.TabIndex = 10
        '
        'nmrc_fractalsPeriod
        '
        Me.nmrc_fractalsPeriod.Location = New System.Drawing.Point(119, 12)
        Me.nmrc_fractalsPeriod.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_fractalsPeriod.Name = "nmrc_fractalsPeriod"
        Me.nmrc_fractalsPeriod.Size = New System.Drawing.Size(52, 22)
        Me.nmrc_fractalsPeriod.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 14)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(107, 17)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Fractals Period:"
        '
        'pnl_sma
        '
        Me.pnl_sma.Controls.Add(Me.cmb_smafields)
        Me.pnl_sma.Controls.Add(Me.Label4)
        Me.pnl_sma.Controls.Add(Me.nmrc_smaperiod)
        Me.pnl_sma.Controls.Add(Me.Label6)
        Me.pnl_sma.Location = New System.Drawing.Point(9, 48)
        Me.pnl_sma.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnl_sma.Name = "pnl_sma"
        Me.pnl_sma.Size = New System.Drawing.Size(397, 44)
        Me.pnl_sma.TabIndex = 11
        '
        'cmb_smafields
        '
        Me.cmb_smafields.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb_smafields.FormattingEnabled = True
        Me.cmb_smafields.Items.AddRange(New Object() {"Open", "Low", "High", "Close", "Volume"})
        Me.cmb_smafields.Location = New System.Drawing.Point(221, 11)
        Me.cmb_smafields.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmb_smafields.Name = "cmb_smafields"
        Me.cmb_smafields.Size = New System.Drawing.Size(121, 24)
        Me.cmb_smafields.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(145, 14)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 17)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "SMA Field:"
        '
        'nmrc_smaperiod
        '
        Me.nmrc_smaperiod.Location = New System.Drawing.Point(87, 12)
        Me.nmrc_smaperiod.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_smaperiod.Name = "nmrc_smaperiod"
        Me.nmrc_smaperiod.Size = New System.Drawing.Size(52, 22)
        Me.nmrc_smaperiod.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 14)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(86, 17)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "SMA Period:"
        '
        'btn_view
        '
        Me.btn_view.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_view.Location = New System.Drawing.Point(1208, 57)
        Me.btn_view.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn_view.Name = "btn_view"
        Me.btn_view.Size = New System.Drawing.Size(115, 36)
        Me.btn_view.TabIndex = 15
        Me.btn_view.Text = "View"
        Me.btn_view.UseVisualStyleBackColor = True
        '
        'pnl_smi
        '
        Me.pnl_smi.Controls.Add(Me.nmrc_KdoubleSmoothing)
        Me.pnl_smi.Controls.Add(Me.Label11)
        Me.pnl_smi.Controls.Add(Me.nmrc_K)
        Me.pnl_smi.Controls.Add(Me.nmrc_Ksmoothing)
        Me.pnl_smi.Controls.Add(Me.Label12)
        Me.pnl_smi.Controls.Add(Me.nmrc_D)
        Me.pnl_smi.Controls.Add(Me.Label13)
        Me.pnl_smi.Controls.Add(Me.Label14)
        Me.pnl_smi.Location = New System.Drawing.Point(9, 48)
        Me.pnl_smi.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnl_smi.Name = "pnl_smi"
        Me.pnl_smi.Size = New System.Drawing.Size(397, 44)
        Me.pnl_smi.TabIndex = 16
        '
        'nmrc_KdoubleSmoothing
        '
        Me.nmrc_KdoubleSmoothing.Location = New System.Drawing.Point(277, 14)
        Me.nmrc_KdoubleSmoothing.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_KdoubleSmoothing.Name = "nmrc_KdoubleSmoothing"
        Me.nmrc_KdoubleSmoothing.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_KdoubleSmoothing.TabIndex = 3
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(197, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(83, 17)
        Me.Label11.TabIndex = 8
        Me.Label11.Text = "K Double S:"
        '
        'nmrc_K
        '
        Me.nmrc_K.Location = New System.Drawing.Point(21, 12)
        Me.nmrc_K.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_K.Name = "nmrc_K"
        Me.nmrc_K.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_K.TabIndex = 1
        '
        'nmrc_Ksmoothing
        '
        Me.nmrc_Ksmoothing.Location = New System.Drawing.Point(151, 14)
        Me.nmrc_Ksmoothing.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_Ksmoothing.Name = "nmrc_Ksmoothing"
        Me.nmrc_Ksmoothing.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_Ksmoothing.TabIndex = 7
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(63, 14)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(92, 17)
        Me.Label12.TabIndex = 6
        Me.Label12.Text = "K Smoothing:"
        '
        'nmrc_D
        '
        Me.nmrc_D.Location = New System.Drawing.Point(344, 14)
        Me.nmrc_D.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_D.Name = "nmrc_D"
        Me.nmrc_D.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_D.TabIndex = 5
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(325, 16)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(22, 17)
        Me.Label13.TabIndex = 4
        Me.Label13.Text = "D:"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(4, 14)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(21, 17)
        Me.Label14.TabIndex = 0
        Me.Label14.Text = "K:"
        '
        'btn_export
        '
        Me.btn_export.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_export.Location = New System.Drawing.Point(1208, 502)
        Me.btn_export.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btn_export.Name = "btn_export"
        Me.btn_export.Size = New System.Drawing.Size(115, 30)
        Me.btn_export.TabIndex = 17
        Me.btn_export.Text = "Export CSV"
        Me.btn_export.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(135, 262)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(1069, 23)
        Me.ProgressBar1.TabIndex = 18
        '
        'pnl_atr
        '
        Me.pnl_atr.Controls.Add(Me.nmrc_atr)
        Me.pnl_atr.Controls.Add(Me.Label10)
        Me.pnl_atr.Location = New System.Drawing.Point(9, 48)
        Me.pnl_atr.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnl_atr.Name = "pnl_atr"
        Me.pnl_atr.Size = New System.Drawing.Size(397, 44)
        Me.pnl_atr.TabIndex = 19
        '
        'nmrc_atr
        '
        Me.nmrc_atr.Location = New System.Drawing.Point(91, 12)
        Me.nmrc_atr.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_atr.Name = "nmrc_atr"
        Me.nmrc_atr.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_atr.TabIndex = 1
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(4, 14)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(85, 17)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "ATR Period:"
        '
        'pnl_cci
        '
        Me.pnl_cci.Controls.Add(Me.nmrc_cciPeriod)
        Me.pnl_cci.Controls.Add(Me.Label7)
        Me.pnl_cci.Location = New System.Drawing.Point(9, 48)
        Me.pnl_cci.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnl_cci.Name = "pnl_cci"
        Me.pnl_cci.Size = New System.Drawing.Size(397, 44)
        Me.pnl_cci.TabIndex = 20
        '
        'nmrc_cciPeriod
        '
        Me.nmrc_cciPeriod.Location = New System.Drawing.Point(91, 12)
        Me.nmrc_cciPeriod.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_cciPeriod.Name = "nmrc_cciPeriod"
        Me.nmrc_cciPeriod.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_cciPeriod.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(4, 14)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(78, 17)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "CCI Period:"
        '
        'pnl_rsi
        '
        Me.pnl_rsi.Controls.Add(Me.nmrc_rsiPeriod)
        Me.pnl_rsi.Controls.Add(Me.Label8)
        Me.pnl_rsi.Location = New System.Drawing.Point(9, 48)
        Me.pnl_rsi.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnl_rsi.Name = "pnl_rsi"
        Me.pnl_rsi.Size = New System.Drawing.Size(397, 44)
        Me.pnl_rsi.TabIndex = 21
        '
        'nmrc_rsiPeriod
        '
        Me.nmrc_rsiPeriod.Location = New System.Drawing.Point(91, 12)
        Me.nmrc_rsiPeriod.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_rsiPeriod.Name = "nmrc_rsiPeriod"
        Me.nmrc_rsiPeriod.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_rsiPeriod.TabIndex = 1
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(4, 14)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(79, 17)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "RSI Period:"
        '
        'pnl_macd
        '
        Me.pnl_macd.Controls.Add(Me.nmrc_macdSignal)
        Me.pnl_macd.Controls.Add(Me.Label16)
        Me.pnl_macd.Controls.Add(Me.nmrc_macdSlow)
        Me.pnl_macd.Controls.Add(Me.Label15)
        Me.pnl_macd.Controls.Add(Me.nmrc_macdFast)
        Me.pnl_macd.Controls.Add(Me.Label9)
        Me.pnl_macd.Location = New System.Drawing.Point(9, 48)
        Me.pnl_macd.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pnl_macd.Name = "pnl_macd"
        Me.pnl_macd.Size = New System.Drawing.Size(397, 44)
        Me.pnl_macd.TabIndex = 23
        '
        'nmrc_macdSignal
        '
        Me.nmrc_macdSignal.Location = New System.Drawing.Point(340, 14)
        Me.nmrc_macdSignal.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_macdSignal.Name = "nmrc_macdSignal"
        Me.nmrc_macdSignal.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_macdSignal.TabIndex = 26
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(292, 15)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(51, 17)
        Me.Label16.TabIndex = 25
        Me.Label16.Text = "Signal:"
        '
        'nmrc_macdSlow
        '
        Me.nmrc_macdSlow.Location = New System.Drawing.Point(217, 14)
        Me.nmrc_macdSlow.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_macdSlow.Name = "nmrc_macdSlow"
        Me.nmrc_macdSlow.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_macdSlow.TabIndex = 3
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(152, 14)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(65, 17)
        Me.Label15.TabIndex = 2
        Me.Label15.Text = "Slow MA:"
        '
        'nmrc_macdFast
        '
        Me.nmrc_macdFast.Location = New System.Drawing.Point(69, 12)
        Me.nmrc_macdFast.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.nmrc_macdFast.Name = "nmrc_macdFast"
        Me.nmrc_macdFast.Size = New System.Drawing.Size(43, 22)
        Me.nmrc_macdFast.TabIndex = 1
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(4, 14)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(63, 17)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Fast MA:"
        '
        'lblProgressStatus
        '
        Me.lblProgressStatus.AutoSize = True
        Me.lblProgressStatus.Location = New System.Drawing.Point(20, 511)
        Me.lblProgressStatus.Name = "lblProgressStatus"
        Me.lblProgressStatus.Size = New System.Drawing.Size(133, 17)
        Me.lblProgressStatus.TabIndex = 24
        Me.lblProgressStatus.Text = "Progress Status ....."
        '
        'uc_Indicator
        '
        Me.uc_Indicator.Location = New System.Drawing.Point(419, 6)
        Me.uc_Indicator.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.uc_Indicator.Name = "uc_Indicator"
        Me.uc_Indicator.Size = New System.Drawing.Size(919, 95)
        Me.uc_Indicator.TabIndex = 2
        '
        'frm_Indicator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1339, 546)
        Me.Controls.Add(Me.lblProgressStatus)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.pnl_macd)
        Me.Controls.Add(Me.pnl_rsi)
        Me.Controls.Add(Me.pnl_cci)
        Me.Controls.Add(Me.pnl_atr)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btn_export)
        Me.Controls.Add(Me.pnl_smi)
        Me.Controls.Add(Me.btn_view)
        Me.Controls.Add(Me.pnl_sma)
        Me.Controls.Add(Me.pnl_fractals)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.pnl_ema)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmb_indicator)
        Me.Controls.Add(Me.uc_Indicator)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.Name = "frm_Indicator"
        Me.Text = "Indicator"
        Me.pnl_ema.ResumeLayout(False)
        Me.pnl_ema.PerformLayout()
        CType(Me.nmrc_emaPeriod, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnl_fractals.ResumeLayout(False)
        Me.pnl_fractals.PerformLayout()
        CType(Me.nmrc_fractalsPeriod, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnl_sma.ResumeLayout(False)
        Me.pnl_sma.PerformLayout()
        CType(Me.nmrc_smaperiod, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnl_smi.ResumeLayout(False)
        Me.pnl_smi.PerformLayout()
        CType(Me.nmrc_KdoubleSmoothing, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nmrc_K, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nmrc_Ksmoothing, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nmrc_D, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnl_atr.ResumeLayout(False)
        Me.pnl_atr.PerformLayout()
        CType(Me.nmrc_atr, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnl_cci.ResumeLayout(False)
        Me.pnl_cci.PerformLayout()
        CType(Me.nmrc_cciPeriod, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnl_rsi.ResumeLayout(False)
        Me.pnl_rsi.PerformLayout()
        CType(Me.nmrc_rsiPeriod, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnl_macd.ResumeLayout(False)
        Me.pnl_macd.PerformLayout()
        CType(Me.nmrc_macdSignal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nmrc_macdSlow, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nmrc_macdFast, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmb_indicator As ComboBox
    Friend WithEvents uc_Indicator As uc_Indicator
    Friend WithEvents Label1 As Label
    Friend WithEvents pnl_ema As Panel
    Friend WithEvents nmrc_emaPeriod As NumericUpDown
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents cmb_emafield As ComboBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Panel4 As Panel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents pnl_fractals As Panel
    Friend WithEvents nmrc_fractalsPeriod As NumericUpDown
    Friend WithEvents Label5 As Label
    Friend WithEvents pnl_sma As Panel
    Friend WithEvents cmb_smafields As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents nmrc_smaperiod As NumericUpDown
    Friend WithEvents Label6 As Label
    Friend WithEvents btn_view As Button
    Friend WithEvents pnl_smi As Panel
    Friend WithEvents nmrc_KdoubleSmoothing As NumericUpDown
    Friend WithEvents Label11 As Label
    Friend WithEvents nmrc_K As NumericUpDown
    Friend WithEvents nmrc_Ksmoothing As NumericUpDown
    Friend WithEvents Label12 As Label
    Friend WithEvents nmrc_D As NumericUpDown
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents btn_export As Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents pnl_atr As Panel
    Friend WithEvents nmrc_atr As NumericUpDown
    Friend WithEvents Label10 As Label
    Friend WithEvents pnl_cci As Panel
    Friend WithEvents nmrc_cciPeriod As NumericUpDown
    Friend WithEvents Label7 As Label
    Friend WithEvents pnl_rsi As Panel
    Friend WithEvents nmrc_rsiPeriod As NumericUpDown
    Friend WithEvents Label8 As Label
    Friend WithEvents pnl_macd As Panel
    Friend WithEvents nmrc_macdSignal As NumericUpDown
    Friend WithEvents Label16 As Label
    Friend WithEvents nmrc_macdSlow As NumericUpDown
    Friend WithEvents Label15 As Label
    Friend WithEvents nmrc_macdFast As NumericUpDown
    Friend WithEvents Label9 As Label
    Friend WithEvents lblProgressStatus As Label
End Class
