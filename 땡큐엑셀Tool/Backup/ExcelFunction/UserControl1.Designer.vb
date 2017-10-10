<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserControl1
    Inherits System.Windows.Forms.UserControl

    'UserControl은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TabControl2 = New System.Windows.Forms.TabControl()
        Me.TabPage7 = New System.Windows.Forms.TabPage()
        Me.tbReplace1 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cbReplace1 = New System.Windows.Forms.CheckBox()
        Me.btnMergeRun2 = New System.Windows.Forms.Button()
        Me.tbTemplate = New System.Windows.Forms.TextBox()
        Me.cbReplace2 = New System.Windows.Forms.CheckBox()
        Me.cbReplace3 = New System.Windows.Forms.CheckBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.tbResultMergeRow = New System.Windows.Forms.TextBox()
        Me.tbReplace2 = New System.Windows.Forms.TextBox()
        Me.tbReplace3 = New System.Windows.Forms.TextBox()
        Me.TabPage8 = New System.Windows.Forms.TabPage()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.btnTrimRun = New System.Windows.Forms.Button()
        Me.tbResultTrimRow = New System.Windows.Forms.TextBox()
        Me.tbTrimSrcRow = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TabPage9 = New System.Windows.Forms.TabPage()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.btnNumberOnlyRun = New System.Windows.Forms.Button()
        Me.tbNumberOnlyOutRow = New System.Windows.Forms.TextBox()
        Me.tbNumberOnlySrcRow = New System.Windows.Forms.TextBox()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.TabPage10 = New System.Windows.Forms.TabPage()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.btnByteLengthRun = New System.Windows.Forms.Button()
        Me.tbByteLengthOutRow = New System.Windows.Forms.TextBox()
        Me.tbByteLengthSrcRow = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.btnCelTypeRun = New System.Windows.Forms.Button()
        Me.tbCelTypeSrcRow = New System.Windows.Forms.TextBox()
        Me.TextBox8 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.SplitContainer1.Panel1.SuspendLayout
        Me.SplitContainer1.Panel2.SuspendLayout
        Me.SplitContainer1.SuspendLayout
        Me.TabControl1.SuspendLayout
        Me.TabPage2.SuspendLayout
        Me.TabControl2.SuspendLayout
        Me.TabPage7.SuspendLayout
        Me.TabPage8.SuspendLayout
        Me.TabPage9.SuspendLayout
        Me.TabPage10.SuspendLayout
        Me.TabPage6.SuspendLayout
        Me.SuspendLayout
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TabControl1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Size = New System.Drawing.Size(230, 700)
        Me.SplitContainer1.SplitterDistance = 671
        Me.SplitContainer1.TabIndex = 1
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Multiline = true
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(230, 671)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.AutoScroll = true
        Me.TabPage2.BackColor = System.Drawing.Color.White
        Me.TabPage2.Controls.Add(Me.TabControl2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(222, 645)
        Me.TabPage2.TabIndex = 5
        Me.TabPage2.Text = "기능"
        Me.TabPage2.ToolTipText = "메시지를 구성하기 위한 편리한 기능을 제공합니다"
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Controls.Add(Me.TabPage8)
        Me.TabControl2.Controls.Add(Me.TabPage9)
        Me.TabControl2.Controls.Add(Me.TabPage10)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Location = New System.Drawing.Point(17, 6)
        Me.TabControl2.Multiline = true
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(189, 423)
        Me.TabControl2.TabIndex = 3
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.tbReplace1)
        Me.TabPage7.Controls.Add(Me.Label5)
        Me.TabPage7.Controls.Add(Me.cbReplace1)
        Me.TabPage7.Controls.Add(Me.btnMergeRun2)
        Me.TabPage7.Controls.Add(Me.tbTemplate)
        Me.TabPage7.Controls.Add(Me.cbReplace2)
        Me.TabPage7.Controls.Add(Me.cbReplace3)
        Me.TabPage7.Controls.Add(Me.Label20)
        Me.TabPage7.Controls.Add(Me.TextBox3)
        Me.TabPage7.Controls.Add(Me.tbResultMergeRow)
        Me.TabPage7.Controls.Add(Me.tbReplace2)
        Me.TabPage7.Controls.Add(Me.tbReplace3)
        Me.TabPage7.Location = New System.Drawing.Point(4, 40)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage7.Size = New System.Drawing.Size(181, 379)
        Me.TabPage7.TabIndex = 0
        Me.TabPage7.Text = "문구변환"
        Me.TabPage7.UseVisualStyleBackColor = true
        '
        'tbReplace1
        '
        Me.tbReplace1.Location = New System.Drawing.Point(107, 209)
        Me.tbReplace1.MaxLength = 1
        Me.tbReplace1.Name = "tbReplace1"
        Me.tbReplace1.Size = New System.Drawing.Size(21, 21)
        Me.tbReplace1.TabIndex = 40
        Me.tbReplace1.Text = "A"
        Me.tbReplace1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = true
        Me.Label5.Location = New System.Drawing.Point(10, 51)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(41, 12)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "[문구]"
        '
        'cbReplace1
        '
        Me.cbReplace1.AutoSize = true
        Me.cbReplace1.Checked = true
        Me.cbReplace1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbReplace1.Location = New System.Drawing.Point(47, 212)
        Me.cbReplace1.Name = "cbReplace1"
        Me.cbReplace1.Size = New System.Drawing.Size(58, 16)
        Me.cbReplace1.TabIndex = 35
        Me.cbReplace1.Text = "@1 열"
        Me.cbReplace1.UseVisualStyleBackColor = true
        '
        'btnMergeRun2
        '
        Me.btnMergeRun2.Location = New System.Drawing.Point(35, 323)
        Me.btnMergeRun2.Name = "btnMergeRun2"
        Me.btnMergeRun2.Size = New System.Drawing.Size(110, 23)
        Me.btnMergeRun2.TabIndex = 38
        Me.btnMergeRun2.Text = "실행"
        Me.btnMergeRun2.UseVisualStyleBackColor = True
        '
        'tbTemplate
        '
        Me.tbTemplate.Location = New System.Drawing.Point(11, 70)
        Me.tbTemplate.Multiline = True
        Me.tbTemplate.Name = "tbTemplate"
        Me.tbTemplate.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.tbTemplate.Size = New System.Drawing.Size(159, 130)
        Me.tbTemplate.TabIndex = 37
        Me.tbTemplate.Text = "@1" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "@2" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "@3"
        '
        'cbReplace2
        '
        Me.cbReplace2.AutoSize = True
        Me.cbReplace2.Checked = True
        Me.cbReplace2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbReplace2.Location = New System.Drawing.Point(47, 238)
        Me.cbReplace2.Name = "cbReplace2"
        Me.cbReplace2.Size = New System.Drawing.Size(58, 16)
        Me.cbReplace2.TabIndex = 26
        Me.cbReplace2.Text = "@2 열"
        Me.cbReplace2.UseVisualStyleBackColor = True
        '
        'cbReplace3
        '
        Me.cbReplace3.AutoSize = True
        Me.cbReplace3.Checked = True
        Me.cbReplace3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbReplace3.Location = New System.Drawing.Point(47, 263)
        Me.cbReplace3.Name = "cbReplace3"
        Me.cbReplace3.Size = New System.Drawing.Size(58, 16)
        Me.cbReplace3.TabIndex = 27
        Me.cbReplace3.Text = "@3 열"
        Me.cbReplace3.UseVisualStyleBackColor = True
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(45, 298)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(57, 12)
        Me.Label20.TabIndex = 33
        Me.Label20.Text = "출  력  열"
        '
        'TextBox3
        '
        Me.TextBox3.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox3.Enabled = False
        Me.TextBox3.Location = New System.Drawing.Point(37, 11)
        Me.TextBox3.Multiline = True
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(110, 44)
        Me.TextBox3.TabIndex = 26
        Me.TextBox3.Text = "* 문구에 포함된 @1,@2,@3을 해당 열값으로 변환"
        '
        'tbResultMergeRow
        '
        Me.tbResultMergeRow.Location = New System.Drawing.Point(107, 295)
        Me.tbResultMergeRow.MaxLength = 1
        Me.tbResultMergeRow.Name = "tbResultMergeRow"
        Me.tbResultMergeRow.Size = New System.Drawing.Size(21, 21)
        Me.tbResultMergeRow.TabIndex = 30
        Me.tbResultMergeRow.Text = "D"
        Me.tbResultMergeRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tbReplace2
        '
        Me.tbReplace2.Location = New System.Drawing.Point(107, 234)
        Me.tbReplace2.MaxLength = 1
        Me.tbReplace2.Name = "tbReplace2"
        Me.tbReplace2.Size = New System.Drawing.Size(21, 21)
        Me.tbReplace2.TabIndex = 22
        Me.tbReplace2.Text = "B"
        Me.tbReplace2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tbReplace3
        '
        Me.tbReplace3.Location = New System.Drawing.Point(107, 259)
        Me.tbReplace3.MaxLength = 1
        Me.tbReplace3.Name = "tbReplace3"
        Me.tbReplace3.Size = New System.Drawing.Size(21, 21)
        Me.tbReplace3.TabIndex = 23
        Me.tbReplace3.Text = "C"
        Me.tbReplace3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.Label24)
        Me.TabPage8.Controls.Add(Me.Label25)
        Me.TabPage8.Controls.Add(Me.btnTrimRun)
        Me.TabPage8.Controls.Add(Me.tbResultTrimRow)
        Me.TabPage8.Controls.Add(Me.tbTrimSrcRow)
        Me.TabPage8.Controls.Add(Me.TextBox2)
        Me.TabPage8.Location = New System.Drawing.Point(4, 40)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage8.Size = New System.Drawing.Size(181, 379)
        Me.TabPage8.TabIndex = 1
        Me.TabPage8.Text = "트림"
        Me.TabPage8.UseVisualStyleBackColor = true
        '
        'Label24
        '
        Me.Label24.AutoSize = true
        Me.Label24.Location = New System.Drawing.Point(36, 82)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(65, 12)
        Me.Label24.TabIndex = 38
        Me.Label24.Text = "결과출력열"
        '
        'Label25
        '
        Me.Label25.AutoSize = true
        Me.Label25.Location = New System.Drawing.Point(36, 53)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(41, 12)
        Me.Label25.TabIndex = 37
        Me.Label25.Text = "열선택"
        '
        'btnTrimRun
        '
        Me.btnTrimRun.Location = New System.Drawing.Point(38, 109)
        Me.btnTrimRun.Name = "btnTrimRun"
        Me.btnTrimRun.Size = New System.Drawing.Size(110, 23)
        Me.btnTrimRun.TabIndex = 34
        Me.btnTrimRun.Text = "실행"
        Me.btnTrimRun.UseVisualStyleBackColor = true
        '
        'tbResultTrimRow
        '
        Me.tbResultTrimRow.Location = New System.Drawing.Point(124, 79)
        Me.tbResultTrimRow.MaxLength = 1
        Me.tbResultTrimRow.Name = "tbResultTrimRow"
        Me.tbResultTrimRow.Size = New System.Drawing.Size(21, 21)
        Me.tbResultTrimRow.TabIndex = 36
        Me.tbResultTrimRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tbTrimSrcRow
        '
        Me.tbTrimSrcRow.Location = New System.Drawing.Point(124, 50)
        Me.tbTrimSrcRow.MaxLength = 1
        Me.tbTrimSrcRow.Name = "tbTrimSrcRow"
        Me.tbTrimSrcRow.Size = New System.Drawing.Size(21, 21)
        Me.tbTrimSrcRow.TabIndex = 35
        Me.tbTrimSrcRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox2.Enabled = false
        Me.TextBox2.Location = New System.Drawing.Point(31, 15)
        Me.TextBox2.Multiline = true
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = true
        Me.TextBox2.Size = New System.Drawing.Size(110, 17)
        Me.TextBox2.TabIndex = 27
        Me.TextBox2.Text = "* 양끝의 공백 제거"
        '
        'TabPage9
        '
        Me.TabPage9.Controls.Add(Me.Label28)
        Me.TabPage9.Controls.Add(Me.Label29)
        Me.TabPage9.Controls.Add(Me.btnNumberOnlyRun)
        Me.TabPage9.Controls.Add(Me.tbNumberOnlyOutRow)
        Me.TabPage9.Controls.Add(Me.tbNumberOnlySrcRow)
        Me.TabPage9.Controls.Add(Me.TextBox7)
        Me.TabPage9.Location = New System.Drawing.Point(4, 40)
        Me.TabPage9.Name = "TabPage9"
        Me.TabPage9.Size = New System.Drawing.Size(181, 379)
        Me.TabPage9.TabIndex = 2
        Me.TabPage9.Text = "숫자추출"
        Me.TabPage9.UseVisualStyleBackColor = true
        '
        'Label28
        '
        Me.Label28.AutoSize = true
        Me.Label28.Location = New System.Drawing.Point(36, 82)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(65, 12)
        Me.Label28.TabIndex = 44
        Me.Label28.Text = "결과출력열"
        '
        'Label29
        '
        Me.Label29.AutoSize = true
        Me.Label29.Location = New System.Drawing.Point(36, 53)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(41, 12)
        Me.Label29.TabIndex = 43
        Me.Label29.Text = "열선택"
        '
        'btnNumberOnlyRun
        '
        Me.btnNumberOnlyRun.Location = New System.Drawing.Point(38, 109)
        Me.btnNumberOnlyRun.Name = "btnNumberOnlyRun"
        Me.btnNumberOnlyRun.Size = New System.Drawing.Size(110, 23)
        Me.btnNumberOnlyRun.TabIndex = 40
        Me.btnNumberOnlyRun.Text = "실행"
        Me.btnNumberOnlyRun.UseVisualStyleBackColor = true
        '
        'tbNumberOnlyOutRow
        '
        Me.tbNumberOnlyOutRow.Location = New System.Drawing.Point(124, 79)
        Me.tbNumberOnlyOutRow.MaxLength = 1
        Me.tbNumberOnlyOutRow.Name = "tbNumberOnlyOutRow"
        Me.tbNumberOnlyOutRow.Size = New System.Drawing.Size(21, 21)
        Me.tbNumberOnlyOutRow.TabIndex = 42
        Me.tbNumberOnlyOutRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tbNumberOnlySrcRow
        '
        Me.tbNumberOnlySrcRow.Location = New System.Drawing.Point(124, 50)
        Me.tbNumberOnlySrcRow.MaxLength = 1
        Me.tbNumberOnlySrcRow.Name = "tbNumberOnlySrcRow"
        Me.tbNumberOnlySrcRow.Size = New System.Drawing.Size(21, 21)
        Me.tbNumberOnlySrcRow.TabIndex = 41
        Me.tbNumberOnlySrcRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox7
        '
        Me.TextBox7.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.TextBox7.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox7.Enabled = false
        Me.TextBox7.Location = New System.Drawing.Point(35, 12)
        Me.TextBox7.Multiline = true
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.ReadOnly = true
        Me.TextBox7.Size = New System.Drawing.Size(110, 30)
        Me.TextBox7.TabIndex = 39
        Me.TextBox7.Text = "* 기호와 문자를 제거하고 숫자만 추출"
        '
        'TabPage10
        '
        Me.TabPage10.Controls.Add(Me.Label26)
        Me.TabPage10.Controls.Add(Me.Label27)
        Me.TabPage10.Controls.Add(Me.btnByteLengthRun)
        Me.TabPage10.Controls.Add(Me.tbByteLengthOutRow)
        Me.TabPage10.Controls.Add(Me.tbByteLengthSrcRow)
        Me.TabPage10.Controls.Add(Me.TextBox6)
        Me.TabPage10.Location = New System.Drawing.Point(4, 40)
        Me.TabPage10.Name = "TabPage10"
        Me.TabPage10.Size = New System.Drawing.Size(181, 379)
        Me.TabPage10.TabIndex = 3
        Me.TabPage10.Text = "Byte(길이)"
        Me.TabPage10.UseVisualStyleBackColor = true
        '
        'Label26
        '
        Me.Label26.AutoSize = true
        Me.Label26.Location = New System.Drawing.Point(36, 82)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(65, 12)
        Me.Label26.TabIndex = 44
        Me.Label26.Text = "결과출력열"
        '
        'Label27
        '
        Me.Label27.AutoSize = true
        Me.Label27.Location = New System.Drawing.Point(36, 53)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(41, 12)
        Me.Label27.TabIndex = 43
        Me.Label27.Text = "열선택"
        '
        'btnByteLengthRun
        '
        Me.btnByteLengthRun.Location = New System.Drawing.Point(38, 109)
        Me.btnByteLengthRun.Name = "btnByteLengthRun"
        Me.btnByteLengthRun.Size = New System.Drawing.Size(110, 23)
        Me.btnByteLengthRun.TabIndex = 40
        Me.btnByteLengthRun.Text = "실행"
        Me.btnByteLengthRun.UseVisualStyleBackColor = true
        '
        'tbByteLengthOutRow
        '
        Me.tbByteLengthOutRow.Location = New System.Drawing.Point(124, 79)
        Me.tbByteLengthOutRow.MaxLength = 1
        Me.tbByteLengthOutRow.Name = "tbByteLengthOutRow"
        Me.tbByteLengthOutRow.Size = New System.Drawing.Size(21, 21)
        Me.tbByteLengthOutRow.TabIndex = 42
        Me.tbByteLengthOutRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tbByteLengthSrcRow
        '
        Me.tbByteLengthSrcRow.Location = New System.Drawing.Point(124, 50)
        Me.tbByteLengthSrcRow.MaxLength = 1
        Me.tbByteLengthSrcRow.Name = "tbByteLengthSrcRow"
        Me.tbByteLengthSrcRow.Size = New System.Drawing.Size(21, 21)
        Me.tbByteLengthSrcRow.TabIndex = 41
        Me.tbByteLengthSrcRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox6
        '
        Me.TextBox6.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.TextBox6.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox6.Enabled = false
        Me.TextBox6.Location = New System.Drawing.Point(27, 11)
        Me.TextBox6.Multiline = true
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.ReadOnly = true
        Me.TextBox6.Size = New System.Drawing.Size(131, 17)
        Me.TextBox6.TabIndex = 39
        Me.TextBox6.Text = "* 바이트(Byte) 구하기"
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.Label33)
        Me.TabPage6.Controls.Add(Me.btnCelTypeRun)
        Me.TabPage6.Controls.Add(Me.tbCelTypeSrcRow)
        Me.TabPage6.Controls.Add(Me.TextBox8)
        Me.TabPage6.Location = New System.Drawing.Point(4, 40)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(181, 379)
        Me.TabPage6.TabIndex = 4
        Me.TabPage6.Text = "셀서식(텍스트)"
        Me.TabPage6.UseVisualStyleBackColor = true
        '
        'Label33
        '
        Me.Label33.AutoSize = true
        Me.Label33.Location = New System.Drawing.Point(36, 53)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(41, 12)
        Me.Label33.TabIndex = 49
        Me.Label33.Text = "열선택"
        '
        'btnCelTypeRun
        '
        Me.btnCelTypeRun.Location = New System.Drawing.Point(38, 109)
        Me.btnCelTypeRun.Name = "btnCelTypeRun"
        Me.btnCelTypeRun.Size = New System.Drawing.Size(110, 23)
        Me.btnCelTypeRun.TabIndex = 46
        Me.btnCelTypeRun.Text = "실행"
        Me.btnCelTypeRun.UseVisualStyleBackColor = true
        '
        'tbCelTypeSrcRow
        '
        Me.tbCelTypeSrcRow.Location = New System.Drawing.Point(124, 50)
        Me.tbCelTypeSrcRow.MaxLength = 1
        Me.tbCelTypeSrcRow.Name = "tbCelTypeSrcRow"
        Me.tbCelTypeSrcRow.Size = New System.Drawing.Size(21, 21)
        Me.tbCelTypeSrcRow.TabIndex = 47
        Me.tbCelTypeSrcRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox8
        '
        Me.TextBox8.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.TextBox8.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox8.Enabled = false
        Me.TextBox8.Location = New System.Drawing.Point(25, 16)
        Me.TextBox8.Multiline = true
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.ReadOnly = true
        Me.TextBox8.Size = New System.Drawing.Size(144, 26)
        Me.TextBox8.TabIndex = 45
        Me.TextBox8.Text = "* 셀을 텍스트형으로 변환"
        '
        'Label1
        '
        Me.Label1.AutoSize = true
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label1.Location = New System.Drawing.Point(0, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ⓒ2014 ThanksXL. All rights reserved."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'UserControl1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7!, 12!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "UserControl1"
        Me.Size = New System.Drawing.Size(230, 700)
        Me.SplitContainer1.Panel1.ResumeLayout(false)
        Me.SplitContainer1.Panel2.ResumeLayout(false)
        Me.SplitContainer1.Panel2.PerformLayout
        CType(Me.SplitContainer1,System.ComponentModel.ISupportInitialize).EndInit
        Me.SplitContainer1.ResumeLayout(false)
        Me.TabControl1.ResumeLayout(false)
        Me.TabPage2.ResumeLayout(false)
        Me.TabControl2.ResumeLayout(false)
        Me.TabPage7.ResumeLayout(false)
        Me.TabPage7.PerformLayout
        Me.TabPage8.ResumeLayout(false)
        Me.TabPage8.PerformLayout
        Me.TabPage9.ResumeLayout(false)
        Me.TabPage9.PerformLayout
        Me.TabPage10.ResumeLayout(false)
        Me.TabPage10.PerformLayout
        Me.TabPage6.ResumeLayout(false)
        Me.TabPage6.PerformLayout
        Me.ResumeLayout(false)

End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents tbReplace1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cbReplace1 As System.Windows.Forms.CheckBox
    Friend WithEvents btnMergeRun2 As System.Windows.Forms.Button
    Friend WithEvents tbTemplate As System.Windows.Forms.TextBox
    Friend WithEvents cbReplace2 As System.Windows.Forms.CheckBox
    Friend WithEvents cbReplace3 As System.Windows.Forms.CheckBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents tbResultMergeRow As System.Windows.Forms.TextBox
    Friend WithEvents tbReplace2 As System.Windows.Forms.TextBox
    Friend WithEvents tbReplace3 As System.Windows.Forms.TextBox
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents btnTrimRun As System.Windows.Forms.Button
    Friend WithEvents tbResultTrimRow As System.Windows.Forms.TextBox
    Friend WithEvents tbTrimSrcRow As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TabPage9 As System.Windows.Forms.TabPage
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents btnNumberOnlyRun As System.Windows.Forms.Button
    Friend WithEvents tbNumberOnlyOutRow As System.Windows.Forms.TextBox
    Friend WithEvents tbNumberOnlySrcRow As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TabPage10 As System.Windows.Forms.TabPage
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents btnByteLengthRun As System.Windows.Forms.Button
    Friend WithEvents tbByteLengthOutRow As System.Windows.Forms.TextBox
    Friend WithEvents tbByteLengthSrcRow As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents btnCelTypeRun As System.Windows.Forms.Button
    Friend WithEvents tbCelTypeSrcRow As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
