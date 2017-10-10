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
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Button범위보내기 = New System.Windows.Forms.Button()
        Me.Button서버소켓종료 = New System.Windows.Forms.Button()
        Me.Button서버소켓시작 = New System.Windows.Forms.Button()
        Me.Button보내기 = New System.Windows.Forms.Button()
        Me.TextBox메시지 = New System.Windows.Forms.TextBox()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Button접속종료 = New System.Windows.Forms.Button()
        Me.Button서버접속 = New System.Windows.Forms.Button()
        Me.ButtonClient보내기 = New System.Windows.Forms.Button()
        Me.TextBoxClient메시지 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
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
        Me.SplitContainer1.Size = New System.Drawing.Size(1108, 700)
        Me.SplitContainer1.SplitterDistance = 661
        Me.SplitContainer1.TabIndex = 1
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1108, 661)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.WebBrowser1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(1100, 635)
        Me.TabPage1.TabIndex = 6
        Me.TabPage1.Text = "Brower"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebBrowser1.Location = New System.Drawing.Point(0, 0)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(1100, 635)
        Me.WebBrowser1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Button범위보내기)
        Me.TabPage2.Controls.Add(Me.Button서버소켓종료)
        Me.TabPage2.Controls.Add(Me.Button서버소켓시작)
        Me.TabPage2.Controls.Add(Me.Button보내기)
        Me.TabPage2.Controls.Add(Me.TextBox메시지)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(1100, 635)
        Me.TabPage2.TabIndex = 7
        Me.TabPage2.Text = "Server"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Button범위보내기
        '
        Me.Button범위보내기.Location = New System.Drawing.Point(486, 35)
        Me.Button범위보내기.Name = "Button범위보내기"
        Me.Button범위보내기.Size = New System.Drawing.Size(114, 81)
        Me.Button범위보내기.TabIndex = 4
        Me.Button범위보내기.Text = "범위보내기"
        Me.Button범위보내기.UseVisualStyleBackColor = True
        '
        'Button서버소켓종료
        '
        Me.Button서버소켓종료.Location = New System.Drawing.Point(123, 6)
        Me.Button서버소켓종료.Name = "Button서버소켓종료"
        Me.Button서버소켓종료.Size = New System.Drawing.Size(110, 23)
        Me.Button서버소켓종료.TabIndex = 3
        Me.Button서버소켓종료.Text = "서버소켓 종료"
        Me.Button서버소켓종료.UseVisualStyleBackColor = True
        '
        'Button서버소켓시작
        '
        Me.Button서버소켓시작.Location = New System.Drawing.Point(3, 6)
        Me.Button서버소켓시작.Name = "Button서버소켓시작"
        Me.Button서버소켓시작.Size = New System.Drawing.Size(114, 23)
        Me.Button서버소켓시작.TabIndex = 2
        Me.Button서버소켓시작.Text = "서버소켓 시작"
        Me.Button서버소켓시작.UseVisualStyleBackColor = True
        '
        'Button보내기
        '
        Me.Button보내기.Location = New System.Drawing.Point(353, 35)
        Me.Button보내기.Name = "Button보내기"
        Me.Button보내기.Size = New System.Drawing.Size(114, 81)
        Me.Button보내기.TabIndex = 1
        Me.Button보내기.Text = "보내기"
        Me.Button보내기.UseVisualStyleBackColor = True
        '
        'TextBox메시지
        '
        Me.TextBox메시지.Location = New System.Drawing.Point(3, 35)
        Me.TextBox메시지.Multiline = True
        Me.TextBox메시지.Name = "TextBox메시지"
        Me.TextBox메시지.Size = New System.Drawing.Size(344, 81)
        Me.TextBox메시지.TabIndex = 0
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.Button접속종료)
        Me.TabPage3.Controls.Add(Me.Button서버접속)
        Me.TabPage3.Controls.Add(Me.ButtonClient보내기)
        Me.TabPage3.Controls.Add(Me.TextBoxClient메시지)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(1100, 635)
        Me.TabPage3.TabIndex = 8
        Me.TabPage3.Text = "Client"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'Button접속종료
        '
        Me.Button접속종료.Location = New System.Drawing.Point(127, 3)
        Me.Button접속종료.Name = "Button접속종료"
        Me.Button접속종료.Size = New System.Drawing.Size(118, 23)
        Me.Button접속종료.TabIndex = 3
        Me.Button접속종료.Text = "접속종료"
        Me.Button접속종료.UseVisualStyleBackColor = True
        '
        'Button서버접속
        '
        Me.Button서버접속.Location = New System.Drawing.Point(3, 3)
        Me.Button서버접속.Name = "Button서버접속"
        Me.Button서버접속.Size = New System.Drawing.Size(118, 23)
        Me.Button서버접속.TabIndex = 2
        Me.Button서버접속.Text = "서버접속"
        Me.Button서버접속.UseVisualStyleBackColor = True
        '
        'ButtonClient보내기
        '
        Me.ButtonClient보내기.Location = New System.Drawing.Point(929, 32)
        Me.ButtonClient보내기.Name = "ButtonClient보내기"
        Me.ButtonClient보내기.Size = New System.Drawing.Size(109, 77)
        Me.ButtonClient보내기.TabIndex = 1
        Me.ButtonClient보내기.Text = "보내기"
        Me.ButtonClient보내기.UseVisualStyleBackColor = True
        '
        'TextBoxClient메시지
        '
        Me.TextBoxClient메시지.Location = New System.Drawing.Point(3, 32)
        Me.TextBoxClient메시지.Multiline = True
        Me.TextBoxClient메시지.Name = "TextBoxClient메시지"
        Me.TextBoxClient메시지.Size = New System.Drawing.Size(920, 77)
        Me.TextBoxClient메시지.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label1.Location = New System.Drawing.Point(0, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(207, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ⓒ2014 땡큐엑셀. All rights reserved."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'UserControl1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "UserControl1"
        Me.Size = New System.Drawing.Size(1108, 700)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.ResumeLayout(False)

End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label

    Public Sub New()

        ' 이 호출은 디자이너에 필요합니다.
        InitializeComponent()

        ' InitializeComponent() 호출 뒤에 초기화 코드를 추가하십시오.

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Button보내기 As System.Windows.Forms.Button
    Friend WithEvents TextBox메시지 As System.Windows.Forms.TextBox
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents ButtonClient보내기 As System.Windows.Forms.Button
    Friend WithEvents TextBoxClient메시지 As System.Windows.Forms.TextBox
    Friend WithEvents Button서버소켓종료 As System.Windows.Forms.Button
    Friend WithEvents Button서버소켓시작 As System.Windows.Forms.Button
    Friend WithEvents Button접속종료 As System.Windows.Forms.Button
    Friend WithEvents Button서버접속 As System.Windows.Forms.Button
    Friend WithEvents Button범위보내기 As System.Windows.Forms.Button
End Class
