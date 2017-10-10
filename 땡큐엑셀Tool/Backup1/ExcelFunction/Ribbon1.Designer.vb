Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 클래스 컴퍼지션 디자이너 지원에 필요합니다.
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '이 호출은 구성 요소 디자이너에 필요합니다.
        InitializeComponent()

    End Sub

    'Component는 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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

    '구성 요소 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 구성 요소 디자이너에 필요합니다.
    '구성 요소 디자이너를 사용하여 수정할 수 있습니다.
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button중간공백하나로 = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.TabExcelFunction = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.tbXlmsg = Me.Factory.CreateRibbonToggleButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button단어카운트 = Me.Factory.CreateRibbonButton
        Me.Button텍스스수식계산 = Me.Factory.CreateRibbonButton
        Me.Button연속공백하나로 = Me.Factory.CreateRibbonButton
        Me.Button바탕색합계 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button네이버 = Me.Factory.CreateRibbonButton
        Me.Button티스토리 = Me.Factory.CreateRibbonButton
        Me.Button티스토리블로그 = Me.Factory.CreateRibbonButton
        Me.Button네이버블로그 = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button210월 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.TabExcelFunction.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        '
        'Button중간공백하나로
        '
        Me.Button중간공백하나로.Label = "연속공백하나로"
        Me.Button중간공백하나로.Name = "Button중간공백하나로"
        '
        'Button2
        '
        Me.Button2.Label = "10월"
        Me.Button2.Name = "Button2"
        '
        'TabExcelFunction
        '
        Me.TabExcelFunction.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TabExcelFunction.Groups.Add(Me.Group1)
        Me.TabExcelFunction.Groups.Add(Me.Group2)
        Me.TabExcelFunction.Groups.Add(Me.Group3)
        Me.TabExcelFunction.Label = "땡큐엑셀Tool"
        Me.TabExcelFunction.Name = "TabExcelFunction"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.tbXlmsg)
        Me.Group1.Label = "Excel기능"
        Me.Group1.Name = "Group1"
        '
        'tbXlmsg
        '
        Me.tbXlmsg.Label = "숨기기"
        Me.tbXlmsg.Name = "tbXlmsg"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button단어카운트)
        Me.Group2.Items.Add(Me.Button텍스스수식계산)
        Me.Group2.Items.Add(Me.Button연속공백하나로)
        Me.Group2.Items.Add(Me.Button바탕색합계)
        Me.Group2.Label = "Group2"
        Me.Group2.Name = "Group2"
        '
        'Button단어카운트
        '
        Me.Button단어카운트.Label = "단어카운트"
        Me.Button단어카운트.Name = "Button단어카운트"
        '
        'Button텍스스수식계산
        '
        Me.Button텍스스수식계산.Label = "텍스트수식계산"
        Me.Button텍스스수식계산.Name = "Button텍스스수식계산"
        '
        'Button연속공백하나로
        '
        Me.Button연속공백하나로.Label = "연속공백하나로"
        Me.Button연속공백하나로.Name = "Button연속공백하나로"
        '
        'Button바탕색합계
        '
        Me.Button바탕색합계.Label = "같은바탕색합계"
        Me.Button바탕색합계.Name = "Button바탕색합계"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button네이버)
        Me.Group3.Items.Add(Me.Button티스토리)
        Me.Group3.Items.Add(Me.Button티스토리블로그)
        Me.Group3.Items.Add(Me.Button네이버블로그)
        Me.Group3.Items.Add(Me.Button1)
        Me.Group3.Items.Add(Me.Button210월)
        Me.Group3.Items.Add(Me.Button3)
        Me.Group3.Items.Add(Me.Button4)
        Me.Group3.Items.Add(Me.Button5)
        Me.Group3.Label = "Group3"
        Me.Group3.Name = "Group3"
        '
        'Button네이버
        '
        Me.Button네이버.Label = "네이버"
        Me.Button네이버.Name = "Button네이버"
        '
        'Button티스토리
        '
        Me.Button티스토리.Label = "티스토리"
        Me.Button티스토리.Name = "Button티스토리"
        '
        'Button티스토리블로그
        '
        Me.Button티스토리블로그.Label = "티스토리블로그"
        Me.Button티스토리블로그.Name = "Button티스토리블로그"
        '
        'Button네이버블로그
        '
        Me.Button네이버블로그.Label = "네이버블로그"
        Me.Button네이버블로그.Name = "Button네이버블로그"
        '
        'Button1
        '
        Me.Button1.Label = "엑셀"
        Me.Button1.Name = "Button1"
        '
        'Button210월
        '
        Me.Button210월.Label = "10월에(엑셀버전)"
        Me.Button210월.Name = "Button210월"
        '
        'Button3
        '
        Me.Button3.Label = "10월에(PPT버전)"
        Me.Button3.Name = "Button3"
        '
        'Button4
        '
        Me.Button4.Label = "10월의(Word버전)"
        Me.Button4.Name = "Button4"
        '
        'Button5
        '
        Me.Button5.Label = "불러오기"
        Me.Button5.Name = "Button5"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.TabExcelFunction)
        Me.TabExcelFunction.ResumeLayout(False)
        Me.TabExcelFunction.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()

    End Sub

    Friend WithEvents TabExcelFunction As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents tbXlmsg As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button단어카운트 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button텍스스수식계산 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button네이버 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button티스토리 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button티스토리블로그 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button네이버블로그 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button연속공백하나로 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button중간공백하나로 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button바탕색합계 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button210월 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
