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
        Me.TabExcelFunction = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.tbXlmsg = Me.Factory.CreateRibbonToggleButton
        Me.TabExcelFunction.SuspendLayout()
        Me.Group1.SuspendLayout()
        '
        'TabExcelFunction
        '
        Me.TabExcelFunction.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TabExcelFunction.Groups.Add(Me.Group1)
        Me.TabExcelFunction.Label = "ThanksXL"
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
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.TabExcelFunction)
        Me.TabExcelFunction.ResumeLayout(False)
        Me.TabExcelFunction.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()

    End Sub

    Friend WithEvents TabExcelFunction As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents tbXlmsg As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
