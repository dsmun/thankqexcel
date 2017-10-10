Imports System.Windows.Forms

Public Class Dialog1

    Dim rngR As Excel.Range
    Dim rngC As Excel.Range


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Dim colorSum As Double



        '범위의 셀을 하나 하나씩 꺼내어
        For Each R In rngR
            '바탕색이 같은것만 더하기
            If R.Interior.ColorIndex = rngC.Interior.ColorIndex Then
                colorSum = colorSum + R.value
            End If
        Next

        Globals.ThisAddIn.Application.ActiveCell.Value = colorSum



        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub



    Private Sub Button범위선_Click(sender As System.Object, e As System.EventArgs) Handles Button범위선.Click
        'Dim rngR As Excel.Range

        '범위선택
        rngR = Globals.ThisAddIn.Application.InputBox("범위선택", "범위", Type:=8)

    End Sub


    Private Sub Button바탕색선택_Click(sender As System.Object, e As System.EventArgs) Handles Button바탕색선택.Click
        'Dim rngC As Excel.Range

        '바탕색선택
        rngC = Globals.ThisAddIn.Application.InputBox("바탕색선택", "바탕색", Type:=8)
    End Sub


End Class
