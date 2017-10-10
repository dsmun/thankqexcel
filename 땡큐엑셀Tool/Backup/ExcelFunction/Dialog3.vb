Imports System.Windows.Forms

Public Class Dialog3

    '정지 or 진행
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        If OK_Button.Text = "진행" Then
            Globals.ThisAddIn.gsStatus = "진행"
            Me.Text = "진행"
            OK_Button.Text = "정지"
        Else
            Globals.ThisAddIn.gsStatus = "정지"
            Me.Text = "정지"
            OK_Button.Text = "진행"
        End If


        'Me.DialogResult = System.Windows.Forms.DialogResult.OK
        'Me.Close()
    End Sub

    '취소
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click

        Globals.ThisAddIn.gsStatus = "취소"
        Me.Text = "취소"

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub


    Private Sub lblPercent_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPercent.TextChanged


        If Val(lblPercent.Text) >= 100 Then

            Globals.ThisAddIn.gsStatus = "완료"
            Me.Text = "완료"

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
            ' or
            'OK_Button_Click(sender, EventArgs.Empty)

        End If


    End Sub

    Private Sub Dialog3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
