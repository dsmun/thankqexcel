Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        tbXlmsg.Checked = True
        tbXlmsg.Label = "패널 숨기기"
    End Sub

    Private Sub tbXlmsg_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles tbXlmsg.Click

        Try

            If tbXlmsg.Checked Then

                For i = 0 To Globals.ThisAddIn.CustomTaskPanes.Count - 1

                    If Globals.ThisAddIn.CustomTaskPanes(i).Title = "Excel-Function Beta" Then
                        Globals.ThisAddIn.CustomTaskPanes(i).Visible = True
                        tbXlmsg.Label = "패널 숨기기"
                        Exit For
                    End If

                Next i

            Else

                For i = 0 To Globals.ThisAddIn.CustomTaskPanes.Count - 1

                    If Globals.ThisAddIn.CustomTaskPanes(i).Title = "Excel-Function Beta" Then
                        Globals.ThisAddIn.CustomTaskPanes(i).Visible = False
                        tbXlmsg.Label = "패널 보이기"
                        Exit For
                    End If

                Next i


                'Globals.ThisAddIn.CustomTaskPanes(0).Visible = False
                'tbXlmsg.Label = "패널 보이기"

            End If


        Catch ex As Exception

            MsgBox(ex.Message)

        End Try


    End Sub
End Class
