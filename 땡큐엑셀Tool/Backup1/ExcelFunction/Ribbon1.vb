Imports Microsoft.Office.Tools.Ribbon
Imports System.Text.RegularExpressions.Regex 'for regex

Public Class Ribbon1


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        tbXlmsg.Checked = True
        tbXlmsg.Label = "패널 숨기기"
    End Sub

    Private Sub tbXlmsg_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles tbXlmsg.Click

        Try

            If tbXlmsg.Checked Then

                For i = 0 To Globals.ThisAddIn.CustomTaskPanes.Count - 1

                    If Globals.ThisAddIn.CustomTaskPanes(i).Title = "땡큐엑셀Tool Beta" Then
                        Globals.ThisAddIn.CustomTaskPanes(i).Visible = True
                        tbXlmsg.Label = "패널 숨기기"
                        Exit For
                    End If

                Next i

            Else

                For i = 0 To Globals.ThisAddIn.CustomTaskPanes.Count - 1

                    If Globals.ThisAddIn.CustomTaskPanes(i).Title = "땡큐엑셀Tool Beta" Then
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


    '단어카운트
    Private Sub Button단어카운트_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button단어카운트.Click

        'MsgBox("단어카운트")


        Dim rng As Excel.Range
        'Dim iCnt As Integer
        Dim v범위
        Dim D
        Dim j
        Dim strA
        Dim r As Excel.Range
        Dim str셀값 As String



        '1. 범위를 루프돌며 사전에 기록
        D = CreateObject("Scripting.Dictionary")    '딕셔너리 선언

        '범위선택
        rng = Globals.ThisAddIn.Application.InputBox("범위선택... 결과는 ActiveCell부터 M행,2열 배열로 출력됩니다. ,", "범위", Type:=8)



        '갯수
        'rng = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbTrimSrcRow.Text & ":" & tbTrimSrcRow.Text)
        'rng = Globals.ThisAddIn.Application.Selection
        'iCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountA(rng)




        '연속된 공백을 제거하기 위해 정규표현식 사용
        Dim strInput As String 
        Dim pattern As String = "\s+"
        Dim replacement As String = " "
        Dim strResult As String
        Dim rgx As New RegularExpressions.Regex(pattern)

        '파싱
        For Each r In rng

            '널이 아닌셀만 파악
            If r.Value <> "" Then

                'MsgBox(r.Value)
                'Split함수를 사용하여 공백으로 단어 분리
                strInput = r.Value
                strResult = rgx.Replace(strInput, replacement)

                v범위 = Globals.ThisAddIn.Application.Transpose(Split(strResult, " ")) '행,열 변환


                'MsgBox(UBound(v범위, 1))
                'MsgBox(v범위(1, 1))


                For i = 1 To UBound(v범위, 1)
                    If Not D.exists(v범위(i, 1)) Then
                        '처음출현하는 단어
                        D.Add(v범위(i, 1), 1)
                    Else
                        '재출현하는 단어는 +1
                        D.Item(v범위(i, 1)) = D.Item(v범위(i, 1)) + 1
                    End If
                Next

            End If


        Next



        '출력
        Dim key As Object
        j = 1
        strA = ""

        'For Each key In D.Keys
        '    Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells(j, "I") = key
        '    Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells(j, "J") = D(key)

        '    j = j + 1
        'Next key

        'ActiveCell부터 출력
        j = 0
        For Each key In D.Keys
            Globals.ThisAddIn.Application.ActiveCell.Offset(j, 0).Value = key
            Globals.ThisAddIn.Application.ActiveCell.Offset(j, 1).Value = D(key)

            j = j + 1
        Next key

        MsgBox("완료되었습니다.", MsgBoxStyle.Information, "땡큐엑셀")



    End Sub



    Private Sub Button네이버_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button네이버.Click
        Globals.ThisAddIn.taskPaneView.goHomePage1("http://www.naver.com/")

    End Sub

    Private Sub Button티스토리_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button티스토리.Click
        'taskPaneView.goHomePage2("http://thank-q.tistory.com")
        Globals.ThisAddIn.taskPaneView.goHomePage1("http://www.tistory.com")
    End Sub

    Private Sub Button티스토리블로그_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button티스토리블로그.Click
        Globals.ThisAddIn.taskPaneView.goHomePage1("http://thank-q.tistory.com")
    End Sub

    Private Sub Button네이버블로그_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button네이버블로그.Click
        Dim strUrl As String
        'Globals.ThisAddIn.taskPaneView.goHomePage1("http://blog.naver.com/dsmun")
        strUrl = "https://onedrive.live.com/edit.aspx?cid=0069e56308e643b3&page=view&resid=69E56308E643B3!2442&parId=69E56308E643B3!103&app=Excel"
        Globals.ThisAddIn.taskPaneView.goHomePage1(strUrl)
    End Sub


    Private Sub Button연속공백하나로_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button연속공백하나로.Click

        Dim rng As Excel.Range
        'Dim iCnt As Integer
        'Dim v범위
        'Dim D
        Dim j As Integer
        'Dim strA
        Dim r As Excel.Range
        'Dim str셀값 As String



        '1. 범위를 루프돌며 사전에 기록
        'D = CreateObject("Scripting.Dictionary")    '딕셔너리 선언

        '범위선택
        rng = Globals.ThisAddIn.Application.InputBox("범위선택... 결과는 ActiveCell부터 출력됩니다. ,", "범위", Type:=8)




        '연속된 공백을 제거하기 위해 정규표현식 사용
        Dim strInput As String
        Dim pattern As String = "\s+"
        Dim replacement As String = " "
        Dim strResult As String
        Dim rgx As New RegularExpressions.Regex(pattern)

        j = 0
        '파싱
        For Each r In rng

            '널이 아닌셀만 진행
            If r.Value <> "" Then

                'MsgBox(r.Value)
                '정규식을 사용하여 중간공백 1개로
                strInput = r.Value
                strResult = rgx.Replace(strInput, replacement)

                Globals.ThisAddIn.Application.ActiveCell.Offset(j, 0).Value = strResult

                'r.Offset(0, 1).Value = strResult
            Else
                Globals.ThisAddIn.Application.ActiveCell.Offset(j, 0).Value = ""
            End If


            j = j + 1

        Next


        MsgBox("완료되었습니다.", MsgBoxStyle.Information, "땡큐엑셀")

    End Sub

    Private Sub Button텍스스수식계산_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button텍스스수식계산.Click

    End Sub


    Private Sub Button바탕색합계_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button바탕색합계.Click
        Dim rngR As Excel.Range
        Dim rngC As Excel.Range
        Dim colorSum As Double


        '범위선택
        rngR = Globals.ThisAddIn.Application.InputBox("범위선택(1/2)", "범위", Type:=8)

        If rngR Is Nothing Then

            MsgBox("선택된 범위가 없습니다.")
            Exit Sub

        End If

        '바탕색선택
        rngC = Globals.ThisAddIn.Application.InputBox("바탕색셀선택(2/2)", "바탕색", Type:=8)
        If rngC Is Nothing Then

            MsgBox("선택된 범위가 없습니다.")
            Exit Sub

        End If


        '범위의 셀을 하나 하나씩 꺼내어
        For Each R In rngR
            '바탕색이 같은것만 더하기
            If R.Interior.ColorIndex = rngC.Interior.ColorIndex Then
                colorSum = colorSum + R.value
            End If
        Next


        Globals.ThisAddIn.Application.ActiveCell.Value = colorSum




        'Dim dlg As Dialog1

        'dlg = New Dialog1

        'dlg.ShowDialog()


    End Sub

    Private Sub Button1_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button1.Click
        Dim strUrl As String
        strUrl = "https://1drv.ms/x/s!ArND5ghj5WkAkwoOupaEczi3NmOV"
        Globals.ThisAddIn.taskPaneView.goHomePage1(strUrl)
    End Sub



    Private Sub Button2_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button210월.Click, Button2.Click
        Dim strUrl As String
        strUrl = "https://1drv.ms/x/s!ArND5ghj5WkAkw3l-J7tYUEsEq1N"
        Globals.ThisAddIn.taskPaneView.goHomePage1(strUrl)
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button3.Click
        Dim strUrl As String
        strUrl = "https://1drv.ms/p/s!ArND5ghj5WkAkw8ekUUQYuZyT-m6"
        Globals.ThisAddIn.taskPaneView.goHomePage1(strUrl)
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button4.Click
        Dim strUrl As String
        strUrl = "https://1drv.ms/w/s!ArND5ghj5WkAkxFVvIB2T02_n1h9"
        Globals.ThisAddIn.taskPaneView.goHomePage1(strUrl)
    End Sub




    Private Sub Button5_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button5.Click

        Dim dlg As Form1
        dlg = New Form1


        Dim strUrl As String
        strUrl = "https://1drv.ms/x/s!ArND5ghj5WkAkw3l-J7tYUEsEq1N"
        dlg.WebBrowser1.Navigate(strUrl)

        'dlg.ShowDialog()
        dlg.Show()


    End Sub


End Class
