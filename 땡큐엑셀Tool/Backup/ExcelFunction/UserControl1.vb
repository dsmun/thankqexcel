Imports System.Threading
Public Class UserControl1

    Class ParameterMerge
        'Friend StrArg As String
        'Friend RetVal As Boolean
        Friend strResultRow As String
        Friend strTemplate As String
        Friend strRep1 As String
        Friend strRep2 As String
        Friend strRep3 As String
        Friend iThreadNum As Integer

        'Friend iMod() As Integer

        'Sub ReArr(ByVal i As Integer)
        '    ReDim iMod(0 To i)
        'End Sub

    End Class



    Delegate Sub cbDBProgress(ByVal iVal1 As Integer, ByVal iVal2 As Integer)
    Public cbDBInsertFunc As cbDBProgress


    Friend gsStatus As String

    Dim lockMergeObj As New Object()
    Dim iModNumMergeThread As Integer = 0 '머지 쓰레드 공유변수
    Dim iSum As Integer = 0 '진행갯수
    Dim iThreadResult() As Integer

    Dim dlgDB As New Dialog3
    Dim lockSumObj As New Object()



    'Delegrate(대리자)
    Public Sub cbDBInsertMethod(ByVal iVal1 As Integer, ByVal iVal2 As Integer)

        '공유변수 Lock
        SyncLock lockSumObj
            iSum = iSum + iVal1
            dlgDB.pbPro.Value = CInt(iSum / iVal2 * 100)
            dlgDB.lblPercent.Text = CInt(iSum / iVal2 * 100).ToString
        End SyncLock

    End Sub



    ' Tab컨트롤 Visible True / False
    Public Sub setTabControl(ByVal bOpt1 As Boolean)

        TabControl1.Visible = bOpt1

    End Sub


    'Config 파일명 Get
    Public Sub getCFGFile(ByRef strOpt1 As String)

        '설정파일은 일단 C:\에 위치
        'strOpt1 = tbCFG.Text
        strOpt1 = "C:\xlCFG.xls"

    End Sub



    Private Sub btnNumberOnlyRun_Click(sender As System.Object, e As System.EventArgs) Handles btnNumberOnlyRun.Click
        Dim iRc As Integer
        Dim rng As Excel.Range
        Dim iCnt As Integer


        '갯수
        rng = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbNumberOnlySrcRow.Text & ":" & tbNumberOnlySrcRow.Text)
        iCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountA(rng)


        If iCnt < 1 Then
            MsgBox("입력된 내용이 없습니다.")
            Exit Sub
        End If


        '길이(Byte)
        Dim startT, endT As DateTime

        startT = Now
        iRc = modNumberOnly(tbNumberOnlySrcRow.Text, tbNumberOnlyOutRow.Text, iCnt)
        endT = Now

        Globals.ThisAddIn.Application.ScreenUpdating = True

    End Sub

    Private Sub btnTrimRun_Click(sender As System.Object, e As System.EventArgs) Handles btnTrimRun.Click
        Dim iRc As Integer
        Dim rng As Excel.Range
        Dim iCnt As Integer


        '갯수
        rng = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbTrimSrcRow.Text & ":" & tbTrimSrcRow.Text)
        iCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountA(rng)


        If iCnt < 1 Then
            MsgBox("입력된 내용이 없습니다.")
            Exit Sub
        End If


        '머지
        Dim startT, endT As DateTime

        startT = Now
        iRc = modTrim(tbTrimSrcRow.Text, tbResultTrimRow.Text, iCnt)
        endT = Now

        Globals.ThisAddIn.Application.ScreenUpdating = True

        '탭 표시
        'TabControl2.SelectedTab = TabPage8
        'TabPage8.Select()
    End Sub

    Private Sub btnCelTypeRun_Click(sender As System.Object, e As System.EventArgs) Handles btnCelTypeRun.Click
        Dim rng As Excel.Range


        Globals.ThisAddIn.Application.ScreenUpdating = False
        rng = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbCelTypeSrcRow.Text & ":" & tbCelTypeSrcRow.Text)
        rng.Select()
        rng.NumberFormatLocal = "@"
        Globals.ThisAddIn.Application.ScreenUpdating = True

        MsgBox("텍스트형으로 변환되었습니다.")
    End Sub

    Private Sub btnByteLengthRun_Click(sender As System.Object, e As System.EventArgs) Handles btnByteLengthRun.Click
        Dim iRc As Integer
        Dim rng As Excel.Range
        Dim iCnt As Integer


        '갯수
        rng = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbByteLengthSrcRow.Text & ":" & tbByteLengthSrcRow.Text)
        iCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountA(rng)


        If iCnt < 1 Then
            MsgBox("입력된 내용이 없습니다.")
            Exit Sub
        End If


        '길이(Byte)
        Dim startT, endT As DateTime

        startT = Now
        iRc = modByteLength(tbByteLengthSrcRow.Text, tbByteLengthOutRow.Text, iCnt)
        endT = Now

        Globals.ThisAddIn.Application.ScreenUpdating = True
    End Sub


    Private Sub btnMergeRun2_Click(sender As System.Object, e As System.EventArgs) Handles btnMergeRun2.Click
        'Dim iRc As Integer
        Dim strWS1 As String
        Dim strNewSheet As String = "전송"
        Dim rng As Excel.Range
        Dim iCnt As Integer
        Dim iRc As Integer

        Globals.ThisAddIn.iProcessingNum = 0

        strWS1 = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.name()

        'rng = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbRcvPhoneRow.Text & "1").End(Excel.XlDirection.xlDown)

        '갯수
        'iCnt = rng.Row
        rng = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbReplace1.Text & ":" & tbReplace1.Text)
        'iCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountA(tbRcvPhoneRow.Text & ":" & tbRcvPhoneRow.Text)
        iCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountA(rng)



        If iCnt < 1 Then
            MsgBox("처리할 내용이 없습니다")
            Exit Sub
        Else

            Dim msg As String
            Dim title As String
            Dim style As MsgBoxStyle
            Dim response As MsgBoxResult

            msg = iCnt & "건 처리합니다"
            'style = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
            style = MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.YesNo
            title = "확인"   ' Define title.
            ' Display message.
            response = MsgBox(msg, style, title)
            If response = MsgBoxResult.Yes Then   ' User chose Yes.

            Else
                Exit Sub
            End If

        End If







        Dim iThreadCnt As Integer
        Dim m As Integer
        Dim params As New ParameterMerge()


        '쓰레드는 1000개 이상일때만 가동됨
        If iCnt >= 100 Then
            'iThreadCnt = Val(tbThreadNum.Text)
            iThreadCnt = Val(shareClass.gstrTHREADNUM)
        Else
            iThreadCnt = 1
        End If

        If iThreadCnt < 1 Then iThreadCnt = 1

        'params.ReArr(iThreadCnt)

        params.strResultRow = tbResultMergeRow.Text
        params.strTemplate = tbTemplate.Text
        params.strRep1 = tbReplace1.Text
        params.strRep2 = tbReplace2.Text
        params.strRep3 = tbReplace3.Text
        params.iThreadNum = iThreadCnt

        iModNumMergeThread = 0 ' 초기화
        iSum = 0 '초기화
        Globals.ThisAddIn.gsStatus = "진행"

        Dim WorkerThread(iThreadCnt) As System.Threading.Thread

        For m = 0 To iThreadCnt - 1
            'params.iMod(m) = m

            WorkerThread(m) = New System.Threading.Thread(AddressOf DoWorkerMerge2)
            WorkerThread(m).Name = "DoWorkerMerge" & m
            WorkerThread(m).Start(params) ' 새 스레드를 시작합니다.

            'Thread1.Join() ' 스레드 1이 끝날 때까지 기다립니다. 
            ' 반환 값을 표시합니다.    
            'MsgBox("스레드 1이 " & Tasks.RetVal & "값을 반환했습니다.")

        Next




        Globals.ThisAddIn.Application.ScreenUpdating = False

        '진행상태창(모달)
        '취소 또는 창 닫힘시 롤백처리
        '완료시 커밋처리
        dlgDB = New Dialog3()
        dlgDB.ShowDialog()

        Select Case dlgDB.DialogResult

            Case Windows.Forms.DialogResult.OK
                iRc = 1
            Case Windows.Forms.DialogResult.Cancel
                iRc = 0
            Case Windows.Forms.DialogResult.Abort
                iRc = -1
            Case Else
                iRc = -1
        End Select


        '행,열 자동조절
        Globals.ThisAddIn.Application.ActiveSheet.columns.AutoFit()
        Globals.ThisAddIn.Application.ActiveSheet.rows.AutoFit()

        Globals.ThisAddIn.Application.ScreenUpdating = True

        If iRc <= 0 Then
            MsgBox("취소 되었습니다", , "결과")
        Else
            MsgBox("완료 되었습니다", , "결과")
        End If






    End Sub



    Sub DoWorkerMerge2(ByVal param As Object)
        Dim iRc As Integer
        Dim strWS1 As String
        Dim rng As Excel.Range
        Dim iCnt As Integer



        Dim p As ParameterMerge
        p = CType(param, ParameterMerge)



        '스레드결과 배열크기 재정의
        Dim l As Integer
        'ReDim iThreadResult(Val(tbThreadNum.Text))
        ReDim iThreadResult(Val(shareClass.gstrTHREADNUM))
        ' 배열초기화
        For l = 0 To UBound(iThreadResult)
            iThreadResult(l) = 0   ' Initialize array.
        Next l



        strWS1 = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.name()

        '갯수
        'rng = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(p.strRep1.ToString & ":" & p.strRep1.ToString)
        rng = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbReplace1.Text & ":" & tbReplace1.Text)
        iCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountA(rng)


        'Dim startT, endT As DateTime

        'startT = Now

        Dim iMod As Integer
        '공유변수
        '스레드갯수
        SyncLock lockMergeObj
            iModNumMergeThread = iModNumMergeThread + 1
            iMod = iModNumMergeThread - 1  '지역변수
        End SyncLock

        'System.Threading.Interlocked.Increment(iModMergeThread) 'interlock

        'Me.Invoke(, New Object() {100})



        ' 
        'If dlgProg.InvokeRequired = True Then '해당 컨트롤이 이 쓰레드에 존재하는지 확인
        '    'True: 해당쓰레드가 아님, 콜백
        '    cbFunc = New changeProgress(AddressOf changeProgressMethod)
        '    'TextBox1 작동, 해당 컨트롤이 쓰레드에 존재하지 않으므로 새로운 컨트롤 생성.
        '    dlgProg.Invoke(cbFunc, New Object() {100})
        'Else
        '    '해당 쓰레드이므로 TextBox에 접근가능
        '    'TextBox1.Text = state.ToString()
        'End If






        'iRc = modMerge2(strWS1  , p.strResultRow, p.strTemplate, p.strRep1, p.strRep2, p.strRep3, iCnt  , p.iThreadNum, iModNumMergeThread - 1)
        '      modMerge2(strOpt1 , strOpt4       , strOpt5      , strOpt6  , strOpt7  , strOpt8  , iOpt9 , iOpt10      , iOpt11                ) 









        'Dim strPhone, strCallBack, 
        Dim strMsg As String
        Dim strRep1, strRep2, strRep3 As String
        Dim i As Integer
        Dim n As Integer



        'Dim trd = New Thread(AddressOf ThreadTask)
        'trd.IsBackground = True
        'trd.Start()


        'Dim uDlg = New UserControl2
        'uDlg.Show()

        'Globals.ThisAddIn.Application.ScreenUpdating = False
        'modMerge2 = 0


        iRc = 1

        For i = 1 To iCnt

            '' Lock인 상태일때
            'If Globals.ThisAddIn.iLock Then
            '    Globals.ThisAddIn.mre.WaitOne()
            'Else

            'End If




            '진행상자의 상태
            Select Case Globals.ThisAddIn.gsStatus

                Case "취소"
                    iRc = -1
                    Exit For

                Case "완료"
                    iRc = 1
                Case "진행"
                    iRc = 0
                Case "정지"
                    iRc = 0
                    '정지일 경우 1/100초씩 Sleep
                    While Globals.ThisAddIn.gsStatus = "정지"
                        Thread.Sleep(100)
                    End While

                    '정지상태에서 취소할경우 For루프를 벗어남
                    If Globals.ThisAddIn.gsStatus = "취소" Then
                        iRc = -1
                        Exit For
                    End If

                Case Else

            End Select






            'Globals.ThisAddIn.mre.WaitOne()

            'If i Mod p.iThreadNum = iMod Then
            '    SyncLock lockMergeProcessingNum
            '        Globals.ThisAddIn.iProcessingNum = Globals.ThisAddIn.iProcessingNum + 1
            '    End SyncLock
            'Else
            '    Continue For
            'End If





            If (i Mod p.iThreadNum = iMod) Then
                n = n + 1 ' 처리갯수 카운드

                Try
                    If dlgDB.InvokeRequired = True Then '해당 컨트롤이 이 쓰레드에 존재하는지 확인
                        'True: 해당쓰레드가 아님, 콜백
                        cbDBInsertFunc = New cbDBProgress(AddressOf cbDBInsertMethod)
                        'TextBox1 작동, 해당 컨트롤이 쓰레드에 존재하지 않으므로 새로운 컨트롤 생성.
                        dlgDB.Invoke(cbDBInsertFunc, New Object() {1, iCnt})
                    Else
                        '해당 쓰레드이므로 TextBox에 접근가능
                        'TextBox1.Text = state.ToString()
                    End If
                Catch ex As Exception '스레드가 처리도중 진행상자를 닫은경우 For루프를 벗어나서 롤백
                    iRc = -1
                    Exit For
                End Try

            Else
                Continue For ' 자기것이 아닌경우 For문으로
            End If














            'For i = iOpt10 To iOpt9 Step iOpt11

            'If itrdStatus Then Exit For


            'ipbValue = Int((i / iOpt9) * 100)  '백분율


            ''1000건마다 한번씩만 화면 Refresh
            ''1000건미만일때는 1건씩 Refresh 
            'If iCnt > 100 Then
            '    If i Mod 100 = 0 Then
            '        'Globals.ThisAddIn.Application.ScreenUpdating = True
            '        'System.Windows.Forms.Application.DoEvents()
            '        'Globals.ThisAddIn.Application.ScreenUpdating = False
            '        'Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt6 & i).select()
            '    Else
            '        'Globals.ThisAddIn.Application.ScreenUpdating = False
            '    End If
            'Else
            '    'Globals.ThisAddIn.Application.ScreenUpdating = True
            'End If





            '셀값
            Try
                'strPhone = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt3 & i).value
                'strCallBack = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt4 & i).value
                'strMsg = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt5 & i).value
                strMsg = p.strTemplate
            Catch ex As Exception
                iRc = -1
                MsgBox(ex.Message)
                'modMerge2 = 0
                Exit For
            End Try

            '치환값


            Dim builder As New StringBuilder(strMsg)
            'builder.Append(strMsg)




            If cbReplace1.Checked Then
                Try
                    strRep1 = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strWS1).Range(p.strRep1 & i).value

                    If strRep1 = Nothing Then
                        strRep1 = ""
                    End If
                    'strMsg = strMsg.Replace("@1", strRep1.ToString)
                Catch ex As Exception
                    strRep1 = ""
                End Try

                builder.Replace("@1", strRep1.ToString)
            Else
                builder.Replace("@1", "")
            End If


            If cbReplace2.Checked Then
                Try
                    strRep2 = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strWS1).Range(p.strRep2 & i).value
                    If strRep2 = Nothing Then
                        strRep2 = ""
                    End If

                    'strMsg = strMsg.Replace("@2", strRep2.ToString)
                Catch ex As Exception
                    strRep2 = ""
                End Try

                builder.Replace("@2", strRep2.ToString)
            Else
                builder.Replace("@2", "")
            End If


            If cbReplace3.Checked Then
                Try
                    strRep3 = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strWS1).Range(p.strRep3 & i).value
                    If strRep3 = Nothing Then
                        strRep3 = ""
                    End If

                    'strMsg = strMsg.Replace("@3", strRep3.ToString)
                Catch ex As Exception
                    strRep3 = ""
                End Try

                builder.Replace("@3", strRep3.ToString)
            Else
                builder.Replace("@3", "")
            End If



            Try
                'Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt2).range("A" & i).value = "'" & strPhone.ToString
                'Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt2).range("B" & i).value = "'" & strCallBack.ToString

                '치환
                'builder.Replace(vbLf, vbNewLine)
                'strMsg = strMsg.Replace(vbLf, vbNewLine)



                'If strRep1.Length > 0 Then
                '    Try
                '        strMsg = strMsg.Replace("@1", strRep1.ToString)
                '    Catch ex As Exception
                '        strMsg = strMsg.Replace("@1", "")
                '    End Try
                'End If


                'If strRep2.Length > 0 Then
                '    Try
                '        strMsg = strMsg.Replace("@2", strRep2.ToString)
                '    Catch ex As Exception
                '        strMsg = strMsg.Replace("@2", "")
                '    End Try
                'End If

                'If strRep3.Length > 0 Then
                '    Try
                '        strMsg = strMsg.Replace("@3", strRep3.ToString)
                '    Catch ex As Exception
                '        strMsg = strMsg.Replace("@3", "")
                '    End Try
                'End If


                ''줄바꿈 기호
                'Try
                '    strMsg = strMsg.Replace(vbLf, vbNewLine)
                'Catch ex As Exception

                'End Try

                'Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strWS1).Range(p.strResultRow & i).value = strMsg.ToString
                Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strWS1).Range(p.strResultRow & i).value = builder.ToString

            Catch ex As Exception
                iRc = -1
                MsgBox(i & "번째 행에 오류가 있습니다.")
                'modMerge2 = 0
                Exit For
            End Try

            'modMerge2 = 1


        Next

        'Try
        'Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt4 & i).select()
        'Catch
        'Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt4 & i - 1).select()
        'End Try

        'Globals.ThisAddIn.Application.ScreenUpdating = True

        'trd.Abort()

        'uDlg.Hide()
        'uDlg.Dispose()

        'MsgBox(i - 1 & "건/ 총" & iOpt9 & " 건 처리됨")

        'ipbValue = 0
        'itrdStatus = 0 '쓰레드 상태







        Dim iSum As Integer '스레드의 처리결과

        '스레드 실행결과
        ' > 0 : ?건 Insert
        ' 0 =< : Error
        iThreadResult(iMod) = iRc


        ' 이 스레드가 정상적으로 끝난 경우
        ' 다른 스레드가 끝나기를 기다린다.
        ' iSum : 스레드의 갯수만큼 순화하며 처리결과의 합을 구한다.
        '   정상처리결과:1, 비정상:-1
        If iRc > 0 Then
            iSum = 0
            While iSum < p.iThreadNum

                For k = 0 To p.iThreadNum
                    iSum = iSum + iThreadResult(k)

                    '어느 한 스레드에 문제가 있을경우
                    If iThreadResult(k) < 0 Then
                        iRc = -1
                        Exit While
                    End If

                    Thread.Sleep(10)
                Next

            End While

        End If




        '' 다른 스레드의 처리결과가 실패일경우 롤백처리함.
        '' DB Insert의 경우 All or Nothing 처리함.
        'Try
        '    If iRc <= 0 Then '롤백
        '        'MsgBox("취소 되었습니다.")

        '    Else
        '        'MsgBox("(" & iMod & ")" & n & "건 완료 되었습니다")

        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
























        'endT = Now

        'Globals.ThisAddIn.Application.ScreenUpdating = True
        'MsgBox(startT & " : " & endT)

    End Sub


    '트림 열선택
    Private Sub Label25_Click(sender As System.Object, e As System.EventArgs) Handles Label25.Click
        Try
            Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbTrimSrcRow.Text & ":" & tbTrimSrcRow.Text).select()
        Catch ex As Exception
        End Try
    End Sub

    '트림 결과출력열
    Private Sub Label24_Click(sender As System.Object, e As System.EventArgs) Handles Label24.Click
        Try
            Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbResultTrimRow.Text & ":" & tbResultTrimRow.Text).select()
        Catch ex As Exception
        End Try
    End Sub

    '트림
    Private Sub tbTrimSrcRow_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbTrimSrcRow.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then
            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub
        End If

        tbTrimSrcRow.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub

    '트림
    Private Sub tbResultTrimRow_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbResultTrimRow.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then
            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub
        End If

        tbResultTrimRow.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub


    '숫자추출
    Private Sub Label29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label29.Click
        Try
            Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbNumberOnlySrcRow.Text & ":" & tbNumberOnlySrcRow.Text).select()
        Catch ex As Exception
        End Try

    End Sub

    '숫자추출
    Private Sub Label28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label28.Click
        Try
            Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbNumberOnlyOutRow.Text & ":" & tbNumberOnlyOutRow.Text).select()
        Catch ex As Exception
        End Try

    End Sub

    '숫자추출
    Private Sub tbNumberOnlySrcRow_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbNumberOnlySrcRow.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then
            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub
        End If

        tbNumberOnlySrcRow.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub

    '숫자추출
    Private Sub tbNumberOnlyOutRow_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbNumberOnlyOutRow.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then
            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub
        End If

        tbNumberOnlyOutRow.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub

    '셀서식
    Private Sub Label33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label33.Click
        Try
            Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbCelTypeSrcRow.Text & ":" & tbCelTypeSrcRow.Text).select()
        Catch ex As Exception
        End Try
    End Sub

    '셀서식
    Private Sub tbCelTypeSrcRow_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbCelTypeSrcRow.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then
            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub
        End If

        tbCelTypeSrcRow.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub

    '문자길이
    Private Sub Label27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label27.Click
        Try
            Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbByteLengthSrcRow.Text & ":" & tbByteLengthSrcRow.Text).select()
        Catch ex As Exception
        End Try
    End Sub

    '문자길이
    Private Sub Label26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label26.Click
        Try
            Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbByteLengthOutRow.Text & ":" & tbByteLengthOutRow.Text).select()
        Catch ex As Exception
        End Try
    End Sub

    '문자길이
    Private Sub tbByteLengthSrcRow_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbByteLengthSrcRow.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then
            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub
        End If

        tbByteLengthSrcRow.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub

    '문자길이
    Private Sub tbByteLengthOutRow_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbByteLengthOutRow.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then
            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub
        End If

        tbByteLengthOutRow.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub



    '치환1 체크박스
    Private Sub cbReplace1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbReplace1.CheckedChanged
        tbReplace1.Enabled = cbReplace1.Checked

        Try
            If cbReplace1.Checked Then
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbReplace1.Text & ":" & tbReplace1.Text).select()
            End If

        Catch ex As Exception

        End Try

    End Sub

    '치환2 체크박스
    Private Sub cbReplace2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbReplace2.CheckedChanged
        tbReplace2.Enabled = cbReplace2.Checked

        Try
            If cbReplace2.Checked Then
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbReplace2.Text & ":" & tbReplace2.Text).select()
            End If

        Catch ex As Exception

        End Try
    End Sub

    '치환3 체크박스
    Private Sub cbReplace3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbReplace3.CheckedChanged
        tbReplace3.Enabled = cbReplace3.Checked

        Try
            If cbReplace3.Checked Then
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbReplace3.Text & ":" & tbReplace3.Text).select()
            End If

        Catch ex As Exception

        End Try
    End Sub

    '문자출력열
    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click
        Try
            Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(tbResultMergeRow.Text & ":" & tbResultMergeRow.Text).select()
        Catch ex As Exception
        End Try

    End Sub

    '문구변화
    Private Sub tbReplace1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbReplace1.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then

            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub

        End If

        tbReplace1.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub

    Private Sub tbReplace2_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbReplace2.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then

            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub

        End If

        tbReplace2.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub

    Private Sub tbReplace3_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tbReplace3.MouseClick
        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then

            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub

        End If

        tbReplace3.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub

    Private Sub tbResultMergeRow_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles tbResultMergeRow.MouseClick

        If Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1) > "Z" Then

            MsgBox("A-Z열까지만 선택가능합니다.")
            Exit Sub

        End If

        tbResultMergeRow.Text = Chr(Asc("A") + Globals.ThisAddIn.Application.ActiveCell.Column() - 1)
    End Sub


End Class
