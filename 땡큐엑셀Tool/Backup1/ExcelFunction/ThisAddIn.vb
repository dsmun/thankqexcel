Public Class ThisAddIn

    Public taskPaneView = New UserControl1

    Friend iProcessingNum As Integer
    Friend gsStatus As String


    Private Sub ThisAddIn_Startup() Handles Me.Startup

        'Dim iRc



        Dim myTaskPane = Me.CustomTaskPanes.Add(taskPaneView, "땡큐엑셀Tool Beta")

        With myTaskPane
            '.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom
            '.Width = 250
            .Visible = True
        End With



        'taskPaneView.init()



        'Globals.ThisAddIn.Application.ScreenUpdating = False


        'Dim strTestPhone As String = ""

        ''Dim sCFGFile = Globals.ThisAddIn.taskPaneView.tbCFG
        ''Dim sCFGFile = "C:\xlCFG.xls"

        Dim sCFGFile As String = ""
        '설정파일
        taskPaneView.getCFGFile(sCFGFile)


        Globals.ThisAddIn.Application.ScreenUpdating = False


        ''설정파일 오픈
        'iRc = modOpenFile(sCFGFile)
        'If iRc Then

        'Else '설정파일이 없을경우 종료
        '    myTaskPane.Dispose()
        '    'MsgBox("설정파일(C:\xlCFG.xls) 확인후 다시 실행하세요")
        '    Exit Sub
        'End If


        ''테스트수신번호
        'strTestPhone = modGetConfig(sCFGFile, "XLCFG", "TESTPHONE")

        ''테스트수신번호 Set
        ''taskPaneView.setTestPhone(strTestPhone)


        ''DB Type : MSSQL,ORACLE
        ''taskPaneView.shareClass.gstrDB_TYPE = modGetConfig(sCFGFile, "XLCFG", "DB_TYPE")
        'shareClass.gstrDB_TYPE = modGetConfig(sCFGFile, "XLCFG", "DB_TYPE")

        ''DB IP
        ''taskPaneView.shareClass.gstrDB_IP = modGetConfig(sCFGFile, "XLCFG", "DBSERVER_IP")
        'shareClass.gstrDB_IP = modGetConfig(sCFGFile, "XLCFG", "DBSERVER_IP")
        ''taskPaneView.shareClass.gstrDB_PORT = modGetConfig(sCFGFile, "XLCFG", "DBSERVER_PORT")
        'shareClass.gstrDB_PORT = modGetConfig(sCFGFile, "XLCFG", "DBSERVER_PORT")
        ''taskPaneView.shareClass.gstrDB_NAME = modGetConfig(sCFGFile, "XLCFG", "DB_NAME")
        'shareClass.gstrDB_NAME = modGetConfig(sCFGFile, "XLCFG", "DB_NAME")
        ''taskPaneView.shareClass.gstrDB_SMSTABLE = modGetConfig(sCFGFile, "XLCFG", "DB_SMSTABLE")
        'shareClass.gstrDB_SMSTABLE = modGetConfig(sCFGFile, "XLCFG", "DB_SMSTABLE")
        ''taskPaneView.shareClass.gstrDB_UID = modGetConfig(sCFGFile, "XLCFG", "DB_UID")
        'shareClass.gstrDB_UID = modGetConfig(sCFGFile, "XLCFG", "DB_UID")
        ''taskPaneView.shareClass.gstrDB_PWD = modGetConfig(sCFGFile, "XLCFG", "DB_PWD")
        'shareClass.gstrDB_PWD = modGetConfig(sCFGFile, "XLCFG", "DB_PWD")


        ''taskPaneView.shareClass.gstrDB_LMSTABLE = modGetConfig(sCFGFile, "XLCFG", "DB_LMSTABLE")
        'shareClass.gstrDB_LMSTABLE = modGetConfig(sCFGFile, "XLCFG", "DB_LMSTABLE")
        ''taskPaneView.shareClass.gstrDB_GRPTABLE = modGetConfig(sCFGFile, "XLCFG", "DB_GRPTABLE")
        'shareClass.gstrDB_GRPTABLE = modGetConfig(sCFGFile, "XLCFG", "DB_GRPTABLE")

        ''taskPaneView.strUSING_SDATE = modGetConfig(sCFGFile, "XLCFG", "USING_SDATE") '라이선스(시작)
        'shareClass.gstrUSING_SDATE = modGetConfig(sCFGFile, "XLCFG", "USING_SDATE") '라이선스(시작)
        ''taskPaneView.strUSING_EDATE = modGetConfig(sCFGFile, "XLCFG", "USING_EDATE") '라이선스(끝)
        'shareClass.gstrUSING_EDATE = modGetConfig(sCFGFile, "XLCFG", "USING_EDATE") '라이선스(끝)

        ''taskPaneView.strTHREADNUM = modGetConfig(sCFGFile, "XLCFG", "THREADNUM") '스레드갯수
        'shareClass.gstrTHREADNUM = modGetConfig(sCFGFile, "XLCFG", "THREADNUM") '스레드갯수

        ''사용자정의 필드값
        'shareClass.gstrTRAN_ETC1 = modGetConfig(sCFGFile, "XLCFG", "TRAN_ETC1")
        'shareClass.gstrTRAN_ETC2 = modGetConfig(sCFGFile, "XLCFG", "TRAN_ETC2")
        'shareClass.gstrTRAN_ETC3 = modGetConfig(sCFGFile, "XLCFG", "TRAN_ETC3")
        'shareClass.gstrTRAN_ETC4 = modGetConfig(sCFGFile, "XLCFG", "TRAN_ETC4")
        'shareClass.gstrTRAN_REFKEY = modGetConfig(sCFGFile, "XLCFG", "TRAN_REFKEY")
        'shareClass.gstrTRAN_ID = modGetConfig(sCFGFile, "XLCFG", "TRAN_ID")

        'shareClass.gstrDB_SMSLOGTABLE = modGetConfig(sCFGFile, "XLCFG", "DB_SMSLOGTABLE")
        'shareClass.gstrDB_LMSLOGTABLE = modGetConfig(sCFGFile, "XLCFG", "DB_LMSLOGTABLE")
        'shareClass.gstrDB_GRPLOGTABLE = modGetConfig(sCFGFile, "XLCFG", "DB_GRPLOGTABLE")
        ''Globals.ThisAddIn.Application.ScreenUpdating = False


        ''설정파일 닫기
        'modCloseFile(sCFGFile)

        Globals.ThisAddIn.Application.ScreenUpdating = True


        '워크북이 없을경우 생성
        Try
            If Globals.ThisAddIn.Application.Worksheets.Count < 1 Then
                Globals.ThisAddIn.Application.Workbooks.Add()
            End If
        Catch ex As Exception
            Globals.ThisAddIn.Application.Workbooks.Add()
        End Try



        myTaskPane.Visible = True


        'Try

        '    '라이선스기간 체크
        '    Select Case UCase(shareClass.gstrDB_TYPE)
        '        Case "ORACLE"
        '            'myTaskPane.Visible = True

        '        Case "MSSQL"
        '            'DB에서 오늘날짜 가져오기
        '            'Dim strToday = getMSSQLToday(shareClass.gstrDB_IP, shareClass.gstrDB_UID, shareClass.gstrDB_PWD, shareClass.gstrDB_NAME)

        '            Dim strToday = Today()

        '            '사용기간(체크)
        '            If (strToday.ToString >= shareClass.gstrUSING_SDATE) And (strToday.ToString <= shareClass.gstrUSING_EDATE) Then
        '                'myTaskPane.Visible = True
        '            Else
        '                'myTaskPane.Visible = False
        '                taskPaneView.setTabControl(False)
        '                MsgBox("Excel-Function 라이선스 기간이 만료되었습니다")
        '            End If

        '        Case Else
        '    End Select

        'Catch ex As Exception
        '    taskPaneView.setTabControl(False)
        '    MsgBox(ex.Message)
        'End Try


        taskPaneView.goHomePage1("http://thank-q.tistory.com")

    End Sub




    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
