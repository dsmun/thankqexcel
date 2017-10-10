Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.IO

Public Class UserControl1

    '서버소켓
    Dim Server As TcpListener 'TCP네트워크 클라이언트에서 연결 수신
    Dim SerClient As TcpClient  'TCP네트워크 서비스에 대한 클라이언트 연결 제공
    Dim myStream As NetworkStream  '네트워크 스트림
    Dim myWrite As StreamWriter  '스트림 쓰기
    Dim Start As Boolean = False  '서버 시작
    Dim myServer As Thread  '스레드


    '클리이언트소켓
    Dim client As TcpClient
    Dim myClientStream As NetworkStream
    Dim myRead As StreamReader
    Dim myReader As Thread



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


    '웹브라우저1
    Public Sub goHomePage1(ByVal strOpt1 As String)

        '설정파일은 일단 C:\에 위치
        Me.WebBrowser1.Navigate(strOpt1)

    End Sub



    Private Sub Button서버소켓시작_Click(sender As System.Object, e As System.EventArgs) Handles Button서버소켓시작.Click

        Start = True

        Dim addr = New IPAddress(0)
        Server = New TcpListener(addr, 9969)
        Server.Start()

        myServer = New Thread(AddressOf ServerStart)
        myServer.Start()

    End Sub


    Private Sub ServerStart()

        MessageView("서버가 실행이 되었습니다.")

        While (Start)
            Try
                SerClient = Server.AcceptTcpClient()
                MessageView("클라이언트가 접속하였습니다.")

                myStream = SerClient.GetStream
                myWrite = New StreamWriter(myStream)

            Catch ex As Exception
                MessageView("클라이언트가 접속을 종료하였습니다.")
                Return
            End Try
        End While

    End Sub


    Private Sub MessageView(ByVal strText As String)

        Me.TextBox메시지.AppendText(strText + Chr(13) + Chr(10))
        Me.TextBox메시지.Focus()
        Me.TextBox메시지.ScrollToCaret()

    End Sub

    Private Sub Button보내기_Click(sender As System.Object, e As System.EventArgs) Handles Button보내기.Click
        Dim msg As String


        'msg = Format(Now(), "yyyy-mm-dd hh:mm:ss")

        msg = Me.TextBox메시지.Text

        If Not myStream Is Nothing And Not myWrite Is Nothing Then

            'myWrite.WriteLine("서버에서 메시지를 보냈습니다.")
            myWrite.WriteLine(msg)
            myWrite.Flush()
        End If
    End Sub



    Private Sub Button범위보내기_Click(sender As System.Object, e As System.EventArgs) Handles Button범위보내기.Click
        Dim msg As String
        Dim adRC As String

        'Globals.ThisAddIn.Application.Selection.row()
        'Globals.ThisAddIn.Application.Selection.column()
        adRC = Globals.ThisAddIn.Application.Selection.row() & "," & Globals.ThisAddIn.Application.Selection.column()

        'adRC = "10,10"
        'msg = adRC & ";" & Me.TextBox메시지.Text
        msg = adRC & ";" & Globals.ThisAddIn.Application.Selection.formula

        If Not myStream Is Nothing And Not myWrite Is Nothing Then

            'myWrite.WriteLine("서버에서 메시지를 보냈습니다.")
            myWrite.WriteLine(msg)
            myWrite.Flush()
        End If

    End Sub



    Private Sub Button서버소켓종료_Click(sender As System.Object, e As System.EventArgs) Handles Button서버소켓종료.Click



        If Not myStream Is Nothing Then
            Try
                'myWrite.WriteLine("서버를 종료합니다")
                myWrite.Flush()
            Catch ex As Exception
            End Try
        End If

        Me.Start = False

        If Not myStream Is Nothing Then
            myWrite.Close()
        End If

        If Not myStream Is Nothing Then
            myStream.Close()
        End If

        If Not SerClient Is Nothing Then
            SerClient.Close()
        End If

        If Not Server Is Nothing Then
            Server.Stop()
        End If

        If Not myServer Is Nothing Then
            myServer.Abort()
        End If

        MessageView("서버를 종료하였습니다")

    End Sub






    '클라이언트단
    Private Sub Button서버접속_Click(sender As System.Object, e As System.EventArgs) Handles Button서버접속.Click
        Try
            client = New TcpClient("127.0.0.1", 9969)
            messageViewClient("서버에 접속 했습니다.")
            myClientStream = client.GetStream
            myRead = New StreamReader(myClientStream)
            myReader = New Thread(AddressOf Receive)
            myReader.Start()

        Catch ex As Exception
            messageViewClient("서버에 접속하지 못 했습니다.")
        End Try
    End Sub


    Private Sub MessageViewClient(ByVal strText As String)
        Me.TextBoxClient메시지.AppendText(strText + Chr(13) + Chr(10))
        Me.TextBoxClient메시지.Focus()
        Me.TextBoxClient메시지.ScrollToCaret()
    End Sub



    Private Sub WriteCell(ByVal strText As String)
        Dim iPos As Integer
        Dim strAddr As String
        Dim strMsg As String
        Dim iRow As Integer
        Dim iCol As Integer
        Dim jPos As Integer

        '주소와 수식 구분자
        iPos = InStr(strText, ";")
        strAddr = Mid(strText, 1, iPos - 1)
        strMsg = Mid(strText, iPos + 1, strText.Length - iPos)

        'strMsg = strText.Substring(iPos + 1, strText.Length - (iPos + 1))
        'strAddr = strText.Substring(0, iPos)

        '행,열 구분자
        jPos = InStr(strAddr, ",")
        iRow = Int(Mid(strAddr, 1, jPos - 1))
        iCol = Int(Mid(strAddr, jPos + 1, Len(strAddr) - jPos))

        'iRow = Int(strAddr.Substring(0, jPos))
        'iCol = Int(strAddr.Substring(jPos + 1, strAddr.Length - (jPos + 1)))

        'MsgBox(strAddr)
        'MsgBox(strMsg)

        'MsgBox(iRow)
        'MsgBox(iCol)


        Globals.ThisAddIn.Application.Cells(iRow, iCol).Value = strMsg
        'Globals.ThisAddIn.Application.ActiveCell.Value = strText

    End Sub



    Private Sub Receive()
        Try
            While (True)
                If (myClientStream.CanRead) Then
                    Dim msg = myRead.ReadLine
                    If (msg.Length > 0) Then
                        MessageViewClient(msg)
                        WriteCell(msg)
                    End If
                End If
            End While
        Catch
            Return
        End Try

    End Sub




    Private Sub Button접속종료_Click(sender As System.Object, e As System.EventArgs) Handles Button접속종료.Click

        Try
            If Not myRead Is Nothing Then
                myRead.Close()
            End If

            If Not myClientStream Is Nothing Then
                myClientStream.Close()
            End If

            If Not SerClient Is Nothing Then
                SerClient.Close()
            End If

            If Not client Is Nothing Then
                client.Close()
            End If

            If Not myReader Is Nothing Then
                myReader.Abort()
            End If
        Catch
            Return
        End Try


        MessageViewClient("접속을 종료하였습니다.")

    End Sub


End Class
