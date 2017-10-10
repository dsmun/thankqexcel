Imports System.Data.SqlClient 'MS-SQL

Module Module1


    'DB에서 오늘날짜 가져오기
    Function getMSSQLToday(ByVal IP As String, ByVal UID As String, ByVal PWD As String, ByVal DB As String)

        Dim constr As String = "server=" & IP & ";uid=" & UID & ";pwd=" & PWD & ";database=" & DB
        Dim strDate As String = ""

        Dim conn = New SqlConnection(constr)
        conn.Open()


        Dim cmd = conn.CreateCommand()
        cmd.CommandType = Data.CommandType.Text

        Dim sql As String

        'SELECT convert(char(10),getdate(),126)
        sql = "SELECT convert(varchar(10),getdate(),126)"
        cmd.CommandText = sql


        Try
            Dim dataReader As SqlDataReader = cmd.ExecuteReader()
            Dim iFieldCnt = dataReader.FieldCount()

            Do While dataReader.Read()
                For j = 0 To iFieldCnt - 1
                    strDate = dataReader(j)
                Next
            Loop

            dataReader.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return strDate

    End Function


    Function getORACLEToday(ByVal IP As String, ByVal UID As String, ByVal PWD As String, ByVal DB As String)

        Dim constr As String = "server=" & IP & ";uid=" & UID & ";pwd=" & PWD & ";database=" & DB
        Dim strDate As String = ""

        Dim conn = New SqlConnection(constr)
        conn.Open()


        Dim cmd = conn.CreateCommand()
        cmd.CommandType = Data.CommandType.Text

        Dim sql As String

        'SELECT convert(char(10),getdate(),126)
        sql = "SELECT convert(varchar(10),getdate(),126)"
        cmd.CommandText = sql


        Try
            Dim dataReader As SqlDataReader = cmd.ExecuteReader()
            Dim iFieldCnt = dataReader.FieldCount()

            Do While dataReader.Read()
                For j = 0 To iFieldCnt - 1
                    strDate = dataReader(j)
                Next
            Loop

            dataReader.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return strDate

    End Function


    Public Function modOpenFile(ByVal strOpt1 As String)

        'Dim wbDB As Excel.Workbook

        Globals.ThisAddIn.Application.ScreenUpdating = False

        Try
            Globals.ThisAddIn.Application.Workbooks.Open(Filename:=strOpt1)
            'wbDB = Globals.ThisAddIn.Application.Workbooks.Open(Filename:=strOpt1)
            modOpenFile = 1
        Catch ex As Exception
            MsgBox(ex.Message)
            modOpenFile = 0
        End Try

        Globals.ThisAddIn.Application.ScreenUpdating = True


    End Function



    Public Function modCloseFile(ByVal strOpt1 As String)

        'Dim wbDB As Excel.Workbook

        Globals.ThisAddIn.Application.ScreenUpdating = False

        'wbDB.Close(SaveChanges:=False)
        Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges:=False)

        Globals.ThisAddIn.Application.ScreenUpdating = True


        modCloseFile = 1

    End Function


    'Config 값리턴
    'strOpt1 : 엑셀파일
    'strOpt2 : 시트
    'strOpt3 : 필드
    Public Function modGetConfig(ByVal strOpt1 As String, ByVal strOpt2 As String, ByVal strOpt3 As String) As String

        Dim rng As Excel.Range
        Dim iCnt As Integer
        Dim wbDB As Excel.Workbook
        Dim ws As Excel.Worksheet

        Dim strRet As String = ""

        'Globals.ThisAddIn.Application.ScreenUpdating = False

        'wbDB = Globals.ThisAddIn.Application.Workbooks.Open(Filename:=strOpt1)
        wbDB = Globals.ThisAddIn.Application.ActiveWorkbook
        'rng = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(strOpt2).Range("A:A")
        ws = wbDB.Sheets(strOpt2)
        rng = ws.Range("A:A")



        Try

            Dim c As Excel.Range
            Dim firstAddress As String

            c = rng.Find(What:=strOpt3)
            'iCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountA(rng)

            If Not c Is Nothing Then
                firstAddress = c.Address
                Do
                    iCnt = iCnt + 1

                    If iCnt = 1 Then
                        strRet = c.Offset(, 1).Value
                    Else
                        strRet = strRet & ";" & c.Offset(, 1).Value
                    End If

                    c = rng.FindNext(c)
                Loop While Not c Is Nothing And c.Address <> firstAddress
            Else
                MsgBox(strOpt3 & " 에 대한 설정값이 없습니다")
            End If



        Catch ex As Exception


        End Try





        'wbDB.Close(SaveChanges:=False)


        'Globals.ThisAddIn.Application.ScreenUpdating = True

        'MsgBox(iCnt & ":" & strRet)


        modGetConfig = strRet

    End Function



    Function ExtractNumber(ByVal InputString As String)

        Dim i As Integer = 0

        Dim Num As String = ""



        For i = Len(InputString) To 1 Step -1

            If IsNumeric(Mid(InputString, i, 1)) Then

                Num = Mid(InputString, i, 1) & Num

            End If

        Next i

        ExtractNumber = Num

    End Function















    Public Function modByteLength(ByVal strOpt1 As String, ByVal strOpt2 As String, ByVal iOpt3 As Integer) As Integer

        Dim iLen As Integer
        Dim i As Integer
        Dim strMsg As String


        modByteLength = 0

        For i = 1 To iOpt3

            '100건마다 한번씩만 화면 Refresh
            '100건미만일때는 1건씩 Refresh 
            If iOpt3 > 100 Then
                If i Mod 100 = 0 Then
                    Globals.ThisAddIn.Application.ScreenUpdating = True
                Else
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                End If
            Else
                Globals.ThisAddIn.Application.ScreenUpdating = True
            End If


            '셀값
            Try
                strMsg = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(strOpt1 & i).value

            Catch ex As Exception
                MsgBox(i & "번째 행에 오류가 있습니다.")
                modByteLength = 0

                Exit For

            End Try



            Try

                Try
                    iLen = System.Text.Encoding.Default.GetByteCount(strMsg)
                Catch ex As Exception
                    MsgBox(i & "번째 행에 오류가 있습니다.")
                End Try


                'Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt4 & i).value = strMsg.ToString
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(strOpt2 & i).value = iLen

            Catch ex As Exception
                MsgBox(i & "번째 행에 오류가 있습니다.")
                modByteLength = 0

                Exit For
            End Try

            modByteLength = 1

        Next


        MsgBox(i - 1 & "건/ 총" & iOpt3 & " 건 처리됨")

    End Function












    '트림
    Public Function modTrim(ByVal strOpt1 As String, ByVal strOpt2 As String, ByVal iOpt3 As Integer) As Integer

        Dim strMsg As String
        Dim strTrimmedMsg As String
        Dim i As Integer


        modTrim = 0
        strTrimmedMsg = ""

        For i = 1 To iOpt3

            '1000건마다 한번씩만 화면 Refresh
            '1000건미만일때는 1건씩 Refresh 
            If iOpt3 > 1000 Then
                If i Mod 1000 = 0 Then
                    Globals.ThisAddIn.Application.ScreenUpdating = True
                Else
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                End If
            Else
                Globals.ThisAddIn.Application.ScreenUpdating = True
            End If


            '셀값
            Try
                'strMsg = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt1 & i).value
                strMsg = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(strOpt1 & i).value

            Catch ex As Exception
                MsgBox(i & "번째 행에 오류가 있습니다.")
                modTrim = 0

                Exit For

            End Try



            Try

                Try
                    strTrimmedMsg = strMsg.Trim()
                Catch ex As Exception
                    MsgBox(i & "번째 행에 오류가 있습니다.")
                End Try


                'Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt4 & i).value = strMsg.ToString
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(strOpt2 & i).value = strTrimmedMsg.ToString

            Catch ex As Exception
                MsgBox(i & "번째 행에 오류가 있습니다.")
                modTrim = 0

                Exit For
            End Try

            modTrim = 1

        Next


        MsgBox(i - 1 & "건/ 총" & iOpt3 & " 건 처리됨")

    End Function





    '숫자만 추출
    Public Function modNumberOnly(ByVal strOpt1 As String, ByVal strOpt2 As String, ByVal iOpt3 As Integer) As Integer

        Dim i As Integer
        Dim strMsg As String
        Dim strExtNumber As String = ""


        modNumberOnly = 0

        For i = 1 To iOpt3

            '100건마다 한번씩만 화면 Refresh
            '100건미만일때는 1건씩 Refresh 
            If iOpt3 > 100 Then
                If i Mod 100 = 0 Then
                    Globals.ThisAddIn.Application.ScreenUpdating = True
                Else
                    Globals.ThisAddIn.Application.ScreenUpdating = False
                End If
            Else
                Globals.ThisAddIn.Application.ScreenUpdating = True
            End If


            '셀값
            Try
                strMsg = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(strOpt1 & i).value

            Catch ex As Exception
                MsgBox(i & "번째 행에 오류가 있습니다.")
                modNumberOnly = 0

                Exit For

            End Try



            Try

                Try
                    strExtNumber = ExtractNumber(strMsg)
                Catch ex As Exception
                    MsgBox(i & "번째 행에 오류가 있습니다.")
                End Try


                'Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(strOpt1).Range(strOpt4 & i).value = strMsg.ToString
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range(strOpt2 & i).value = strExtNumber

            Catch ex As Exception
                MsgBox(i & "번째 행에 오류가 있습니다.")
                modNumberOnly = 0

                Exit For
            End Try

            modNumberOnly = 1

        Next


        MsgBox(i - 1 & "건/ 총" & iOpt3 & " 건 처리됨")

    End Function

End Module
