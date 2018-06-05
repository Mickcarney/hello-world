Module Module1
    Private Structure Header
        <VBFixedString(2)> Public recordType As String
        <VBFixedString(21)> Public recordTitle As String
        <VBFixedString(11)> Public fileDate As String
        <VBFixedString(11)> Public applic As String
        <VBFixedString(11)> Public code As String
        <VBFixedString(50)> Public authority As String
    End Structure
    Private Structure Trailer
        <VBFixedString(2)> Public recordType As String
        <VBFixedString(9)> Public numRecs As String
        <VBFixedString(50)> Public totAmt As String
    End Structure
    Private Structure Detail
        <VBFixedString(2)> Public recordType As String
        <VBFixedString(21)> Public accRef As String
        <VBFixedString(71)> Public debtorName As String
        <VBFixedString(11)> Public pCode As String
        <VBFixedString(206)> Public addr As String
        <VBFixedString(21)> Public phoneNum As String
        <VBFixedString(2)> Public fill1 As String
        <VBFixedString(2)> Public fill2 As String
        <VBFixedString(2)> Public fill3 As String
        <VBFixedString(13)> Public liabFrom As String
        <VBFixedString(4)> Public fill4 As String
        <VBFixedString(13)> Public amt As String
        <VBFixedString(11)> Public yrFrom As String
        <VBFixedString(11)> Public yrTo As String
        <VBFixedString(11)> Public liabOrderDate As String
        <VBFixedString(11)> Public forPostCode As String
        <VBFixedString(206)> Public forAddr As String
        <VBFixedString(21)> Public caseRef As String
        <VBFixedString(120)> Public jointLiable As String
        'and the additional bits...
        <VBFixedString(128)> Public emailAddr As String
        <VBFixedString(40)> Public altPhone1 As String
        <VBFixedString(40)> Public altPhone2 As String
        <VBFixedString(40)> Public altPhone3 As String
        <VBFixedString(40)> Public altPhone4 As String
    End Structure
    Public strInputFile As String
    Public strOutputFile As String
    Private strErrorLog As String
    Private connHB As New IngConnection
    Private IngConnHB As New Ingres.Client.IngresConnection
    Private IngConnNDR As New Ingres.Client.IngresConnection
    <VBFixedString(128)> Private sEmail As String
    <VBFixedString(40)> Private sP1 As String
    <VBFixedString(40)> Private sP2 As String
    <VBFixedString(40)> Private sP3 As String
    <VBFixedString(40)> Private sP4 As String
    Private strErrMess As String



    Friend Sub ReadFileIn()
        Dim objReader As New System.IO.StreamReader(strInputFile)
        Dim newHeader As New Header
        Dim newTrailer As New Trailer
        Dim nd() As Detail
        Dim x As Integer = 0
        Dim strLine As String
        'ReDim Preserve nd(x)

        Try

        
        While Not objReader.EndOfStream
            strLine = objReader.ReadLine
            If Left(strLine, 1) = "H" Then
                newHeader.recordType = Left(strLine, 2)
                newHeader.recordTitle = Mid(strLine, 3, 21)
                newHeader.fileDate = Mid(strLine, 24, 11)
                newHeader.applic = Mid(strLine, 35, 11)
                newHeader.code = Mid(strLine, 46, 11)
                newHeader.authority = Mid(strLine, 57, 50)
            End If
            If Left(strLine, 1) = "D" Then
                ReDim Preserve nd(x)
                nd(x).recordType = Left(strLine, 2)
                nd(x).accRef = Mid(strLine, 3, 21)
                nd(x).debtorName = Mid(strLine, 24, 71)
                nd(x).pCode = Mid(strLine, 95, 11)
                nd(x).addr = Mid(strLine, 106, 206)
                nd(x).phoneNum = Mid(strLine, 312, 21)
                nd(x).fill1 = Mid(strLine, 333, 2)
                nd(x).fill2 = Mid(strLine, 335, 2)
                nd(x).fill3 = Mid(strLine, 337, 2)
                nd(x).liabFrom = Mid(strLine, 339, 13)
                nd(x).fill4 = Mid(strLine, 352, 4)
                nd(x).amt = Mid(strLine, 356, 13)
                nd(x).yrFrom = Mid(strLine, 369, 11)
                nd(x).yrTo = Mid(strLine, 380, 11)
                nd(x).liabOrderDate = Mid(strLine, 391, 11)
                nd(x).forPostCode = Mid(strLine, 402, 11)
                nd(x).forAddr = Mid(strLine, 413, 206)
                nd(x).caseRef = Mid(strLine, 619, 21)
                nd(x).jointLiable = Mid(strLine, 640, 120)
                If newHeader.applic = "CT         " Then
                    GetCTExtraBits(Left(nd(x).accRef, 7))
                ElseIf newHeader.applic = "NDR        " Then
                    GetNDRExtraBits(nd(x).accRef)
                End If


                nd(x).emailAddr = sEmail
                nd(x).altPhone1 = sP1
                nd(x).altPhone2 = sP2
                nd(x).altPhone3 = sP3
                nd(x).altPhone4 = sP4

                x = x + 1

            End If


                If Left(strLine, 1) = "T" Then
                    newTrailer.recordType = Left(strLine, 2)
                    newTrailer.numRecs = Mid(strLine, 3, 9)
                    newTrailer.totAmt = Mid(strLine, 12, 50)
                End If
        End While
        WriteOutputFile(newHeader.recordType + newHeader.recordTitle + newHeader.fileDate + newHeader.applic + newHeader.code + newHeader.authority)

        For i = 0 To UBound(nd)
            WriteOutputFile(nd(i).recordType & nd(i).accRef & nd(i).debtorName & _
                            nd(i).pCode & nd(i).addr & nd(i).phoneNum & _
                            nd(i).fill1 & nd(i).fill2 & nd(i).fill3 & _
                            nd(i).liabFrom & nd(i).fill4 & nd(i).amt & _
                            nd(i).yrFrom & nd(i).yrTo & nd(i).liabOrderDate & _
                            nd(i).forPostCode & nd(i).forAddr & nd(i).caseRef & _
                            nd(i).jointLiable & nd(i).emailAddr & _
                            " " & nd(i).altPhone1 & " " & nd(i).altPhone2 & " " & nd(i).altPhone3 & _
                            " " & nd(i).altPhone4)
        Next
        WriteOutputFile(newTrailer.recordType + newTrailer.numRecs + newTrailer.totAmt)
        Catch ex As Exception
            Dim strTextToWrite As String

            strTextToWrite = "Sub 'ReadFileIn' returned error: " + vbCrLf + ex.ToString + vbCrLf
            WriteErrLog(strTextToWrite)
        End Try
    End Sub
    Friend Sub WriteOutputFile(strText As String)

        Try
            My.Computer.FileSystem.WriteAllText(strOutputFile, strText & vbCrLf, True)
        Catch ex As Exception
            Dim strTextToWrite As String

            strTextToWrite = "Sub 'WriteOutputFile' returned error: " + vbCrLf + ex.ToString + vbCrLf
            WriteErrLog(strTextToWrite)
        End Try
    End Sub
    Friend Sub FileNames()
        Dim finfo As System.IO.FileInfo
        Dim fname As String = Nothing
        Dim fpath As String = Nothing
        Dim strNew As String = Nothing


        finfo = My.Computer.FileSystem.GetFileInfo(strInputFile)
        fname = finfo.Name
        fpath = finfo.DirectoryName
        strNew = fname + "_EXTRA"
        strOutputFile = fpath + "\" + strNew
        strErrorLog = fpath + "\bail_file_err.log"

    End Sub
    Friend Sub OpenNDRConnection()
        Try
            IngConnNDR = connHB.OpenNRConn

        Catch ex As Exception
            Dim strTextToWrite As String

            strTextToWrite = "Sub 'OpenNDRConnection' returned error: " + vbCrLf + ex.ToString + vbCrLf
            WriteErrLog(strTextToWrite)
        End Try
    End Sub
    Friend Sub OpenConnection()
        Try
            IngConnHB = connHB.OpenHBConn



            'MsgBox(IngNetConn.State.ToString)

        Catch ex As Exception


            Dim strTextToWrite As String

            strTextToWrite = "Sub 'OpenConnection' returned error: " + vbCrLf + ex.ToString + vbCrLf
            WriteErrLog(strTextToWrite)
        End Try
    End Sub
    Friend Sub GetCTExtraBits(strRef As String)
        Dim Cmd As New Ingres.Client.IngresCommand
        Dim reader As Ingres.Client.IngresDataReader
        Cmd.Connection = IngConnHB
        Cmd.CommandType = Data.CommandType.Text
        sEmail = " "
        sP1 = " "
        sP2 = " "
        sP3 = " "
        sP4 = " "

        Try
            Cmd.CommandText = "select pad(cast(a.email_addr as varchar(120))), pad(cast(b.phonenum1 as varchar(32))), pad(cast(b.phonenum2 as varchar(32))), pad(cast(b.phonenum3 as varchar(32))), pad(cast(b.phonenum4 as varchar(32))) " & _
                "from syemail a, syphone b where a.reference_id = " & strRef & " and b.ref = " & strRef


            reader = Cmd.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    sEmail = "e-mail: " & reader.GetString(0)
                    sP1 = "Alt. Phone1: " & reader.GetString(1)
                    sP2 = "Alt. Phone2: " & reader.GetString(2)
                    sP3 = "Alt. Phone3: " & reader.GetString(3)
                    sP4 = "Alt. Phone4: " & reader.GetString(4)
                End While
            End If

        Catch ex As Exception
            Dim strTextToWrite As String

            strTextToWrite = "Sub 'GetCTExtraBits' returned error: " + vbCrLf + ex.ToString + vbCrLf
            WriteErrLog(strTextToWrite)

        Finally

            reader.Close()


        End Try
    End Sub
    Friend Sub GetNDRExtraBits(strRef As String)
        Dim Cmd As New Ingres.Client.IngresCommand
        Dim reader As Ingres.Client.IngresDataReader
        Cmd.Connection = IngConnNDR
        Cmd.CommandType = Data.CommandType.Text
        sEmail = " "
        sP1 = " "
        sP2 = " "
        sP3 = " "
        sP4 = " "

        Try
            Cmd.CommandText = "select pad(cast(a.email_addr as varchar(120))), pad(cast(b.phonenum1 as varchar(32))), pad(cast(b.phonenum2 as varchar(32))), pad(cast(b.phonenum3 as varchar(32))), pad(cast(b.phonenum4 as varchar(32))) " & _
                "from syemail a, syphone b where a.reference_id = " & Left(strRef, 7) & " and b.ref = '" & strRef & "'"


            reader = Cmd.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    sEmail = "e-mail: " & reader.GetString(0)
                    sP1 = "Alt. Phone1: " & reader.GetString(1)
                    sP2 = "Alt. Phone2: " & reader.GetString(2)
                    sP3 = "Alt. Phone3: " & reader.GetString(3)
                    sP4 = "Alt. Phone4: " & reader.GetString(4)
                End While
            End If

        Catch ex As Exception
            Dim strTextToWrite As String

            strTextToWrite = "Sub 'GetNDRExtraBits' returned error: " + vbCrLf + ex.ToString + vbCrLf
            WriteErrLog(strTextToWrite)

        Finally

            reader.Close()


        End Try
    End Sub
    Friend Sub CloseConnection()
        Try
            'IngNetConn = conn.OpenHBConn()
            'MsgBox(IngNetConn.State.ToString)
            If IngConnHB.State = ConnectionState.Open Then
                IngConnHB.Close()
            End If
            If IngConnNDR.State = ConnectionState.Open Then
                IngConnNDR.Close()
            End If
        Catch ex As Exception

            Dim strTextToWrite As String

            strTextToWrite = "Sub 'CloseConnection' returned error: " + vbCrLf + ex.ToString + vbCrLf
            WriteErrLog(strTextToWrite)
        Finally
            IngConnHB.Dispose()

        End Try




    End Sub
    Public Sub WriteErrLog(ByVal strText As String, Optional ByVal dt As String = "", Optional ByVal tm As String = "")
        Try
            My.Computer.FileSystem.WriteAllText(strErrorLog, "------------------------------------------------------------" & vbCrLf, True)

            My.Computer.FileSystem.WriteAllText(strErrorLog, Date.Today.ToLongDateString & " " & Now.ToShortTimeString & " " & strText & vbCrLf, True)
            My.Computer.FileSystem.WriteAllText(strErrorLog, "------------------------------------------------------------" & vbCrLf, True)
        Catch ex As Exception
            strErrMess = ex.ToString
            WriteErrLog("Sub 'WriteErrorLog' returned error " & strErrMess)
        End Try


    End Sub
End Module
