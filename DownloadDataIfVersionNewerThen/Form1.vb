Imports System.Net.Sockets
Imports System.Text
Imports System.IO
Imports System.Threading

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim UrlToDownload As String = "http://agoge.com.br/SQLEXPR_x86_ENU.exe"
        'Dim UrlToDownload As String = "http://download.microsoft.com/download/c/2/8/c28cc7df-c9d2-453b-9292-ae7d242dfeca/SQLEXPR_x86_ENU.exe"

        MsgBox("I Will download " & UrlToDownload & " if date > 2000-01-01")

        Dim j As Integer
        Dim er As Integer

        Dim SQLPath As String = IO.Path.Combine(Application.StartupPath, "SQLEXPR_x86_ENU.exe")

        Do
            Try
                DownloadData(UrlToDownload, SQLPath, New Date(2000, 1, 1), Nothing,
                    Sub(i As Integer)
                        If i = 25 Or i = 50 Or i = 75 Or i = 100 Then
                            Debug.Print(i)
                            Thread.Sleep(5 * 1000)
                        End If
                    End Sub)

                My.Computer.FileSystem.DeleteFile(SQLPath)
                j += 1
                Debug.Print("Status: " & er & "/" & j)
            Catch ex As Exception
                er += 1
                j += 1
                Debug.Print("Status: " & er & "/" & j)
            End Try
        Loop

        MsgBox("OK")
        End
    End Sub

    Friend Shared Function DateToW3CDate(D As Date) As String
        Return D.ToString("ddd, dd MMM yyyy HH:mm:ss", New System.Globalization.CultureInfo("en-US")) & " GMT"
    End Function

    Friend Shared Function DownloadData(URL As String, FilePath As String, Optional IfModifiedSince As Nullable(Of Date) = Nothing, Optional ByRef ResponseHeaders As List(Of String) = Nothing, Optional Progress As Action(Of Integer) = Nothing) As Nullable(Of Date)
        Dim Uri As New Uri(URL)

        Dim SB As New StringBuilder

        Dim AlreadyDownloadedBytes As Integer

        If My.Computer.FileSystem.FileExists(FilePath) Then

            AlreadyDownloadedBytes = My.Computer.FileSystem.GetFileInfo(FilePath).Length

            SB.AppendLine(String.Format("GET {0} HTTP/1.1", Uri.AbsolutePath))
            SB.AppendLine(String.Format("Host: {0}", Uri.Host))
            SB.AppendLine(String.Format("Range: bytes={0}-", AlreadyDownloadedBytes))

            SB.AppendLine()
            SB.AppendLine()
        Else
            SB.AppendLine(String.Format("GET {0} HTTP/1.1", Uri.AbsolutePath))
            SB.AppendLine(String.Format("Host: {0}", Uri.Host))

            If IfModifiedSince IsNot Nothing Then
                Dim StrDate As String = DateToW3CDate(IfModifiedSince)
                SB.AppendLine("If-Modified-Since: " & StrDate)
            End If

            SB.AppendLine()
            SB.AppendLine()
        End If



        
        Dim LastModified As Nullable(Of Date)

        Dim RequestString As String = SB.ToString
        Dim RequestBuff As Byte() = System.Text.Encoding.ASCII.GetBytes(RequestString)

        Dim TcpClient As New TcpClient
        TcpClient.Connect(Uri.Host, Uri.Port)

        Using NetworkStream As NetworkStream = TcpClient.GetStream
            NetworkStream.Write(RequestBuff, 0, RequestBuff.Length)

            'Dim Headers As New List(Of String)
            ResponseHeaders = New List(Of String)
            Dim BuffSize As Integer = 65535 '20 * 1024 * 1024  ' 64 * 1024
            Dim Buff(BuffSize - 1) As Byte
            Dim DownloadedBytes As Integer
            Dim BytesRead As Integer

            Dim DataInit As Byte()

            Do
                BytesRead = NetworkStream.Read(Buff, 0, BuffSize)

                Dim S As String = System.Text.Encoding.ASCII.GetString(Buff, 0, BytesRead)

                Dim PHeader As Integer = S.IndexOf(vbCrLf & vbCrLf)

                If PHeader = 0 Then
                    Dim ReceivedHeaders As String() = S.Split(vbCrLf)
                    ResponseHeaders.AddRange(ReceivedHeaders)
                Else
                    Dim HeaderPart As String = S.Substring(0, PHeader)
                    Dim ReceivedHeaders As String() = HeaderPart.Split(vbCrLf)

                    For Each RH In ReceivedHeaders
                        ResponseHeaders.Add(RH.Trim)
                    Next

                    Dim DataStart As Integer = PHeader + 4
                    Dim DataCount = BytesRead - DataStart
                    ReDim DataInit(DataCount - 1)

                    Buffer.BlockCopy(Buff, DataStart, DataInit, 0, DataCount)
                    Exit Do
                End If
            Loop

            Dim Headers As String = Join(ResponseHeaders.ToArray, vbCrLf)
            'MsgBox(Headers)

            'HTTP/1.1 304 Not Modified
            'HTTP/1.1 200 OK
            Dim ResponseLine As String = ResponseHeaders.Item(0)
            Dim ResponseParts As String() = ResponseLine.Split(" ")
            Dim ResponseCode As Integer = ResponseParts(1)

            Select Case ResponseCode
                Case 304
                    Return Nothing
                Case 200
                Case 206
                Case Else
                    Throw New Exception("The server respond with: " & ResponseLine)
            End Select

            Dim LastModifiedHeader As String = ResponseHeaders.SingleOrDefault(Function(x) x.StartsWith("Last-Modified"))

            If Not String.IsNullOrEmpty(LastModifiedHeader) Then
                Dim LastModifiedStr As String = LastModifiedHeader.Substring(LastModifiedHeader.IndexOf(":") + 1)
                LastModified = CDate(LastModifiedStr)
            End If


            Dim ContentLengthHeader As String = ResponseHeaders.Single(Function(x) x.StartsWith("Content-Length"))
            Dim ContentLength As Integer = ContentLengthHeader.Split(":")(1).Trim


            Dim P As Integer

            Dim T As New Thread(
                Sub()

                        Do Until DownloadedBytes = ContentLength
                            P = DownloadedBytes / ContentLength * 100

                            If Progress IsNot Nothing Then
                                Progress.Invoke(P)
                            End If

                            Thread.Sleep(100)
                        Loop

                    End Sub)

            T.IsBackground = True
            T.Start()


            Using FileStream As New FileStream(FilePath, FileMode.Append)
                FileStream.Write(DataInit, 0, DataInit.Length)
                DownloadedBytes += DataInit.Length

                Do Until DownloadedBytes = ContentLength
                    BytesRead = NetworkStream.Read(Buff, 0, BuffSize)
                    FileStream.Write(Buff, 0, BytesRead)

                    DownloadedBytes += BytesRead
                Loop
            End Using

        End Using

        TcpClient.Close()

        Return LastModified

    End Function

End Class
