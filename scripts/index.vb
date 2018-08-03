Imports System.IO

Sub Main(ByVal report As String)
    'Nothing to see here.  Move along.
End Sub

Function index(ByVal dir As String)
    Dim logName
    Dim rtn

    rtn = "<H1>Index of " & Path.GetDirectoryName(dir) & "</H1>"
    logName = "index.vb"

    Try
        Dim strPath As String = System.IO.Path.GetFullPath("html") & Path.GetDirectoryName(dir)
        If (Directory.Exists(strPath)) Then
            Dim fs, fo, x

            Dim fileEntries As String() = Directory.GetFiles(strPath)
            Dim fileName As String
            rtn = rtn & "<TABLE border='1'><TR><TH>File Name</TH><TH>Last Modified</TH></TR>"
            For Each fileName In fileEntries
                hs.writelog(logName, fileName)
                rtn = rtn & "<TR><TD><A href='" & Path.GetDirectoryName(dir) & "/" & Path.GetFileName(fileName) & "'>" & Path.GetFileName(fileName) & "</A></TD><TD>" & File.GetLastWriteTime(fileName) & "</TD></TR>"
            Next fileName

            Return rtn & "</TABLE>"
        Else
            Return "Folder does not exist:" & strPath
        End If
    Catch ex As Exception
        Return ex.Message
        hs.WriteLog(logName, ex.Message)
    End Try


End Function