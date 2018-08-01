<%@ Page Language="VB" %>

<script runat="server">

    Dim hs As Scheduler.hsapplication

    Sub Page_Load(Sender As Object, E As EventArgs)
        hs = Context.Items("Content")
    End Sub

    Private Function GetHeadContent(PageTitle As String) As String
        Dim header
        header = ""
        Try
            header = hs.GetPageHeader("", PageTitle, "", "", False, False, True, False, False)

        Catch ex As Exception
            hs.WriteLog(PageTitle & " - Get Head", ex.Message)
        End Try

        Return header
    End Function

    Private Function GetBodyContent(PageTitle As String) As String
        Try
            Return hs.GetPageHeader("", PageTitle, "", "", False, True, False, True, False)
        Catch ex As Exception
            hs.WriteLog(PageTitle & " - Get Body", ex.Message)
        End Try
        Return ""
    End Function

</script>
<html>
<head runat="server">
<%          Response.Write(GetHeadContent("Index"))%>

</head>
<body><%
          Response.Write(GetBodyContent("Index"))
          Response.Write(hs.Runscriptfunc("index.vb", "index", Request.Path, True, False))
%>
</body>
</html>