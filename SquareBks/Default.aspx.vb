Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Web.Services
Imports System.IO


Partial Public Class _Default
    Inherits System.Web.UI.Page
    Public Shared upload_dir As String = "c:\inetpub\wwwroot\squarebks\squarebks\ul\"
    Public Shared download_dir As String = "c:\inetpub\wwwroot\squarebks\squarebks\dl\"
    Public Shared sqlconn As String = "provider=SQLOLEDB;database=BBRE;server=compaq-web\SQLEXPRESS;user id=ChelseaChat;pwd=Chelsea12"
    Public Shared Function OpenExcel(file As String, ByRef sheet As String) As String
        Dim procids1 As New ArrayList
        Dim procids2 As New ArrayList
        Dim procarray As Array = Process.GetProcessesByName("EXCEL")
        Dim oxl As New Excel.Application

        oxl.Workbooks.Open(file, Notify:=False)
        oxl.DisplayAlerts = False
        Dim procarray2 As Array = Process.GetProcessesByName("EXCEL")
        OpenExcel = ""
        For Each pro As Process In procarray
            procids1.Add(pro.Id)
        Next
        For Each pro As Process In procarray2
            If Not procids1.Contains(pro.Id) Then OpenExcel = pro.Id
        Next
        sheet = "[" & oxl.ActiveSheet.Name & "$]"
    End Function
    <WebMethod()> _
    Public Shared Function ProcessQBFile(f As String) As String
        Dim ls As New ADODB.Connection
        Dim le As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim rt As New ADODB.Recordset
        Dim file = upload_dir & f
        Dim newitems = ""
        Dim item = ""
        Dim value = ""
        Dim sheet = ""
        Dim tr = "OpenExcel"
        Dim CPID = ""
        Try
            ProcessQBFile = ""
            CPID = OpenExcel(file, sheet)
            tr = "connecting"
            le.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & file & "';Extended Properties=Excel 8.0;")
            ls.Open(sqlconn)
            tr = "select"
            rt.Open("select * from " & sheet, le)
            Do While Not rt.EOF
                If Not IsDBNull(rt(3).Value) Then
                    value = Replace(Trim(rt(3).Value), "'", "")
                    rs.Open("select * from QB_itemlist where name='" & value & "'", ls)
                    item = "<option>" & value & "</option>"
                    If Not rs.EOF Then
                        If rs(1).Value <> 1 Then ProcessQBFile += item
                    Else
                        newitems += item
                    End If
                    rs.Close()
                End If
                rt.MoveNext()
            Loop
            rt.Close()
            le.Close()
            ls.Close()
            If ProcessQBFile = "" Then ProcessQBFile = "<option>No New Items</option>" Else If newitems <> "" Then ProcessQBFile += "|" & newitems

        Catch ex As Exception
            ProcessQBFile = "<option>Error on " & tr & ": " & ex.Message & "</option>"
        End Try
        Shell("taskkill /pid " & CPID & " /f")
    End Function
    <WebMethod()> _
    Public Shared Function ProcessSQFile(f As String) As String
        Dim ls As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim file = upload_dir & f
        Dim la As Array
        Dim name, pp As String
        Dim sr As New StreamReader(file)
        Dim unknownpps = ""
        Dim cellstyle = "<td style='border-bottom-width:thin;border-bottom-style:solid;'>"
        Dim QBIs = "<option></option>"

        ProcessSQFile = ""
        ls.Open(sqlconn)
        rs.Open("select '<option value=" & Chr(34) & "' + id + '" & Chr(34) & " >' + name + '</option>' as expr1 from QB_Itemlist order by expr1", ls)
        If Not rs.EOF Then QBIs += rs.ToString
        rs.Close()
        sr.ReadLine()
        Do While sr.Peek >= 0
            la = Split(sr.ReadLine, Chr(44))
            name = Replace(Replace(la(0), "'", ""), Chr(34), "")
            pp = Replace(Replace(la(1), "'", ""), Chr(34), "")
            rs.Open("select * from Square_Itemlist where name=N'" & name & "' and PricePoint=N'" & pp & "'", ls)
            If rs.EOF Then unknownpps += "<tr>" & cellstyle & name & "</td>" & cellstyle & pp & "</td>" & cellstyle & "<select>" & QBIs & "</select></td></tr>"
            rs.Close()
        Loop
        sr.Close()
        ls.Close()
        If unknownpps <> "" Then ProcessSQFile = "New Square Name/Price Point(s)<table id='SQTable'><tr><th>Name</th><th>Price Point</th><th>QB Assignment</th></tr>" & _
                    unknownpps & "<tr><td><input type='button' value='save' onclick=" & Chr(34) & "SaveSQItems('" & f & "');" & Chr(34) & " /></td></tr></table>"
    End Function
    <WebMethod()> _
    Public Shared Function AddQBItems(s As String) As String
        Dim ls As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sa = s.Split("|")
        AddQBItems = ""
        Try
            ls.Open(sqlconn)
            For Each item In sa
                item = Replace(Replace(item, "'", ""), Chr(34), "")
                If item <> "" Then ls.Execute("insert into QB_Itemlist (Name,cat) values ('" & item & "',0)")
            Next
            rs.Open("select '<option>' + name + '</option>' from QB_Itemlist order by name", ls)
            If Not rs.EOF Then AddQBItems += rs.GetString
            ls.Close()
        Catch ex As Exception
            AddQBItems = "error: " & ex.Message
        End Try
    End Function
    <WebMethod()> _
    Public Shared Function SubmitSQItems(args As String) As String
        Dim ls As New ADODB.Connection
        args = Left(args, args.Length - 4)
        Dim argsa = Split(args, "////")

        SubmitSQItems = ""
        Try
            ls.Open(sqlconn)
            For Each row In argsa
                If Not row.Contains("undefined") Then
                    Dim cells = Split(row, "|")
                    ls.Execute("insert into Square_Itemlist (name,pricepoint,qb_item) values (N'" & cells(0) & "',N'" & cells(1) & "',N'" & cells(2) & "')")
                End If
            Next
            ls.Close()
        Catch ex As Exception
            SubmitSQItems = ex.Message
        End Try
    End Function
    <WebMethod()> _
    Public Shared Function DoOutput(f As String) As String
        Dim ls As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim file = upload_dir & f
        Dim sr As New StreamReader(file)
        Dim la As Array
        Dim name, pp As String
        Dim cnt As Integer
        Dim sales As Decimal
        Dim rows = ""
        Dim ofile = download_dir & "OUT_" & f
        If Not Directory.Exists(download_dir) Then Directory.CreateDirectory(download_dir)
        Dim sw As New StreamWriter(ofile)
        sw.WriteLine("Name,Qty,Total Sales")

        Try
            DoOutput = ""
            ls.Open(sqlconn)
            ls.Execute("update QB_itemlist set qty=0,sales=0")
            sr.ReadLine()
            Do While sr.Peek >= 0
                la = Split(sr.ReadLine, Chr(44))
                name = Replace(Replace(la(0), "'", ""), Chr(34), "")
                pp = Replace(Replace(la(1), "'", ""), Chr(34), "")
                cnt = la(4)
                sales = CDec(Right(la(5), la(5).length - 1))
                rs.Open("select QB_item from Square_Itemlist where name=N'" & name & "' and PricePoint=N'" & pp & "'", ls)
                If Not rs.EOF Then If Not IsDBNull(rs(0).Value) Then ls.Execute("update QB_itemlist set qty = qty + " & cnt & ",sales = sales + " & sales & " where id = " & rs(0).Value)
                rs.Close()
            Loop
            sr.Close()
            rs.Open("select name,qty,sales from QB_Itemlist where qty>0 order by name", ls)
            Do While Not rs.EOF
                sw.WriteLine(rs(0).Value & Chr(44) & rs(1).Value & Chr(44) & rs(2).Value)
                rows += "<tr><td>" & rs(0).Value & "</td><td>" & rs(1).Value & "</td><td>$" & rs(2).Value & "</tr>"
                rs.MoveNext()
            Loop
            rs.Close()
            sw.Close()
            If rows <> "" Then DoOutput = "QB Item Counts for '" & f & "'<table id='OutputTable'><tr><th>Name</th><th>Qty</th><th>Total Sales</th></tr>" & rows & "</table>|dl/" & Path.GetFileName(ofile) Else DoOutput = "File shows nothing sold."
        Catch ex As Exception
            DoOutput = ex.Message
        End Try
    End Function
    <WebMethod> _
    Public Shared Function bf_proc(f As String) As String
        Dim ls As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim path = upload_dir & f
        Dim sheet = ""
        Dim CPID = ""
        Try
            CPID = OpenExcel(path, sheet)
            bf_proc = "<table><thead><tr><th>Date</th><th>Description</th><th>Withdrawls</th><th>Deposits</th><th>Balance</th></tr></thead><tbody>"
            ls.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';Extended Properties=Excel 8.0;")
            rs.Open("select * from " & sheet, ls)
            Do While Not rs.EOF

                'bf_proc += "<tr><td>" & rs(0).Value & "</td><td>" & rs(1).Value & "</td><td>" & rs(2).Value & "</td><td>" & rs(3).Value & "</td><td>" & rs(4).Value & "</td></tr>"
                rs.MoveNext()
            Loop
            bf_proc += "</tbody></table>"
        Catch ex As Exception
            bf_proc = ex.Message & ex.StackTrace
        End Try
        Shell("taskkill /f /pid " & CPID)
    End Function
    Sub AjaxFileUpload1_UploadComplete(sender As Object, e As AjaxControlToolkit.AjaxFileUploadEventArgs)
        ul_save(AjaxFileUpload1, e)
    End Sub
    Sub AjaxFileUpload2_UploadComplete(sender As Object, e As AjaxControlToolkit.AjaxFileUploadEventArgs)
        ul_save(AjaxFileUpload2, e)
    End Sub  
    Sub AjaxFileUpload3_UploadComplete(sender As Object, e As AjaxControlToolkit.AjaxFileUploadEventArgs) Handles AjaxFileUpload3.UploadComplete
        ul_save(AjaxFileUpload3, e)
    End Sub
    Protected Sub ul_save(afu As AjaxControlToolkit.AjaxFileUpload, e As AjaxControlToolkit.AjaxFileUploadEventArgs)
        If Not Directory.Exists(upload_dir) Then Directory.CreateDirectory(upload_dir)
        afu.SaveAs(upload_dir & e.FileName)
    End Sub
End Class