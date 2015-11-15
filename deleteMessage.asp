<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/AFDatabase.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_AFDatabase_STRING
    MM_editCmd.CommandText = "DELETE FROM messages WHERE ID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "messages.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If

End If
%>
<%
Dim rs_deleteMessage__MMColParam
rs_deleteMessage__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rs_deleteMessage__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rs_deleteMessage
Dim rs_deleteMessage_cmd
Dim rs_deleteMessage_numRows

Set rs_deleteMessage_cmd = Server.CreateObject ("ADODB.Command")
rs_deleteMessage_cmd.ActiveConnection = MM_AFDatabase_STRING
rs_deleteMessage_cmd.CommandText = "SELECT * FROM messages WHERE ID = ?" 
rs_deleteMessage_cmd.Prepared = true
rs_deleteMessage_cmd.Parameters.Append rs_deleteMessage_cmd.CreateParameter("param1", 5, 1, -1, rs_deleteMessage__MMColParam) ' adDouble

Set rs_deleteMessage = rs_deleteMessage_cmd.Execute
rs_deleteMessage_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 1
Repeat1__index = 0
rs_deleteMessage_numRows = rs_deleteMessage_numRows + Repeat1__numRows
%>
<!DOCTYPE HTML>
<html><!-- InstanceBegin template="/Templates/AliciaFaithTemplate.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width; initial-scale=1.0">
<link rel="icon" type="image/png" href="images/favicon.jpg">
<script src="js/modernizr.js"></script>
<script src="//ajax.googleapis.com/ajax/libs/jquery/jquery.min.js"></script>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Alicia Faith Jewelry</title>
<!-- InstanceEndEditable -->
<link href="css/main.css" rel="stylesheet" type="text/css"/>
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
</head>

<body>
<header>
<h1>Alicia Faith Jewelry</h1>
<nav>
	<a href="default.htm">About</a>
    <a href="gallery.asp">Gallery</a>
    <a href="calendar.htm">Calendar</a>
    <a href="contact.asp">Contact</a>
</nav>
</header>
<div id="headerSpacing"></div>
<div id="content">
<!-- InstanceBeginEditable name="content" -->
<h2>You are about to permanantly delete:</h2>
<form METHOD="POST" name="form1" action="<%=MM_editAction%>">
  <table border="1">
    <tr>
      <td>Sender</td>
      <td>Subject</td>
    </tr>
    <tr>
      <td><%=(rs_deleteMessage.Fields.Item("sender").Value)%></td>
      <td><%=(rs_deleteMessage.Fields.Item("subject").Value)%></td>
    </tr>
  </table>
  <p>
    <input type="submit" name="button" id="button" value="Delete">
  </p>
  <input type="hidden" name="MM_delete" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_deleteMessage.Fields.Item("ID").Value %>">
</form>
<p>&nbsp;</p>
<!-- InstanceEndEditable -->
</div>
<footer>
&copy;2013 Alicia Faith Hubbard &nbsp; &nbsp; <a href="messages.asp" id="admin">Administrator</a>
</footer>
</body>
<!-- InstanceEnd --></html>
<%
rs_deleteMessage.Close()
Set rs_deleteMessage = Nothing
%>
