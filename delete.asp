<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--author Alicia Hubbard, April 2013-->
<!--#include file="Connections/AFDatabase.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
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
    MM_editCmd.CommandText = "DELETE FROM Jewelry WHERE itemNumber = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "records.asp"
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
Dim rs_delete__MMColParam
rs_delete__MMColParam = "1"
If (Request.QueryString("itemNumber") <> "") Then 
  rs_delete__MMColParam = Request.QueryString("itemNumber")
End If
%>
<%
Dim rs_delete
Dim rs_delete_cmd
Dim rs_delete_numRows

Set rs_delete_cmd = Server.CreateObject ("ADODB.Command")
rs_delete_cmd.ActiveConnection = MM_AFDatabase_STRING
rs_delete_cmd.CommandText = "SELECT * FROM Jewelry WHERE itemNumber = ?" 
rs_delete_cmd.Prepared = true
rs_delete_cmd.Parameters.Append rs_delete_cmd.CreateParameter("param1", 5, 1, -1, rs_delete__MMColParam) ' adDouble

Set rs_delete = rs_delete_cmd.Execute
rs_delete_numRows = 0
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
<h2>You are about to Delete:</h2>
<p>&nbsp;</p>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
<table border="1" cellpadding="5">
  <tr>
    <td>itemNumber</td>
    <td>Description</td>
    <td>Title</td>
    <td>smImage</td>
    <td>lgImage</td>
    <td>sold</td>
    <td>type</td>
  </tr>
  <tr>
    <td><%=(rs_delete.Fields.Item("itemNumber").Value)%></td>
    <td><%=(rs_delete.Fields.Item("Description").Value)%></td>
    <td><%=(rs_delete.Fields.Item("Title").Value)%></td>
    <td><%=(rs_delete.Fields.Item("smImage").Value)%></td>
    <td><%=(rs_delete.Fields.Item("lgImage").Value)%></td>
    <td><%=(rs_delete.Fields.Item("sold").Value)%></td>
    <td><%=(rs_delete.Fields.Item("type").Value)%></td>
  </tr>
</table>
<input name="Delete" type="submit" value="Delete">&nbsp;
<a href="records.asp"><input type="button" id="button" value="Cancel"></a>
<input type="hidden" name="MM_delete" value="form1">
<input type="hidden" name="MM_recordId" value="<%= rs_delete.Fields.Item("itemNumber").Value %>">
</form>

<!-- InstanceEndEditable -->
</div>
<footer>
&copy;2013 Alicia Faith Hubbard &nbsp; &nbsp; <a href="messages.asp" id="admin">Administrator</a>
</footer>
</body>
<!-- InstanceEnd --></html>
<%
rs_delete.Close()
Set rs_delete = Nothing
%>
