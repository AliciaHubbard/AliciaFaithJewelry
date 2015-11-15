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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_AFDatabase_STRING
    MM_editCmd.CommandText = "INSERT INTO messages (sender, [email address], subject, message) VALUES (?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 255, Request.Form("sender")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 255, Request.Form("email_address")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 255, Request.Form("subject")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 255, Request.Form("message")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "thankYou.asp"
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
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_AFDatabase_STRING
Recordset1_cmd.CommandText = "SELECT * FROM messages" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
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
<h2>Contact Me</h2>
<p>&nbsp;</p>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap align="right">Name:</td>
      <td><input type="text" name="sender" value="" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Email Address:</td>
      <td><input type="text" name="email_address" value="" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Subject:</td>
      <td><input type="text" name="subject" value="" size="32"></td>
    </tr>
    <tr>
      <td nowrap align="right" valign="top">Message:</td>
      <td valign="baseline"><textarea name="message" cols="50" rows="5"></textarea></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">&nbsp;</td>
      <td><input type="submit" value="Insert record"></td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
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
Recordset1.Close()
Set Recordset1 = Nothing
%>
