<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--author Alicia Hubbard, April 2013-->
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "default.htm"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
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
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_AFDatabase_STRING
    MM_editCmd.CommandText = "UPDATE Jewelry SET itemNumber = ?, [Description] = ?, Title = ?, smImage = ?, lgImage = ?, sold = ?, type = ? WHERE itemNumber = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("itemNumber"), Request.Form("itemNumber"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 203, 1, 1073741823, Request.Form("Description")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 255, Request.Form("Title")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 255, Request.Form("smImage")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 255, Request.Form("lgImage")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("sold"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 202, 1, 255, Request.Form("type")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
Dim rs_edit__MMColParam
rs_edit__MMColParam = "1"
If (Request.QueryString("itemNumber") <> "") Then 
  rs_edit__MMColParam = Request.QueryString("itemNumber")
End If
%>
<%
Dim rs_edit
Dim rs_edit_cmd
Dim rs_edit_numRows

Set rs_edit_cmd = Server.CreateObject ("ADODB.Command")
rs_edit_cmd.ActiveConnection = MM_AFDatabase_STRING
rs_edit_cmd.CommandText = "SELECT * FROM Jewelry WHERE itemNumber = ?" 
rs_edit_cmd.Prepared = true
rs_edit_cmd.Parameters.Append rs_edit_cmd.CreateParameter("param1", 5, 1, -1, rs_edit__MMColParam) ' adDouble

Set rs_edit = rs_edit_cmd.Execute
rs_edit_numRows = 0
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
<h2>Edit Item</h2>
<span> <a href="messages.asp">Messages</a> &nbsp; &nbsp;  <a href="records.asp">Items</a> &nbsp; &nbsp; <a href="add.asp">Add an Item</a>&nbsp; &nbsp; <a href="<%= MM_Logout %>">Log out</a></span>
<p>&nbsp;</p>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap align="right">ItemNumber:</td>
      <td>
      <input type="text" name="itemNumber" value="<%=(rs_edit.Fields.Item("itemNumber").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Description:</td>
      <td><textarea name="Description" cols="45" rows="5"><%=(rs_edit.Fields.Item("Description").Value)%></textarea></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Title:</td>
      <td><input type="text" name="Title" value="<%=(rs_edit.Fields.Item("Title").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">SmImage:</td>
      <td><input type="text" name="smImage" value="<%=(rs_edit.Fields.Item("smImage").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">LgImage:</td>
      <td><input type="text" name="lgImage" value="<%=(rs_edit.Fields.Item("lgImage").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Sold:</td>
      <td><input type="checkbox" name="sold" value=1 ></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Type:</td>
      <td><input type="text" name="type" value="<%=(rs_edit.Fields.Item("type").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">&nbsp;</td>
      <td><input type="submit" value="Update record">
        <a href="records.asp"><input type="button" id="button" value="Cancel"></a></td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= rs_edit.Fields.Item("itemNumber").Value %>">
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
rs_edit.Close()
Set rs_edit = Nothing
%>
