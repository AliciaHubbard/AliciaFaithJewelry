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
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 255, Request.Form("emailAddress")) ' adVarWChar
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
<link href="css/validation.css" rel="stylesheet" type="text/css"/>
<script type="text/javascript">
function validate(){
	var name= validateName();
	var email= validateEmail();
	var subject= validateSubject();
	var message= validateMessage();
	
	if(name && email && subject && message){
		return true;
	}
	else {return false;}
}

function validateName(){
	if(document.form1.sender.value == "" || document.form1.sender.value.length<3 || !document.form1.sender.value.match(/^[A-Za-z ]+$/) || document.form1.sender.value.length>50){
		document.form1.sender.className="invalid";
		document.getElementById("nameHint").style.display='inline';
		return false;
	}
	else {
		document.form1.sender.className="null";
		document.getElementById("nameHint").style.display='none';
		return true;
	}
}

function validateEmail(){
	if(document.form1.emailAddress.value == "" || !document.form1.emailAddress.value.match(/^\S+@\S+\.\S+$/) || document.form1.emailAddress.value.length>50){
		document.form1.emailAddress.className="invalid";
		document.getElementById("emailHint").style.display='inline';
		return false;
	}
	else {
		document.form1.emailAddress.className="null";
		document.getElementById("emailHint").style.display='none';
		return true;
	}
}

function validateSubject(){
	if(document.form1.subject.value == "" || document.form1.subject.value.length<3 || document.form1.subject.value.length>40){
		document.form1.subject.className="invalid";
		document.getElementById("subjectHint").style.display='inline';
		return false;
	}
	else {
		document.form1.subject.className="null";
		document.getElementById("subjectHint").style.display='none';
		return true;
	}
}

function validateMessage(){
	if(document.form1.message.value == "" || document.form1.message.value.length<3 || document.form1.message.value.length>500){
		document.form1.message.className="invalid";
		document.getElementById("messageHint").style.display='inline';
		return false;
	}
	else {
		document.form1.message.className="null";
		document.getElementById("messageHint").style.display='none';
		return true;
	}
}


</script>
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
<form method="POST" action="<%=MM_editAction%>" name="form1" onsubmit="return validate();">
  <table align="center">
  	<tr valign="baseline">
    	<td>&nbsp;</td>
    	<td><span id="nameHint">name must be at least 3 characters<br/>and contain only letters and spaces</span></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Name:</td>
      <td><input type="text" name="sender" value="" size="32" onfocus="document.form1.sender.className='null'" onBlur="validateName();"><br/><span id="emailHint">enter an email like: Alicia@mail.com</span></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Email Address:</td>
      <td><input type="text" name="emailAddress" value="" size="32" onfocus="document.form1.emailAddress.className='null'" onBlur="validateEmail();"><br/><span id="subjectHint">enter a subject beteween 3 and 40 characters</span></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Subject:</td>
      <td><input type="text" name="subject" value="" size="32" onfocus="document.form1.subject.className='null'" onBlur="validateSubject();"><br/><span id="messageHint">enter a message beteween 3 and 500 characters</span></td>
    </tr>
    <tr>
      <td nowrap align="right" valign="top">Message:</td>
      <td valign="baseline"><textarea name="message" value="" cols="25" rows="5" onfocus="document.form1.message.className='null'" onBlur="validateMessage();"></textarea></td>
    </tr>
  </table>
  <input type="submit" value="Send">
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
