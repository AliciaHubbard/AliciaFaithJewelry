<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/AFDatabase.asp" -->
<%
Dim rsItem__MMColParam
rsItem__MMColParam = "1"
If (Request.QueryString("itemNumber") <> "") Then 
  rsItem__MMColParam = Request.QueryString("itemNumber")
End If
%>
<%
Dim rsItem
Dim rsItem_cmd
Dim rsItem_numRows

Set rsItem_cmd = Server.CreateObject ("ADODB.Command")
rsItem_cmd.ActiveConnection = MM_AFDatabase_STRING
rsItem_cmd.CommandText = "SELECT * FROM Jewelry WHERE itemNumber = ?" 
rsItem_cmd.Prepared = true
rsItem_cmd.Parameters.Append rsItem_cmd.CreateParameter("param1", 5, 1, -1, rsItem__MMColParam) ' adDouble

Set rsItem = rsItem_cmd.Execute
rsItem_numRows = 0
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
<h2><%=(rsItem.Fields.Item("Title").Value)%></h2><br/>
<img src="<%=(rsItem.Fields.Item("lgImage").Value)%>"/><br/>
<p><%=(rsItem.Fields.Item("Description").Value)%></p><br/>

<!-- InstanceEndEditable -->
</div>
<footer>
&copy;2013 Alicia Faith Hubbard &nbsp; &nbsp; <a href="messages.asp" id="admin">Administrator</a>
</footer>
</body>
<!-- InstanceEnd --></html>
<%
rsItem.Close()
Set rsItem = Nothing
%>
