<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--author Alicia Hubbard, April 2013-->
<!--#include file="Connections/AFDatabase.asp" -->
<%
Dim rsJewelry
Dim rsJewelry_cmd
Dim rsJewelry_numRows

Set rsJewelry_cmd = Server.CreateObject ("ADODB.Command")
rsJewelry_cmd.ActiveConnection = MM_AFDatabase_STRING
rsJewelry_cmd.CommandText = "SELECT [Description], Title, smImage, lgImage FROM Jewelry" 
rsJewelry_cmd.Prepared = true

Set rsJewelry = rsJewelry_cmd.Execute
rsJewelry_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rsJewelry_numRows = rsJewelry_numRows + Repeat1__numRows
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
<style>

@-webkit-keyframes gallery {
  0%   { opacity: .5; }
  100% { opacity: 1; }
}

@-moz-keyframes gallery {
  0%   { opacity: .5; }
  100% { opacity: 1; }
}

@-o-keyframes gallery {
  0%   { opacity: .5; }
  100% { opacity: 1; }
}

@keyframes gallery
{
  0%   { opacity: .5; }
  100% { opacity: 1; }
}

@-webkit-keyframes description
{
0%   {height:0%;
	  visibility:hidden;}
50% {height:0%;}
100% {height:100%;}
}

@-moz-keyframes description
{
0%   {height:0%;
	  visibility:hidden;}
50% {height:0%;}
100% {height:80%;}
}

@-o-keyframes description
{
0%   {height:0%;
	  visibility:hidden;}
50% {height:0%;}
100% {height:80%;}
}

@keyframes description
{
0%   {height:0%;
	  visibility:hidden;}
50% {height:0%;}
100% {height:80%;}
}

#gallery a:focus .large{	
	-webkit-animation: gallery 1s linear;
	-moz-animation: gallery 1s linear;
	-o-animation: gallery 1s linear;
	animation:gallery 1s linear;
}

#gallery a:focus .description{
	-webkit-animation:description 2s ease-in-out;
	-moz-animation:description 2s ease-in-out;
	-o-animation:description 2s ease-in-out;
	animation:description 2s ease-in-out;
}

</style>
<link href="css/gallery.css" rel="stylesheet" type="text/css"/>
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
<h2>Gallery</h2>
<div id="gallery">
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsJewelry.EOF)) 
%>
  <a href="#" tabindex="1"><img src="images/<%=(rsJewelry.Fields.Item("smImage").Value)%>" alt="<%=(rsJewelry.Fields.Item("title").Value)%>"></img>
    <div class="large"><img src="images/<%=(rsJewelry.Fields.Item("lgImage").Value)%>"><div class="description"><h2><%=(rsJewelry.Fields.Item("title").Value)%></h2><p><%=(rsJewelry.Fields.Item("description").Value)%></p></div></div></a>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsJewelry.MoveNext()
Wend
%>
</div>
<br/>
<!-- InstanceEndEditable -->
</div>
<footer>
&copy;2013 Alicia Faith Hubbard &nbsp; &nbsp; <a href="messages.asp" id="admin">Administrator</a>
</footer>
</body>
<!-- InstanceEnd --></html>
<%
rsJewelry.Close()
Set rsJewelry = Nothing
%>