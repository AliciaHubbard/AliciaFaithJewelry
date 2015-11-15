<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--author Alicia Hubbard, April 2013-->
<%
dim mail
Set mail=Server.CreateObject("CDO.Message")
mail.To = "lah2561@email.vccs.edu"
mail.From = Request.Form("email")
mail.Subject = Request.Form("subject")
mail.TextBody = Request.Form("message")
mail.Send()
Response.Write("Mail Sent!")
Set mail = nothing
%>
