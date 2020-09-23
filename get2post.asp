<%
'-- convert URL parameters to form fields
%>
<html>
<head><title>Get2Post by Klemens Schmid</title></head>
<body>
<h2>HTTP Get to Post Converter</h2>
This service converts your URL request into a form post. Thus it simulates a form submit
with input fields corresponding to the URL parameters. <br/>
You can pick up the source of this ASP page from here. My FormSniffer2 helps you building 
the URL from the actual Web page you want to replay.
<p>I promise to not evaluate your request in any way. I can't guarantee that somebody else
hacks my Web site and sniffs your data. If you have needs for better security publish this 
ASP page in your intranet or even on your own machine with IIS installed.
Please also notice this 
<a href="http://www.disclaimer.de/disclaimer.htm">disclaimer</a>
<p>For more services like this and a lot of developer tools enhancing Microsoft Outlook, WAP and
the Internet visit my <a href="http://www.schmidks.de">homepage</a>.
<br/><br/>

<form action="<%=Request.Querystring("g2p_action")%>" method="POST">
<%
For Each fld in Request.QueryString
	If fld <> "g2p_action" Then
%>
<input type="hidden" name="<%=fld%>" value="<%=Request.QueryString(fld)%>"/>
<%
	End If
Next
%>
<input type="submit" name="g2p_dummy" value="Re-submit request as form post">
</form>