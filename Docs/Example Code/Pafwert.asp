<html>
<body>
<h1>Pafwert Password Generator</h1>
<pre>
<%
Dim oPaf, sPass
Set oPaf=CreateObject("Pafwert.PafwertLib")

For i=1 to 5
	'Do 
		sPass = oPaf.GeneratePassword
	'Loop Until Len(sPass)>7 'Enforce a minimum length
	Response.Write Server.HTMLEncode(sPass) & "<br/>"
Next
%>
</pre>
<p><i>Refresh this page for another set of passwords</i></p>
<body>
<html>

