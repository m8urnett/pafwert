Dim oPaf, sPass, iQty
Set oPaf=CreateObject("Pafwert.PafwertLib")

If WScript.Arguments.Count Then 
	iQty=WScript.Arguments(0)
Else
	iQty=5
End If
	
For i=1 to iQty
	Do 
		sPass = oPaf.GeneratePassword
	Loop Until Len(sPass)>7 'Enforce a minimum length
	Wscript.Echo sPass
Next


