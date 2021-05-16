Function generateGUID
'	Reads scriptlet to generate new GUID
	Dim getGUID
	Set getGUID = CreateObject("Scriptlet.TypeLib")
	generateGUID = Left(getGUID.Guid, 38)
End Function

' Dump the generated output to console
wscript.echo generateGUID

