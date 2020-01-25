Dim sitos
Dim numeros
Sitos = "https://nhentai.net/g/"
Numeros = InputBox("Enter your magic number","NH Surfer V.1.0 Beta")
Set PWSobj = (CreateObject("Wscript.Shell"))
If numeros = "" then
Wscript.Quit
End If
If IsNumeric(numeros) then
PWSobj.Run "chrome -incognito -url "&Sitos &numeros
End If
Wscript.Quit