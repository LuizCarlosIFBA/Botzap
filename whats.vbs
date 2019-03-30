Dim startmundo
Set Fso = CreateObject("Scripting.FileSystemObject")
     Set InputFile = fso.OpenTextFile("C:\Users\Luiz\Desktop\Botzap\contatos.txt")
'ESTE SCRIPT FOI DESENVOLVIDO PELO STARTMUNDO
'TODOS OS DIREITOS RESERVADOS copartilhe o video ajude mais pessoas
Do While Not (InputFile.atEndOfStream)
     contatos = InputFile.ReadLine
     set startmundo = CreateObject("WScript.Shell")
'startmundo.Run("https://web.whatsapp.com/")
 wscript.sleep 5000
startmundo.SendKeys "{TAB}"
wscript.sleep 600
startmundo.SendKeys"" & contatos
wscript.sleep 3000
startmundo.SendKeys "{ENTER}"
wscript.sleep 2000
startmundo.SendKeys "^(v)"
wscript.sleep 6000  
startmundo.SendKeys "Ola eu sou um Bot feito por Cati, daqui a um tempo sera frequente me comunicar por ele"
wscript.sleep 9000
startmundo.SendKeys "{ENTER}"     
Loop
Wscript.Quit
